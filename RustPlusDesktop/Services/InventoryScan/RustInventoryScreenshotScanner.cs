using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using OpenCvSharp;
using Tesseract;
using Rect = System.Windows.Rect;

namespace RustPlusDesk.Services.InventoryScan
{
    public sealed class RustInventoryScreenshotScanner
    {
        private const int Columns = 6;
        private const int MatchSize = 72;

        private const double SlotGapRatio = 0.055;

        private const double MinScore = 0.58;
        private const double MinMargin = 0.035;

        public IReadOnlyList<RecognizedInventoryStack> Scan(
            Mat screenMat,
            Rect physicalRect,
            IReadOnlyList<(string ShortName, string DisplayName, Mat TemplateMat)> templates)
        {
            if (screenMat == null || screenMat.Empty())
                throw new ArgumentException("screenMat is empty.", nameof(screenMat));

            if (templates == null || templates.Count == 0)
                throw new ArgumentException("No templates were provided.", nameof(templates));

            using var screenBgr = EnsureBgr(screenMat);

            double pitch = physicalRect.Width / Columns;
            double gap = pitch * SlotGapRatio;
            double slotSize = pitch - gap;

            int rows = Math.Max(1, (int)Math.Floor(physicalRect.Height / pitch + 0.10));

            var results = new List<RecognizedInventoryStack>();

            string debugDir = Path.Combine(AppContext.BaseDirectory, "inventory-scan-debug");
            Directory.CreateDirectory(debugDir);
            ClearOldDebug(debugDir);

            // Estimate Rust slot background from the selected region.
            Vec3b slotBg = EstimateGridBackground(screenBgr, physicalRect);

            var normalizedTemplates = BuildTemplates(templates, slotBg);

            using var tesseract = CreateTesseractEngine();

            try
            {
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < Columns; c++)
                    {
                        int slotIndex = r * Columns + c;

                        int x = (int)Math.Round(physicalRect.X + c * pitch);
                        int y = (int)Math.Round(physicalRect.Y + r * pitch);
                        int w = (int)Math.Round(slotSize);
                        int h = (int)Math.Round(slotSize);

                        if (!IsValidRect(x, y, w, h, screenBgr.Width, screenBgr.Height))
                            continue;

                        using var slot = new Mat(screenBgr, new OpenCvSharp.Rect(x, y, w, h));

                        if (IsProbablyEmptySlot(slot, slotBg))
                        {
                            WriteDebug(debugDir, slotIndex, "EMPTY", 0, 0, slot);
                            continue;
                        }

                        using var preparedSlot = PrepareSlotForMatch(slot, slotBg);

                        MatchResult match = MatchSlot(preparedSlot, normalizedTemplates);

                        WriteDebug(debugDir, slotIndex, match.ShortName, match.Score, match.Margin, preparedSlot);

                        if (!match.Accepted)
                            continue;

                        int quantity = ReadQuantity(slot, tesseract);

                        results.Add(new RecognizedInventoryStack
                        {
                            ShortName = match.ShortName,
                            DisplayName = match.DisplayName,
                            Quantity = quantity,
                            IconConfidence = match.Score,
                            SlotIndex = slotIndex
                        });
                    }
                }
            }
            finally
            {
                foreach (var t in normalizedTemplates)
                    t.Dispose();
            }

            return results;
        }

        private static List<TemplateEntry> BuildTemplates(
            IReadOnlyList<(string ShortName, string DisplayName, Mat TemplateMat)> templates,
            Vec3b slotBg)
        {
            var list = new List<TemplateEntry>();

            foreach (var t in templates)
            {
                if (string.IsNullOrWhiteSpace(t.ShortName))
                    continue;

                if (t.TemplateMat == null || t.TemplateMat.Empty())
                    continue;

                using var composed = ComposeTemplateOnSlotBackground(t.TemplateMat, slotBg);
                using var prepared = PrepareTemplateForMatch(composed, slotBg);

                list.Add(new TemplateEntry
                {
                    ShortName = t.ShortName,
                    DisplayName = string.IsNullOrWhiteSpace(t.DisplayName) ? t.ShortName : t.DisplayName,
                    Bgr = prepared.Clone(),
                    Gray = ToGrayPrepared(prepared),
                    Edge = ToEdgePrepared(prepared),
                    Mean = Cv2.Mean(prepared)
                });
            }

            return list;
        }

        private static Mat ComposeTemplateOnSlotBackground(Mat source, Vec3b slotBg)
        {
            if (source.Channels() == 4)
                return ComposeBgraOnBackground(source, slotBg);

            using var bgr = EnsureBgr(source);

            // If no alpha exists, still place it on a slot-colored canvas.
            var canvas = NewSolidBgr(MatchSize, MatchSize, slotBg);

            using var resized = new Mat();
            Cv2.Resize(bgr, resized, new OpenCvSharp.Size(MatchSize, MatchSize), 0, 0, InterpolationFlags.Area);

            resized.CopyTo(canvas);
            return canvas;
        }

        private static Mat ComposeBgraOnBackground(Mat bgra, Vec3b slotBg)
        {
            int srcW = bgra.Width;
            int srcH = bgra.Height;

            // Find non-transparent bounds.
            Mat[] channels = Cv2.Split(bgra);
            using var alpha = channels[3];

            try
            {
                using var points = new Mat();
                Cv2.FindNonZero(alpha, points);

                if (points.Empty() || points.Rows == 0)
                {
                    return NewSolidBgr(MatchSize, MatchSize, slotBg);
                }

                OpenCvSharp.Rect bounds = Cv2.BoundingRect(points);

                bounds.X = Math.Max(0, bounds.X - 2);
                bounds.Y = Math.Max(0, bounds.Y - 2);
                bounds.Width = Math.Min(srcW - bounds.X, bounds.Width + 4);
                bounds.Height = Math.Min(srcH - bounds.Y, bounds.Height + 4);

                using var cropped = new Mat(bgra, bounds);

                int maxIcon = (int)Math.Round(MatchSize * 0.82);
                double scale = Math.Min(maxIcon / (double)cropped.Width, maxIcon / (double)cropped.Height);

                int newW = Math.Max(1, (int)Math.Round(cropped.Width * scale));
                int newH = Math.Max(1, (int)Math.Round(cropped.Height * scale));

                using var resized = new Mat();
                Cv2.Resize(cropped, resized, new OpenCvSharp.Size(newW, newH), 0, 0, InterpolationFlags.Area);

                var canvas = NewSolidBgra(MatchSize, MatchSize, slotBg);

                int ox = (MatchSize - newW) / 2;
                int oy = (MatchSize - newH) / 2;

                using var roi = new Mat(canvas, new OpenCvSharp.Rect(ox, oy, newW, newH));
                AlphaBlendOnto(resized, roi);

                using var canvasBgr = new Mat();
                Cv2.CvtColor(canvas, canvasBgr, ColorConversionCodes.BGRA2BGR);

                return canvasBgr.Clone();
            }
            finally
            {
                foreach (var ch in channels)
                    ch.Dispose();
            }
        }

        private static void AlphaBlendOnto(Mat srcBgra, Mat dstBgraRoi)
        {
            for (int y = 0; y < srcBgra.Rows; y++)
            {
                for (int x = 0; x < srcBgra.Cols; x++)
                {
                    Vec4b s = srcBgra.At<Vec4b>(y, x);
                    double a = s.Item3 / 255.0;

                    if (a <= 0.001)
                        continue;

                    Vec4b d = dstBgraRoi.At<Vec4b>(y, x);

                    byte b = (byte)Math.Round(s.Item0 * a + d.Item0 * (1.0 - a));
                    byte g = (byte)Math.Round(s.Item1 * a + d.Item1 * (1.0 - a));
                    byte r = (byte)Math.Round(s.Item2 * a + d.Item2 * (1.0 - a));

                    dstBgraRoi.Set(y, x, new Vec4b(b, g, r, 255));
                }
            }
        }

        private static Mat PrepareSlotForMatch(Mat slot, Vec3b slotBg)
        {
            using var bgr = EnsureBgr(slot);

            var cleaned = bgr.Clone();

            int w = cleaned.Width;
            int h = cleaned.Height;

            FillRect(cleaned, new OpenCvSharp.Rect(0, 0, Math.Max(1, (int)(w * 0.055)), h), slotBg); // green condition strip
            FillRect(cleaned, new OpenCvSharp.Rect((int)(w * 0.50), (int)(h * 0.60), w - (int)(w * 0.50), h - (int)(h * 0.60)), slotBg); // count text
            FillRect(cleaned, new OpenCvSharp.Rect(0, 0, w, Math.Max(1, (int)(h * 0.035))), slotBg); // top border
            FillRect(cleaned, new OpenCvSharp.Rect(0, h - Math.Max(1, (int)(h * 0.035)), w, Math.Max(1, (int)(h * 0.035))), slotBg); // bottom border
            FillRect(cleaned, new OpenCvSharp.Rect(0, 0, Math.Max(1, (int)(w * 0.035)), h), slotBg); // left border
            FillRect(cleaned, new OpenCvSharp.Rect(w - Math.Max(1, (int)(w * 0.035)), 0, Math.Max(1, (int)(w * 0.035)), h), slotBg); // right border

            var resized = new Mat();
            Cv2.Resize(cleaned, resized, new OpenCvSharp.Size(MatchSize, MatchSize), 0, 0, InterpolationFlags.Area);
            cleaned.Dispose();

            Cv2.GaussianBlur(resized, resized, new OpenCvSharp.Size(3, 3), 0);
            return resized;
        }

        private static Mat PrepareTemplateForMatch(Mat templateBgr, Vec3b slotBg)
        {
            using var bgr = EnsureBgr(templateBgr);

            var resized = new Mat();
            Cv2.Resize(bgr, resized, new OpenCvSharp.Size(MatchSize, MatchSize), 0, 0, InterpolationFlags.Area);

            Cv2.GaussianBlur(resized, resized, new OpenCvSharp.Size(3, 3), 0);
            return resized;
        }

        private static MatchResult MatchSlot(Mat preparedSlot, IReadOnlyList<TemplateEntry> templates)
        {
            using var slotGray = ToGrayPrepared(preparedSlot);
            using var slotEdge = ToEdgePrepared(preparedSlot);

            Scalar slotMean = Cv2.Mean(preparedSlot);

            double best = double.NegativeInfinity;
            double second = double.NegativeInfinity;

            TemplateEntry? bestTemplate = null;

            foreach (var t in templates)
            {
                double grayScore = TemplateScore(slotGray, t.Gray);
                double edgeScore = TemplateScore(slotEdge, t.Edge);
                double meanScore = MeanColorScore(slotMean, t.Mean);

                double score =
                    grayScore * 0.48 +
                    edgeScore * 0.42 +
                    meanScore * 0.10;

                if (score > best)
                {
                    second = best;
                    best = score;
                    bestTemplate = t;
                }
                else if (score > second)
                {
                    second = score;
                }
            }

            if (bestTemplate == null)
            {
                return new MatchResult
                {
                    Accepted = false,
                    Score = 0,
                    Margin = 0
                };
            }

            if (double.IsNegativeInfinity(second))
                second = 0;

            double margin = best - second;

            bool accepted =
                best >= MinScore &&
                margin >= MinMargin;

            return new MatchResult
            {
                ShortName = bestTemplate.ShortName,
                DisplayName = bestTemplate.DisplayName,
                Score = best,
                Margin = margin,
                Accepted = accepted
            };
        }

        private static double TemplateScore(Mat a, Mat b)
        {
            using var result = new Mat();

            Cv2.MatchTemplate(a, b, result, TemplateMatchModes.CCoeffNormed);
            Cv2.MinMaxLoc(result, out _, out double max, out _, out _);

            if (double.IsNaN(max) || double.IsInfinity(max))
                return -1;

            return max;
        }

        private static double MeanColorScore(Scalar a, Scalar b)
        {
            double diff =
                Math.Abs(a.Val0 - b.Val0) +
                Math.Abs(a.Val1 - b.Val1) +
                Math.Abs(a.Val2 - b.Val2);

            return 1.0 - Math.Min(1.0, diff / 765.0);
        }

        private static Mat ToGrayPrepared(Mat bgr)
        {
            var gray = new Mat();
            Cv2.CvtColor(bgr, gray, ColorConversionCodes.BGR2GRAY);
            Cv2.EqualizeHist(gray, gray);
            return gray;
        }

        private static Mat ToEdgePrepared(Mat bgr)
        {
            using var gray = new Mat();
            Cv2.CvtColor(bgr, gray, ColorConversionCodes.BGR2GRAY);

            var edge = new Mat();
            Cv2.Canny(gray, edge, 40, 120);

            return edge;
        }

        private static bool IsProbablyEmptySlot(Mat slot, Vec3b slotBg)
        {
            using var bgr = EnsureBgr(slot);

            int w = bgr.Width;
            int h = bgr.Height;

            int yMax = Math.Max(1, (int)Math.Round(h * 0.70));
            using var upper = new Mat(bgr, new OpenCvSharp.Rect(0, 0, w, yMax));

            int foreground = 0;
            int total = upper.Width * upper.Height;

            for (int y = 0; y < upper.Rows; y++)
            {
                for (int x = 0; x < upper.Cols; x++)
                {
                    Vec3b p = upper.At<Vec3b>(y, x);

                    double diff =
                        Math.Abs(p.Item0 - slotBg.Item0) +
                        Math.Abs(p.Item1 - slotBg.Item1) +
                        Math.Abs(p.Item2 - slotBg.Item2);

                    if (diff > 52)
                        foreground++;
                }
            }

            double ratio = foreground / Math.Max(1.0, total);

            return ratio < 0.040;
        }

        private static int ReadQuantity(Mat slot, TesseractEngine tesseract)
        {
            int w = slot.Width;
            int h = slot.Height;

            int x = Clamp((int)Math.Round(w * 0.48), 0, w - 1);
            int y = Clamp((int)Math.Round(h * 0.60), 0, h - 1);
            int cw = Clamp((int)Math.Round(w * 0.50), 1, w - x);
            int ch = Clamp((int)Math.Round(h * 0.36), 1, h - y);

            using var count = new Mat(slot, new OpenCvSharp.Rect(x, y, cw, ch));
            using var bgr = EnsureBgr(count);

            using var gray = new Mat();
            Cv2.CvtColor(bgr, gray, ColorConversionCodes.BGR2GRAY);

            using var resized = new Mat();
            Cv2.Resize(gray, resized, new OpenCvSharp.Size(gray.Width * 4, gray.Height * 4), 0, 0, InterpolationFlags.Cubic);

            using var bright = new Mat();
            Cv2.ConvertScaleAbs(resized, bright, 1.55, 25);

            using var thresh = new Mat();
            Cv2.Threshold(bright, thresh, 0, 255, ThresholdTypes.Binary | ThresholdTypes.Otsu);

            using var inverted = new Mat();
            Cv2.BitwiseNot(thresh, inverted);

            try
            {
                Cv2.ImEncode(".png", inverted, out byte[] bytes);
                using var pix = Pix.LoadFromMemory(bytes);
                using var page = tesseract.Process(pix, PageSegMode.SingleWord);

                return ParseRustStackCount(page.GetText());
            }
            catch
            {
                return 1;
            }
        }

        private static void EnsureTrainedDataExists(string trainedDataPath)
        {
            if (File.Exists(trainedDataPath) && new FileInfo(trainedDataPath).Length > 1000000)
                return;

            try
            {
                string dir = Path.GetDirectoryName(trainedDataPath)!;
                Directory.CreateDirectory(dir);

                using var client = new System.Net.Http.HttpClient();
                client.Timeout = TimeSpan.FromSeconds(30);
                var data = client.GetByteArrayAsync("https://github.com/tesseract-ocr/tessdata_fast/raw/main/eng.traineddata").GetAwaiter().GetResult();
                File.WriteAllBytes(trainedDataPath, data);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to download required Tesseract language data: {ex.Message}", ex);
            }
        }

        private static TesseractEngine CreateTesseractEngine()
        {
            string baseDir = AppContext.BaseDirectory
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            string tessDataDir = Path.Combine(baseDir, "tessdata");
            string trainedData = Path.Combine(tessDataDir, "eng.traineddata");

            EnsureTrainedDataExists(trainedData);

            Environment.SetEnvironmentVariable("TESSDATA_PREFIX", tessDataDir);

            var engine = new TesseractEngine(tessDataDir, "eng", EngineMode.LstmOnly);
            engine.SetVariable("tessedit_char_whitelist", "0123456789xX×ftFT ");
            engine.SetVariable("classify_bln_numeric_mode", "1");

            return engine;
        }

        private static int ParseRustStackCount(string rawText)
        {
            if (string.IsNullOrWhiteSpace(rawText))
                return 1;

            string cleaned = rawText
                .ToLowerInvariant()
                .Replace("x", "")
                .Replace("×", "")
                .Replace("ft", "")
                .Replace(" ", "")
                .Replace("\n", "")
                .Replace("\r", "")
                .Replace("\t", "")
                .Trim();

            string digits = new string(cleaned.Where(char.IsDigit).ToArray());

            if (int.TryParse(digits, out int qty) && qty > 0)
                return qty;

            return 1;
        }

        private static Vec3b EstimateGridBackground(Mat screenBgr, Rect rect)
        {
            int x = Clamp((int)Math.Round(rect.X), 0, screenBgr.Width - 1);
            int y = Clamp((int)Math.Round(rect.Y), 0, screenBgr.Height - 1);
            int w = Clamp((int)Math.Round(rect.Width), 1, screenBgr.Width - x);
            int h = Clamp((int)Math.Round(rect.Height), 1, screenBgr.Height - y);

            using var roi = new Mat(screenBgr, new OpenCvSharp.Rect(x, y, w, h));

            var samples = new List<Vec3b>();

            int stepX = Math.Max(1, roi.Width / 20);
            int stepY = Math.Max(1, roi.Height / 20);

            for (int yy = 0; yy < roi.Height; yy += stepY)
            {
                for (int xx = 0; xx < roi.Width; xx += stepX)
                {
                    Vec3b p = roi.At<Vec3b>(yy, xx);

                    // Rust slot background is mid-dark gray, not very bright, not pure black.
                    int sum = p.Item0 + p.Item1 + p.Item2;
                    if (sum > 70 && sum < 260)
                        samples.Add(p);
                }
            }

            if (samples.Count == 0)
                return new Vec3b(52, 52, 49);

            byte b = (byte)samples.Select(p => (int)p.Item0).OrderBy(v => v).ElementAt(samples.Count / 2);
            byte g = (byte)samples.Select(p => (int)p.Item1).OrderBy(v => v).ElementAt(samples.Count / 2);
            byte r = (byte)samples.Select(p => (int)p.Item2).OrderBy(v => v).ElementAt(samples.Count / 2);

            return new Vec3b(b, g, r);
        }

        public static Mat BitmapSourceToMat(BitmapSource bitmapSource)
        {
            if (bitmapSource == null)
                throw new ArgumentNullException(nameof(bitmapSource));

            if (bitmapSource.Format != PixelFormats.Bgra32)
            {
                bitmapSource = new FormatConvertedBitmap(bitmapSource, PixelFormats.Bgra32, null, 0);
            }

            int width = bitmapSource.PixelWidth;
            int height = bitmapSource.PixelHeight;

            var mat = new Mat(height, width, MatType.CV_8UC4);

            int stride = width * 4;
            bitmapSource.CopyPixels(new Int32Rect(0, 0, width, height), mat.Data, stride * height, stride);

            return mat;
        }

        public static Mat EnsureBgr(Mat src)
        {
            if (src == null || src.Empty())
                throw new ArgumentException("Source Mat is empty.", nameof(src));

            var dst = new Mat();

            if (src.Channels() == 4)
                Cv2.CvtColor(src, dst, ColorConversionCodes.BGRA2BGR);
            else if (src.Channels() == 1)
                Cv2.CvtColor(src, dst, ColorConversionCodes.GRAY2BGR);
            else
                src.CopyTo(dst);

            return dst;
        }

        private static Mat NewSolidBgr(int w, int h, Vec3b bg)
        {
            return new Mat(h, w, MatType.CV_8UC3, new Scalar(bg.Item0, bg.Item1, bg.Item2));
        }

        private static Mat NewSolidBgra(int w, int h, Vec3b bg)
        {
            return new Mat(h, w, MatType.CV_8UC4, new Scalar(bg.Item0, bg.Item1, bg.Item2, 255));
        }

        private static void FillRect(Mat mat, OpenCvSharp.Rect rect, Vec3b color)
        {
            rect.X = Clamp(rect.X, 0, mat.Width - 1);
            rect.Y = Clamp(rect.Y, 0, mat.Height - 1);
            rect.Width = Clamp(rect.Width, 1, mat.Width - rect.X);
            rect.Height = Clamp(rect.Height, 1, mat.Height - rect.Y);

            Cv2.Rectangle(mat, rect, new Scalar(color.Item0, color.Item1, color.Item2), -1);
        }

        private static bool IsValidRect(int x, int y, int w, int h, int maxW, int maxH)
        {
            return x >= 0 &&
                   y >= 0 &&
                   w > 0 &&
                   h > 0 &&
                   x + w <= maxW &&
                   y + h <= maxH;
        }

        private static int Clamp(int value, int min, int max)
        {
            if (max < min)
                return min;

            if (value < min)
                return min;

            if (value > max)
                return max;

            return value;
        }

        private static void ClearOldDebug(string debugDir)
        {
            try
            {
                foreach (string file in Directory.GetFiles(debugDir, "slot_*.png"))
                    File.Delete(file);
            }
            catch
            {
                // ignore debug cleanup failure
            }
        }

        private static void WriteDebug(
            string debugDir,
            int slotIndex,
            string name,
            double score,
            double margin,
            Mat mat)
        {
            try
            {
                string safe = string.IsNullOrWhiteSpace(name)
                    ? "unknown"
                    : string.Join("_", name.Split(Path.GetInvalidFileNameChars()));

                string path = Path.Combine(
                    debugDir,
                    $"slot_{slotIndex:00}_{safe}_{score:0.000}_{margin:0.000}.png");

                Cv2.ImWrite(path, mat);
            }
            catch
            {
                // debug should not break scan
            }
        }

        private sealed class TemplateEntry : IDisposable
        {
            public string ShortName { get; set; } = "";
            public string DisplayName { get; set; } = "";
            public Mat Bgr { get; set; } = new();
            public Mat Gray { get; set; } = new();
            public Mat Edge { get; set; } = new();
            public Scalar Mean { get; set; }

            public void Dispose()
            {
                Bgr.Dispose();
                Gray.Dispose();
                Edge.Dispose();
            }
        }

        private sealed class MatchResult
        {
            public string ShortName { get; set; } = "";
            public string DisplayName { get; set; } = "";
            public double Score { get; set; }
            public double Margin { get; set; }
            public bool Accepted { get; set; }
        }
    }
}