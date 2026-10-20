using RustPlusDesk.Core.Cameras;

namespace RustPlusDesk.Infrastructure.Cameras;

/// <summary>
/// Reference GrabCameraSnapshotAsync: a single picture of a camera, for the tile.
///
/// The camera is opened for a moment, a few frames are let through - the picture sharpens over the frames - and the
/// last one is taken; the camera is closed again, so the tiles cost the server one camera at a time.
/// </summary>
public sealed class CameraSnapshotSource(ICameraSessionFactory sessions) : ICameraSnapshotSource
{
    public async Task<byte[]?> GrabAsync(string profilesPath, string profileKey, string cameraId, CancellationToken ct)
    {
        ICameraSession? session = null;
        try
        {
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
            timeout.CancelAfter(TimeSpan.FromSeconds(12));
            session = await sessions.OpenAsync(profilesPath, profileKey, cameraId, timeout.Token).ConfigureAwait(false);
            session.TargetFps = 15;

            byte[]? latest = null;
            int frames = 0;
            var ready = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            void OnFrame(byte[] png)
            {
                latest = png;
                if (++frames >= 40) ready.TrySetResult(true);
            }

            session.FrameRendered += OnFrame;
            try { await ready.Task.WaitAsync(TimeSpan.FromSeconds(8), timeout.Token).ConfigureAwait(false); }
            catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
            catch { /* Whatever accumulated is used. */ }
            finally { session.FrameRendered -= OnFrame; }
            return latest;
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
        catch { return null; }
        finally
        {
            if (session is not null) try { await session.DisposeAsync().ConfigureAwait(false); } catch { /* best effort */ }
        }
    }
}
