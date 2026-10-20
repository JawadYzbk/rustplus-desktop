using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RustPlusDesk.Core.Chat;

namespace RustPlusDesk.CrossPlatform.Tests;

/// <summary>
/// The reference app answers chat commands with strings from RustPlusDesktop\Properties\Resources.resx. The port
/// keeps them verbatim (same wording, same punctuation, same dashes), so every ported answer is compared here
/// against the resource it came from. A drifted character (an ASCII hyphen where the reference has an en dash, a
/// doubled space) fails the port even though a normalized comparison would not notice.
/// </summary>
[TestClass]
public sealed class ChatTextParityTests
{
    [TestMethod]
    public void PortedChatTexts_AreVerbatimCopiesOfTheReferenceResources()
    {
        var resources = ReadReferenceResources();
        // Every answer the reference reads from its resources keeps that wording, so a drifted character (an
        // ASCII hyphen where the reference has an en dash, a doubled space) fails the port even though a
        // normalized comparison would not notice.
        foreach (var (key, fallback) in ChatCommandTexts.Fallbacks
            .Where(pair => ChatCommandTexts.Keys.ContainsKey(pair.Key))
            .Select(pair => (Key: ChatCommandTexts.Keys[pair.Key], Fallback: pair.Value)))
        {
            Assert.IsTrue(resources.ContainsKey(key), $"The reference resources do not define {key}.");
            Assert.AreEqual(resources[key], fallback, $"{key} drifted from the reference wording.");
        }
    }

    [TestMethod]
    public void WhoAnswer_IsTheReferenceLiteral()
    {
        var source = ReferenceSource("Views", "MainWindow", "Map", "MainWindow.Map.ChatCommands.cs");
        Assert.IsTrue(source.Contains($"\"{ChatCommandTexts.WhoResponse}\"", StringComparison.Ordinal),
            "The who answer is a literal in the reference chat-command handler.");
    }

    [TestMethod]
    public void OilRigFallbackTexts_AreTheReferenceLiterals()
    {
        var source = ReferenceSource("Views", "MainWindow", "MainWindow.OilRigTimers.cs");
        Assert.IsTrue(source.Contains($"\"{ChatCommandTexts.OilRigNoAlarm}\"", StringComparison.Ordinal),
            "The no-alarm answer is a literal in the reference oil rig handler.");
        Assert.IsTrue(source.Contains("GetString(\"CodeUiNone\") ?? \"None\"", StringComparison.Ordinal),
            "Reference BuildSingleRigAnswer falls back to the literal None.");
        Assert.IsFalse(ReadReferenceResources().ContainsKey("CodeUiNone"),
            "CodeUiNone is not in the reference resources, so the None fallback is what users see.");
        Assert.AreEqual("None", ChatCommandTexts.OilRigNone[(ChatCommandTexts.OilRigNone.IndexOf(':') + 2)..]);
    }

    /// <summary>
    /// The folder can be copied out of the repository, and the reference sources then are not there to compare
    /// against. The check is skipped in that case rather than failed: the port itself does not need the reference.
    /// </summary>
    private static string? RepositoryRoot()
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
        {
            if (File.Exists(Path.Combine(directory.FullName, "RustPlusDesktop", "Properties", "Resources.resx")))
                return directory.FullName;
        }
        return null;
    }

    /// <summary>The reference resources, or a skip when this copy was taken out of the repository.</summary>
    private static Dictionary<string, string> ReadReferenceResources()
    {
        var root = RepositoryRoot();
        if (root is null) Assert.Inconclusive("The reference WPF sources are not beside this copy.");
        var document = XDocument.Load(Path.Combine(root, "RustPlusDesktop", "Properties", "Resources.resx"));
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var data in document.Root!.Elements("data"))
        {
            var name = data.Attribute("name")?.Value;
            var value = data.Element("value")?.Value;
            if (name is not null && value is not null) values[name] = value;
        }
        Assert.IsTrue(values.Count > 1000, "The reference resources were not read.");
        return values;
    }

    /// <summary>The reference source file, or a skip when this copy was taken out of the repository.</summary>
    private static string ReferenceSource(params string[] parts)
    {
        var root = RepositoryRoot();
        if (root is null) Assert.Inconclusive("The reference WPF sources are not beside this copy.");
        return File.ReadAllText(Path.Combine(new[] { root, "RustPlusDesktop" }.Concat(parts).ToArray()));
    }
}
