using System.Net;
using System.Net.Sockets;
using System.Net.WebSockets;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ProtoBuf;
using RustPlusApi;
using RustPlusContracts;
using RustPlusDesk.Infrastructure.Connections;

namespace RustPlusDesk.CrossPlatform.Tests;

[TestClass]
public sealed class TeamTransportTests
{
    [TestMethod]
    public async Task RealSdkSocket_ProjectsUtcWipeAndOfflineTeamMembers_WithoutExposingPairingCredentials()
    {
        using var reservation = new TcpListener(IPAddress.Loopback, 0);
        reservation.Start();
        int port = ((IPEndPoint)reservation.LocalEndpoint).Port;
        reservation.Stop();
        using var listener = new HttpListener();
        listener.Prefixes.Add($"http://127.0.0.1:{port}/");
        listener.Start();
        using var lifetime = new CancellationTokenSource(TimeSpan.FromSeconds(15));
        var serving = ServeAsync(listener, lifetime.Token);
        var transport = new RustPlusTransport(new RustPlusConnection("127.0.0.1", port, 76561198000000001, 123456789));
        try
        {
            await transport.InitializeAsync(lifetime.Token);
            Assert.AreEqual(DateTimeOffset.FromUnixTimeSeconds(1_792_000_000), transport.InitialStatus.WipeTime);
            Assert.AreEqual(1u, transport.InitialStatus.Players);
            var members = await transport.ReadTeamAsync(lifetime.Token);
            Assert.AreEqual(2, members.Count);
            Assert.AreEqual("76561198000000002", members[1].SteamId);
            Assert.AreEqual("Offline teammate", members[1].Name);
            Assert.AreEqual(false, members[1].IsOnline);
            Assert.AreEqual(true, members[1].IsAlive);
            string safe = System.Text.Json.JsonSerializer.Serialize(members);
            Assert.IsFalse(safe.Contains("123456789"));
            Assert.IsFalse(safe.Contains("PlayerToken"));
        }
        finally
        {
            await transport.DisposeAsync();
            lifetime.Cancel();
            listener.Stop();
            try { await serving; }
            catch (Exception ex) when (ex is OperationCanceledException or HttpListenerException or WebSocketException) { }
        }
    }

    private static async Task ServeAsync(HttpListener listener, CancellationToken ct)
    {
        var context = await listener.GetContextAsync().WaitAsync(ct);
        using var socket = (await context.AcceptWebSocketAsync(null)).WebSocket;
        byte[] buffer = new byte[16_384];
        while (!ct.IsCancellationRequested && socket.State == WebSocketState.Open)
        {
            using var requestBytes = new MemoryStream();
            WebSocketReceiveResult received;
            do
            {
                received = await socket.ReceiveAsync(new ArraySegment<byte>(buffer), ct);
                if (received.MessageType == WebSocketMessageType.Close)
                {
                    await socket.CloseOutputAsync(WebSocketCloseStatus.NormalClosure, "fixture complete", ct);
                    return;
                }
                await requestBytes.WriteAsync(buffer.AsMemory(0, received.Count), ct);
                Assert.IsTrue(requestBytes.Length <= buffer.Length);
            } while (!received.EndOfMessage);
            requestBytes.Position = 0;
            var request = Serializer.Deserialize<AppRequest>(requestBytes);
            Assert.AreEqual(76561198000000001ul, request.PlayerId);
            Assert.AreEqual(123456789, request.PlayerToken);
            var response = new AppResponse { Seq = request.Seq };
            if (request.GetInfo is not null)
                response.Info = new AppInfo { Name = "Loopback fixture", WipeTime = 1_792_000_000, Players = 1, MaxPlayers = 100, QueuedPlayers = 0 };
            else if (request.GetTime is not null)
                response.Time = new AppTime { Time = 12, Sunrise = 8, Sunset = 20, DayLengthMinutes = 60, TimeScale = 1 };
            else if (request.GetTeamInfo is not null)
                response.TeamInfo = new AppTeamInfo
                {
                    LeaderSteamId = 76561198000000001,
                    Members =
                    {
                        new AppTeamInfo.Member { SteamId = 76561198000000001, Name = "Self", IsOnline = true, IsAlive = true },
                        new AppTeamInfo.Member { SteamId = 76561198000000002, Name = "Offline teammate", IsOnline = false, IsAlive = true }
                    }
                };
            else response.Success = new AppSuccess();
            using var outgoing = new MemoryStream();
            Serializer.Serialize(outgoing, new AppMessage { Response = response });
            await socket.SendAsync(outgoing.ToArray(), WebSocketMessageType.Binary, true, ct);
        }
    }
}
