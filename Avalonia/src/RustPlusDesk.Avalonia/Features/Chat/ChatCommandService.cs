using RustPlusDesk.Core.Chat;

namespace RustPlusDesk.Avalonia.Features.Chat;

/// <summary>
/// Reference ProcessChatCommands wiring: every live chat message is offered to the command engine, answers go to the
/// channel they were asked in, and the profile's response delay applies to each answer. The engine decides; this class
/// only supplies the state (server population and clock, AFK members, event and cargo state) and performs the actions.
/// </summary>
public sealed class ChatCommandService(ChatViewModel chat)
{
    private readonly ChatCommandCooldown cooldown = new();

    /// <summary>The selected profile's chat-command settings; defaults until the profile has been read.</summary>
    public ChatCommandSettings Settings { get; set; } = ChatCommandSettings.Defaults;

    /// <summary>The answer inputs for the current connection (null = nothing connected).</summary>
    public Func<ChatCommandState>? State { get; set; }

    /// <summary>Reference real.PromoteToLeaderAsync; null when the live session cannot hand out leadership.</summary>
    public Func<ulong, CancellationToken, Task<bool>>? PromoteToLeader { get; set; }

    /// <summary>Reference profile.CustomTimers.Add: the timer is stored on the profile before the answer is posted.</summary>
    public Func<ChatCustomTimer, Task>? StartTimer { get; set; }

    public Action<string>? Log { get; set; }
    public Func<DateTime> Now { get; set; } = () => DateTime.UtcNow;

    /// <summary>
    /// Reference TriggerLogicEngineOnChatCommand: every message that is offered to the engine is passed on, so the
    /// logic rules that wait for a command see it whether or not the engine answers it.
    /// </summary>
    public Action<string>? CommandReceived { get; set; }

    /// <summary>Starts answering the live chat feed (history is not offered: only messages that arrive count).</summary>
    public void Attach() => chat.LiveMessage += OnMessage;

    /// <summary>Reference connect reset: the two-second cooldown starts over with the session.</summary>
    public void Reset() => cooldown.Reset();

    private void OnMessage(ChatMessage message) => _ = HandleAsync(message, CancellationToken.None);

    /// <summary>One incoming message: log the decision, perform the action, then answer in the channel asked.</summary>
    public async Task<ChatCommandReply?> HandleAsync(ChatMessage message, CancellationToken ct)
    {
        var request = new ChatCommandRequest(message.Channel, message.Name, message.SteamId, message.Text);
        CommandReceived?.Invoke(message.Text);
        var reply = ChatCommands.Process(request, State?.Invoke() ?? new ChatCommandState(),
            Settings, cooldown, Now());
        if (reply.Log is { } line) Log?.Invoke(line);

        if (reply.PromoteToLeader && PromoteToLeader is { } promote)
            _ = PromoteAsync(promote, message, ct);

        if (reply.NewCustomTimer is { } timer && StartTimer is { } startTimer)
        {
            try { await startTimer(timer).ConfigureAwait(false); }
            catch (Exception) { Log?.Invoke("[ChatCommand] The timer could not be saved."); }
        }

        if (reply.Deferred is { } action)
        {
            ChatCommandReply? outcome = null;
            try { outcome = await action(ct).ConfigureAwait(false); }
            catch (OperationCanceledException) { }
            catch (Exception) { Log?.Invoke("[ChatCommand] The command's device action failed."); }
            if (outcome?.Log is { } actionLine) Log?.Invoke(actionLine);
            reply = outcome ?? ChatCommandReply.Ignored;
        }

        if (reply.Message is { } text)
        {
            try
            {
                await chat.SendCommandReplyAsync(message.Channel, text, TimeSpan.FromSeconds(Settings.ResponseDelaySeconds));
            }
            catch (OperationCanceledException) { }
            catch (Exception) { Log?.Invoke("[ChatCommand] The answer could not be sent to chat."); }
        }

        if (reply.FollowUp is { Count: > 0 } steps)
            _ = SendFollowUpAsync(message.Channel, steps);

        return reply;
    }

    /// <summary>
    /// Reference follow-up answers: the list sends the device commands three seconds after the standard list, and
    /// detailed upkeep spaces tool cupboards at least two seconds apart. Each send still waits the response delay.
    /// </summary>
    private async Task SendFollowUpAsync(ChatChannel channel, IReadOnlyList<ChatMessageStep> steps)
    {
        foreach (var step in steps)
        {
            try
            {
                if (step.DelaySeconds > 0) await Task.Delay(TimeSpan.FromSeconds(step.DelaySeconds));
                await chat.SendCommandReplyAsync(channel, step.Text, TimeSpan.FromSeconds(Settings.ResponseDelaySeconds));
            }
            catch (OperationCanceledException) { return; }
            catch (Exception) { Log?.Invoke("[ChatCommand] The answer could not be sent to chat."); }
        }
    }

    /// <summary>Reference: the promotion request is fired and the answer is posted without waiting for it.</summary>
    private async Task PromoteAsync(Func<ulong, CancellationToken, Task<bool>> promote, ChatMessage message, CancellationToken ct)
    {
        try { await promote(message.SteamId, ct); }
        catch (OperationCanceledException) { }
        catch (Exception) { Log?.Invoke($"[ChatCommand] Promote of {message.Name} failed."); }
    }
}
