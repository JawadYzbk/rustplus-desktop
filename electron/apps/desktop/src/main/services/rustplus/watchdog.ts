/**
 * Chat priming + status watchdog (audit RUSTPLUS_CONNECTIVITY §4, §5).
 *
 * Chat is pull-to-prime: without issuing one GetTeamChat/GetClanChat request per connection no
 * broadcast events arrive. Flags live on the ConnectionCore state so they reset per connection.
 *
 * Watchdog: 5 consecutive failed status polls in the 10 s loop → silent connection refresh.
 */
import { rq, type ProtocolApi } from "./protocol.js";
import type { ConnectionState } from "./connection-core.js";

export class ChatPrimer {
  /** Provider form: the live protocol endpoint changes on every reconnect. */
  constructor(private readonly getProtocol: () => ProtocolApi) {}

  /** Issues one team-chat prime; no-ops when already primed on this connection. */
  async primeTeamChat(state: Pick<ConnectionState, "teamChatPrimed">): Promise<boolean> {
    if (state.teamChatPrimed) return false;
    await this.getProtocol().send(rq.getTeamChat());
    state.teamChatPrimed = true;
    return true;
  }

  async primeClanChat(state: Pick<ConnectionState, "clanChatPrimed">): Promise<boolean> {
    if (state.clanChatPrimed) return false;
    await this.getProtocol().send(rq.getTeamChat()); // clan chat rides getClanChat in newer protos; 2.5.0 uses getTeamChat
    state.clanChatPrimed = true;
    return true;
  }
}

export class StatusWatchdog {
  static readonly FAILURE_LIMIT = 5;

  private consecutive = 0;

  constructor(
    private readonly limit = StatusWatchdog.FAILURE_LIMIT,
    private readonly onSilentRefresh: () => void,
  ) {}

  recordFailure(): void {
    this.consecutive += 1;
    if (this.consecutive >= this.limit) {
      this.consecutive = 0; // refresh and start counting anew
      this.onSilentRefresh();
    }
  }

  recordSuccess(): void {
    this.consecutive = 0;
  }

  get current(): number {
    return this.consecutive;
  }
}
