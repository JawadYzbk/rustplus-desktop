import { z } from "zod";

/**
 * Typed IPC channel framework.
 *
 * Security posture (ELECTRON_ARCHITECTURE §3): every invoke channel is declared here with a zod request
 * schema and a zod response schema. The main process registers handlers only for declared channels and
 * validates every incoming payload; the preload validates before sending and after receiving. A renderer
 * asking for an undeclared channel is denied by construction (no handler exists) and logged.
 */

export interface ChannelDef<TRequest extends z.ZodTypeAny, TResponse extends z.ZodTypeAny> {
  /** Literal channel name; also the registry key. */
  readonly name: string;
  readonly request: TRequest;
  readonly response: TResponse;
  /** Human-readable description surfaced in diagnostics and the future IPC inspector. */
  readonly description: string;
}

export function defineChannel<TRequest extends z.ZodTypeAny, TResponse extends z.ZodTypeAny>(
  name: string,
  request: TRequest,
  response: TResponse,
  description: string,
): ChannelDef<TRequest, TResponse> {
  return { name, request, response, description };
}

/** Error envelope used for every rejected/failed invoke. */
export const ipcErrorSchema = z.object({
  code: z.enum(["bad_request", "unknown_channel", "handler_error", "unavailable"]),
  message: z.string().max(2000),
});

export type IpcError = z.infer<typeof ipcErrorSchema>;
