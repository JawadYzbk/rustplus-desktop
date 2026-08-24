export class CloudApiError extends Error {
  readonly isConflict: boolean;

  constructor(readonly status: number, readonly path: string, readonly reason: string) {
    super(`cloud ${status} ${path}: ${reason}`);
    this.name = "CloudApiError";
    this.isConflict = status === 409;
  }
}

type QueryValue = string | number | boolean | null | undefined;

export interface CloudRequestOptions {
  json?: unknown;
  query?: Record<string, QueryValue | QueryValue[]>;
}

export interface CloudApiClientOptions {
  baseUrl: string;
  clientVersion: string;
  token?: () => string | null;
  fetchImpl?: typeof fetch;
  onUnauthorized?: () => void;
  timeoutMs?: number;
}

/** Single main-process HTTP choke point for the Laravel API. */
export class CloudApiClient {
  private readonly fetchImpl: typeof fetch;
  private readonly timeoutMs: number;

  constructor(private readonly options: CloudApiClientOptions) {
    this.fetchImpl = options.fetchImpl ?? globalThis.fetch;
    this.timeoutMs = options.timeoutMs ?? 15_000;
  }

  async request<T>(method: string, path: string, options: CloudRequestOptions = {}): Promise<T> {
    const url = new URL(`${this.options.baseUrl.replace(/\/+$/, "")}/v1/${path.replace(/^\/+/, "")}`);
    for (const [key, value] of Object.entries(options.query ?? {})) {
      for (const item of Array.isArray(value) ? value : [value]) {
        if (item !== null && item !== undefined) url.searchParams.append(key, String(item));
      }
    }

    const headers = new Headers({
      Accept: "application/json",
      "X-Client-Version": this.options.clientVersion,
    });
    const token = this.options.token?.();
    if (token) headers.set("Authorization", `Bearer ${token}`);
    if (options.json !== undefined) {
      headers.set("Content-Type", "application/json");
    }

    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), this.timeoutMs);
    let response: Response;
    try {
      response = await this.fetchImpl(url, {
        method,
        headers,
        body: options.json === undefined ? undefined : JSON.stringify(options.json),
        signal: controller.signal,
      });
    } catch (error) {
      throw new CloudApiError(0, path, error instanceof Error ? error.message : String(error));
    } finally {
      clearTimeout(timeout);
    }

    const body = await readBody(response);
    if (!response.ok) {
      if (response.status === 401 || response.status === 403) this.options.onUnauthorized?.();
      throw new CloudApiError(response.status, path, safeReason(body));
    }
    if (body === undefined) return undefined as T;
    return unwrapData(body) as T;
  }
}

async function readBody(response: Response): Promise<unknown> {
  const text = await response.text();
  if (!text) return undefined;
  try {
    return JSON.parse(text) as unknown;
  } catch {
    return text;
  }
}

function unwrapData(body: unknown): unknown {
  if (typeof body === "object" && body !== null && !Array.isArray(body) && "data" in body) {
    return (body as { data: unknown }).data;
  }
  return body;
}

function safeReason(body: unknown): string {
  if (typeof body === "string") return body.slice(0, 200);
  if (typeof body !== "object" || body === null || Array.isArray(body)) return "request failed";
  const record = body as Record<string, unknown>;
  const errors = record["errors"];
  if (typeof errors === "object" && errors !== null) {
    for (const value of Object.values(errors as Record<string, unknown>)) {
      if (Array.isArray(value) && typeof value[0] === "string") return value[0].slice(0, 200);
    }
  }
  return typeof record["message"] === "string" ? record["message"].slice(0, 200) : "request failed";
}
