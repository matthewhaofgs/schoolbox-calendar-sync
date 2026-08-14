/**
 * Dependency-free Microsoft Graph client for Relay.
 *
 * Authentication uses an Entra application and the OAuth 2.0 client-credentials
 * flow. Client secrets are deliberately kept private and are never included in
 * errors, diagnostics, URLs, or logs by this module.
 */

export const MICROSOFT_GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
export const MICROSOFT_GRAPH_DEFAULT_SCOPE =
  "https://graph.microsoft.com/.default";
export const MICROSOFT_REQUIRED_APPLICATION_PERMISSIONS = [
  "User.Read.All",
  "Calendars.ReadWrite",
] as const;

export const MICROSOFT_USER_SELECT_FIELDS = [
  "id",
  "displayName",
  "givenName",
  "surname",
  "userPrincipalName",
  "mail",
  "proxyAddresses",
  "accountEnabled",
  "userType",
] as const;

export interface MicrosoftEntraCredentials {
  tenantId: string;
  clientId: string;
  clientSecret: string;
}

export interface MicrosoftAccessToken {
  accessToken: string;
  expiresAtMs: number;
  tokenType: string;
}

export interface MicrosoftGraphUser {
  id: string;
  displayName?: string | null;
  givenName?: string | null;
  surname?: string | null;
  userPrincipalName?: string;
  mail?: string | null;
  proxyAddresses?: string[];
  accountEnabled?: boolean | null;
  userType?: string | null;
  [key: string]: unknown;
}

export interface MicrosoftGraphCollection<T> {
  "@odata.context"?: string;
  "@odata.nextLink"?: string;
  value?: T[];
}

export interface ListMicrosoftUsersOptions {
  top?: number;
  select?: readonly string[];
  includeDisabled?: boolean;
  signal?: AbortSignal;
  onPage?: (progress: {
    pageNumber: number;
    pageItems: number;
    accumulatedItems: number;
    hasNextPage: boolean;
  }) => void | Promise<void>;
}

export type MicrosoftCalendarColor =
  | "auto"
  | "lightBlue"
  | "lightGreen"
  | "lightOrange"
  | "lightGray"
  | "lightYellow"
  | "lightTeal"
  | "lightPink"
  | "lightBrown"
  | "lightRed"
  | "maxColor";

export interface MicrosoftCalendar {
  id?: string;
  name?: string;
  color?: MicrosoftCalendarColor | string;
  changeKey?: string;
  canEdit?: boolean;
  canShare?: boolean;
  canViewPrivateItems?: boolean;
  isDefaultCalendar?: boolean;
  owner?: { name?: string; address?: string };
  [key: string]: unknown;
}

export interface MicrosoftCalendarInput {
  name: string;
  color?: MicrosoftCalendarColor;
}

export interface MicrosoftCalendarOptions {
  signal?: AbortSignal;
  ifMatchChangeKey?: string;
}

export interface MicrosoftEventDateTime {
  dateTime: string;
  timeZone: string;
}

export interface MicrosoftEventLocation {
  displayName?: string;
  locationType?: string;
  uniqueId?: string;
  uniqueIdType?: string;
  [key: string]: unknown;
}

export interface MicrosoftEventInput {
  subject?: string;
  body?: {
    contentType: "text" | "html";
    content: string;
  };
  start: MicrosoftEventDateTime;
  end: MicrosoftEventDateTime;
  isAllDay?: boolean;
  location?: MicrosoftEventLocation;
  locations?: MicrosoftEventLocation[];
  showAs?:
    | "free"
    | "tentative"
    | "busy"
    | "oof"
    | "workingElsewhere"
    | "unknown";
  sensitivity?: "normal" | "personal" | "private" | "confidential";
  categories?: string[];
  reminderMinutesBeforeStart?: number;
  isReminderOn?: boolean;
  /** Client-supplied idempotency key supported by Microsoft Graph on create. */
  transactionId?: string;
  [key: string]: unknown;
}

export interface MicrosoftEvent extends MicrosoftEventInput {
  id?: string;
  changeKey?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  webLink?: string;
  isCancelled?: boolean;
}

export interface MicrosoftEventMutationOptions extends MicrosoftCalendarOptions {
  /** Omit to target the user's primary calendar. */
  calendarId?: string;
}

export interface MicrosoftConnectionTestOptions {
  /**
   * Mailbox object ID or user principal name. Required for write verification;
   * a directory sample is used only when verifyWriteAccess is explicitly false.
   */
  targetUserId?: string;
  /** Create and immediately remove a temporary calendar to verify write access. */
  verifyWriteAccess?: boolean;
  signal?: AbortSignal;
}

export interface MicrosoftConnectionTestResult {
  ok: true;
  tenantId: string;
  clientId: string;
  directory: {
    ok: true;
    sampleUsers: number;
    hasMoreUsers: boolean;
  };
  calendar: {
    ok: true;
    targetUserId: string;
    primaryCalendarId?: string;
    secondaryCalendarManagement: boolean;
  };
}

export type MicrosoftFetch = (
  input: RequestInfo | URL,
  init?: RequestInit,
) => Promise<Response>;

export type MicrosoftSleep = (
  delayMs: number,
  signal?: AbortSignal | null,
) => Promise<void>;

export interface MicrosoftGraphClientOptions {
  fetch?: MicrosoftFetch;
  now?: () => number;
  sleep?: MicrosoftSleep;
  /** A cached token is refreshed this many seconds before expiry. */
  tokenExpirySkewSeconds?: number;
  /** Number of retries after the first request for 429 and transient 5xx responses. */
  maxRetries?: number;
  retryBaseDelayMs?: number;
  maxRetryDelayMs?: number;
}

interface MicrosoftOAuthTokenResponse {
  access_token?: unknown;
  expires_in?: unknown;
  token_type?: unknown;
  error?: unknown;
}

interface MicrosoftGraphErrorOptions {
  status: number;
  code?: string;
  requestId?: string;
  date?: string;
  method?: string;
  retryAfterMs?: number;
}

/** A privacy-safe HTTP/OAuth error suitable for Relay diagnostics. */
export class MicrosoftGraphError extends Error {
  readonly status: number;
  readonly code: string;
  readonly requestId?: string;
  readonly date?: string;
  readonly method?: string;
  readonly retryAfterMs?: number;

  constructor(message: string, options: MicrosoftGraphErrorOptions) {
    super(message);
    this.name = "MicrosoftGraphError";
    this.status = options.status;
    this.code = options.code ?? "unknown_error";
    this.requestId = options.requestId;
    this.date = options.date;
    this.method = options.method;
    this.retryAfterMs = options.retryAfterMs;
  }
}

export class MicrosoftConfigurationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "MicrosoftConfigurationError";
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

const GUID_PATTERN =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const EMPTY_GUID = "00000000-0000-0000-0000-000000000000";

function requiredGuid(value: unknown, name: string): string {
  if (typeof value !== "string") {
    throw new MicrosoftConfigurationError(`${name} must be a GUID.`);
  }
  const normalized = value.trim().toLowerCase();
  if (!GUID_PATTERN.test(normalized) || normalized === EMPTY_GUID) {
    throw new MicrosoftConfigurationError(`${name} must be a non-empty GUID.`);
  }
  return normalized;
}

function requiredSecret(value: unknown): string {
  if (
    typeof value !== "string" ||
    value.length === 0 ||
    value.length > 4096 ||
    value.trim() !== value ||
    /[\u0000-\u001f\u007f]/.test(value)
  ) {
    throw new MicrosoftConfigurationError(
      "clientSecret must be a non-empty secret without surrounding whitespace or control characters.",
    );
  }
  return value;
}

/** Parses and strictly validates the three Entra application credentials. */
export function parseMicrosoftCredentials(
  input: MicrosoftEntraCredentials | string | unknown,
): MicrosoftEntraCredentials {
  let value: unknown = input;
  if (typeof input === "string") {
    try {
      value = JSON.parse(input) as unknown;
    } catch {
      throw new MicrosoftConfigurationError(
        "The Microsoft credential is not valid JSON.",
      );
    }
  }
  if (!isRecord(value)) {
    throw new MicrosoftConfigurationError(
      "Microsoft credentials must be an object.",
    );
  }
  return {
    tenantId: requiredGuid(value.tenantId, "tenantId"),
    clientId: requiredGuid(value.clientId, "clientId"),
    clientSecret: requiredSecret(value.clientSecret),
  };
}

function boundedInteger(
  value: number | undefined,
  fallback: number,
  minimum: number,
  maximum: number,
  name: string,
): number {
  const resolved = value ?? fallback;
  if (!Number.isInteger(resolved) || resolved < minimum || resolved > maximum) {
    throw new RangeError(`${name} must be an integer from ${minimum} to ${maximum}.`);
  }
  return resolved;
}

function abortReason(signal?: AbortSignal | null): unknown {
  return signal?.reason ?? new DOMException("Microsoft request was aborted", "AbortError");
}

function throwIfAborted(signal?: AbortSignal | null): void {
  if (signal?.aborted) throw abortReason(signal);
}

function abortable<T>(promise: Promise<T>, signal?: AbortSignal | null): Promise<T> {
  if (!signal) return promise;
  if (signal.aborted) return Promise.reject(abortReason(signal));
  return new Promise<T>((resolve, reject) => {
    const aborted = () => reject(abortReason(signal));
    signal.addEventListener("abort", aborted, { once: true });
    promise.then(
      (value) => {
        signal.removeEventListener("abort", aborted);
        resolve(value);
      },
      (error) => {
        signal.removeEventListener("abort", aborted);
        reject(error);
      },
    );
  });
}

const defaultSleep: MicrosoftSleep = (delayMs, signal) => {
  throwIfAborted(signal);
  return new Promise<void>((resolve, reject) => {
    const timer = setTimeout(() => {
      signal?.removeEventListener("abort", aborted);
      resolve();
    }, delayMs);
    const aborted = () => {
      clearTimeout(timer);
      reject(abortReason(signal));
    };
    signal?.addEventListener("abort", aborted, { once: true });
  });
};

function normalizeResourceId(value: string, name: string): string {
  const normalized = value.trim();
  if (!normalized || normalized.length > 1024 || /[\u0000-\u001f\u007f]/.test(normalized)) {
    throw new TypeError(`${name} is required and must not contain control characters.`);
  }
  return normalized;
}

function normalizeSelect(select?: readonly string[]): string[] {
  const fields = select ?? MICROSOFT_USER_SELECT_FIELDS;
  const normalized = [...new Set(fields.map((field) => field.trim()))];
  if (
    normalized.length === 0 ||
    normalized.some((field) => !/^[A-Za-z][A-Za-z0-9]*$/.test(field))
  ) {
    throw new TypeError("select must contain one or more Microsoft Graph field names.");
  }
  if (!normalized.includes("id")) normalized.unshift("id");
  return normalized;
}

function parseRetryAfterMs(
  response: Response,
  nowMs: number,
): number | undefined {
  const millisecondHeader = response.headers.get("x-ms-retry-after-ms");
  if (millisecondHeader !== null) {
    const milliseconds = Number(millisecondHeader);
    if (Number.isFinite(milliseconds) && milliseconds >= 0) return milliseconds;
  }
  const header = response.headers.get("retry-after");
  if (header === null) return undefined;
  const seconds = Number(header);
  if (Number.isFinite(seconds) && seconds >= 0) return seconds * 1000;
  const dateMs = Date.parse(header);
  return Number.isFinite(dateMs) ? Math.max(0, dateMs - nowMs) : undefined;
}

function isTransientStatus(status: number): boolean {
  return status === 429 || status === 500 || status === 502 || status === 503 || status === 504;
}

async function discardResponse(response: Response): Promise<void> {
  try {
    await response.arrayBuffer();
  } catch {
    // A failed body drain must not hide the response status that caused a retry.
  }
}

async function readResponseBody(response: Response): Promise<unknown> {
  if (response.status === 204 || response.status === 205) return undefined;
  const text = await response.text();
  if (!text) return undefined;
  try {
    return JSON.parse(text) as unknown;
  } catch {
    return undefined;
  }
}

function graphErrorMetadata(body: unknown): {
  code?: string;
  requestId?: string;
  date?: string;
} {
  if (!isRecord(body)) return {};
  const error = body.error;
  if (typeof error === "string") return { code: error };
  if (!isRecord(error)) return {};
  const inner = isRecord(error.innerError)
    ? error.innerError
    : isRecord(error.innererror)
      ? error.innererror
      : undefined;
  const requestId = inner?.["request-id"] ?? inner?.requestId;
  return {
    code: typeof error.code === "string" ? error.code : undefined,
    requestId: typeof requestId === "string" ? requestId : undefined,
    date: typeof inner?.date === "string" ? inner.date : undefined,
  };
}

function throwMicrosoftGraphError(
  response: Response,
  body: unknown,
  method: string,
  nowMs: number,
): never {
  const metadata = graphErrorMetadata(body);
  const code = metadata.code ?? "unknown_error";
  const requestId =
    response.headers.get("request-id") ??
    response.headers.get("client-request-id") ??
    metadata.requestId;
  const date = response.headers.get("date") ?? metadata.date;
  const retryAfterMs = parseRetryAfterMs(response, nowMs);
  throw new MicrosoftGraphError(
    `Microsoft Graph ${method.toUpperCase()} request failed with HTTP ${response.status} (${code}).`,
    {
      status: response.status,
      code,
      requestId: requestId ?? undefined,
      date: date ?? undefined,
      method: method.toUpperCase(),
      retryAfterMs,
    },
  );
}

export function isMicrosoftGraphNotFound(
  error: unknown,
): error is MicrosoftGraphError {
  if (!(error instanceof MicrosoftGraphError)) return false;
  const code = error.code.toLowerCase();
  return (
    error.status === 404 ||
    code === "erroritemnotfound" ||
    code === "request_resourcenotfound" ||
    code === "resourcenotfound"
  );
}

function isMicrosoftGraphForbidden(error: unknown): error is MicrosoftGraphError {
  return error instanceof MicrosoftGraphError && error.status === 403;
}

export function isMicrosoftGraphConflict(
  error: unknown,
): error is MicrosoftGraphError {
  if (!(error instanceof MicrosoftGraphError)) return false;
  const code = error.code.toLowerCase();
  return (
    error.status === 409 ||
    code === "erroritemalreadyexists" ||
    code === "request_conflict" ||
    code === "conflict"
  );
}

function addImmutableIdPreference(headers: Headers): void {
  const existing = headers.get("prefer");
  if (/idtype\s*=\s*"?immutableid"?/i.test(existing ?? "")) return;
  headers.set(
    "prefer",
    existing ? `${existing}, IdType="ImmutableId"` : 'IdType="ImmutableId"',
  );
}

function validateEventInput(event: MicrosoftEventInput): void {
  for (const [name, boundary] of [
    ["start", event.start],
    ["end", event.end],
  ] as const) {
    if (
      !isRecord(boundary) ||
      typeof boundary.dateTime !== "string" ||
      !boundary.dateTime.trim() ||
      typeof boundary.timeZone !== "string" ||
      !boundary.timeZone.trim()
    ) {
      throw new TypeError(`${name} must contain dateTime and timeZone.`);
    }
  }
  if (
    event.transactionId !== undefined &&
    (typeof event.transactionId !== "string" ||
      !event.transactionId.trim() ||
      event.transactionId.length > 255 ||
      /[\u0000-\u001f\u007f]/.test(event.transactionId))
  ) {
    throw new TypeError("transactionId must be a non-empty string of at most 255 characters.");
  }
}

function temporaryCalendarCleanupError(
  calendarId: string,
  error: unknown,
): MicrosoftGraphError {
  const graphError =
    error instanceof MicrosoftGraphError ? error : undefined;
  return new MicrosoftGraphError(
    `Relay created a temporary Microsoft calendar but could not remove it. Delete the calendar with opaque ID ${JSON.stringify(calendarId)} from the selected test mailbox, then run the connection test again.`,
    {
      status: graphError?.status ?? 500,
      code: "temporary_calendar_cleanup_failed",
      requestId: graphError?.requestId,
      date: graphError?.date,
      method: "DELETE",
      retryAfterMs: graphError?.retryAfterMs,
    },
  );
}

export class MicrosoftGraphClient {
  readonly tenantId: string;
  readonly clientId: string;

  private readonly clientSecret: string;
  private readonly fetchFn: MicrosoftFetch;
  private readonly now: () => number;
  private readonly sleep: MicrosoftSleep;
  private readonly tokenExpirySkewMs: number;
  private readonly maxRetries: number;
  private readonly retryBaseDelayMs: number;
  private readonly maxRetryDelayMs: number;
  private cachedToken?: MicrosoftAccessToken;
  private pendingToken?: Promise<MicrosoftAccessToken>;

  constructor(
    credentialsInput: MicrosoftEntraCredentials | string | unknown,
    options: MicrosoftGraphClientOptions = {},
  ) {
    const credentials = parseMicrosoftCredentials(credentialsInput);
    this.tenantId = credentials.tenantId;
    this.clientId = credentials.clientId;
    this.clientSecret = credentials.clientSecret;
    this.fetchFn = options.fetch ?? globalThis.fetch;
    this.now = options.now ?? Date.now;
    this.sleep = options.sleep ?? defaultSleep;
    const skewSeconds = options.tokenExpirySkewSeconds ?? 60;
    if (!Number.isFinite(skewSeconds) || skewSeconds < 0 || skewSeconds >= 3600) {
      throw new RangeError(
        "tokenExpirySkewSeconds must be at least 0 and less than 3600.",
      );
    }
    this.tokenExpirySkewMs = skewSeconds * 1000;
    this.maxRetries = boundedInteger(options.maxRetries, 3, 0, 8, "maxRetries");
    this.retryBaseDelayMs = boundedInteger(
      options.retryBaseDelayMs,
      250,
      1,
      60_000,
      "retryBaseDelayMs",
    );
    this.maxRetryDelayMs = boundedInteger(
      options.maxRetryDelayMs,
      5_000,
      1,
      120_000,
      "maxRetryDelayMs",
    );
    if (this.maxRetryDelayMs < this.retryBaseDelayMs) {
      throw new RangeError("maxRetryDelayMs must be at least retryBaseDelayMs.");
    }
  }

  private tokenUrl(): URL {
    return new URL(
      `https://login.microsoftonline.com/${encodeURIComponent(this.tenantId)}/oauth2/v2.0/token`,
    );
  }

  private async fetchWithRetry(
    input: URL,
    init: RequestInit,
  ): Promise<Response> {
    for (let attempt = 0; ; attempt += 1) {
      throwIfAborted(init.signal);
      const response = await this.fetchFn(input, init);
      if (!isTransientStatus(response.status) || attempt >= this.maxRetries) {
        return response;
      }
      const requestedDelay = parseRetryAfterMs(response, this.now());
      const exponentialDelay = this.retryBaseDelayMs * 2 ** attempt;
      // Graph explicitly asks clients to wait for the complete Retry-After
      // period. The configured cap applies only to Relay's own exponential
      // fallback when the response omits that header.
      const delayMs = requestedDelay ?? Math.min(this.maxRetryDelayMs, exponentialDelay);
      await discardResponse(response);
      await this.sleep(delayMs, init.signal);
    }
  }

  private async exchangeAccessToken(): Promise<MicrosoftAccessToken> {
    const body = new URLSearchParams({
      client_id: this.clientId,
      client_secret: this.clientSecret,
      scope: MICROSOFT_GRAPH_DEFAULT_SCOPE,
      grant_type: "client_credentials",
    });
    const response = await this.fetchWithRetry(this.tokenUrl(), {
      method: "POST",
      headers: {
        accept: "application/json",
        "content-type": "application/x-www-form-urlencoded",
      },
      body: body.toString(),
    });
    const parsed = (await readResponseBody(response)) as MicrosoftOAuthTokenResponse;
    if (!response.ok) {
      throwMicrosoftGraphError(response, parsed, "POST", this.now());
    }
    if (
      !isRecord(parsed) ||
      typeof parsed.access_token !== "string" ||
      !parsed.access_token
    ) {
      throw new MicrosoftGraphError(
        "Microsoft's OAuth token response did not contain an access token.",
        {
          status: response.status,
          code: "invalid_token_response",
          method: "POST",
        },
      );
    }
    const expiresIn =
      typeof parsed.expires_in === "number" && Number.isFinite(parsed.expires_in)
        ? Math.max(1, parsed.expires_in)
        : 3600;
    return {
      accessToken: parsed.access_token,
      expiresAtMs: this.now() + expiresIn * 1000,
      tokenType:
        typeof parsed.token_type === "string" && parsed.token_type
          ? parsed.token_type
          : "Bearer",
    };
  }

  async getAccessTokenInfo(signal?: AbortSignal | null): Promise<MicrosoftAccessToken> {
    throwIfAborted(signal);
    if (
      this.cachedToken &&
      this.cachedToken.expiresAtMs - this.now() > this.tokenExpirySkewMs
    ) {
      return this.cachedToken;
    }
    if (!this.pendingToken) {
      this.pendingToken = this.exchangeAccessToken()
        .then((token) => {
          this.cachedToken = token;
          return token;
        })
        .finally(() => {
          this.pendingToken = undefined;
        });
    }
    return abortable(this.pendingToken, signal);
  }

  async getAccessToken(signal?: AbortSignal | null): Promise<string> {
    return (await this.getAccessTokenInfo(signal)).accessToken;
  }

  clearTokenCache(): void {
    this.cachedToken = undefined;
  }

  private graphUrl(...segments: string[]): URL {
    return new URL(
      `${MICROSOFT_GRAPH_BASE_URL}/${segments
        .map((segment) => encodeURIComponent(segment))
        .join("/")}`,
    );
  }

  private validateNextLink(nextLink: string, expectedPath: RegExp): URL {
    let url: URL;
    try {
      url = new URL(nextLink, `${MICROSOFT_GRAPH_BASE_URL}/`);
    } catch {
      throw new MicrosoftGraphError(
        "Microsoft Graph returned an invalid pagination link.",
        { status: 200, code: "invalid_next_link" },
      );
    }
    const base = new URL(MICROSOFT_GRAPH_BASE_URL);
    if (
      url.protocol !== "https:" ||
      url.origin !== base.origin ||
      !expectedPath.test(url.pathname)
    ) {
      throw new MicrosoftGraphError(
        "Microsoft Graph returned an untrusted pagination link.",
        { status: 200, code: "untrusted_next_link" },
      );
    }
    return url;
  }

  private async authorizedRequest<T>(
    url: URL,
    init: RequestInit = {},
    immutableId = false,
  ): Promise<T> {
    const method = init.method ?? "GET";
    const request = async (forceRefresh: boolean): Promise<Response> => {
      if (forceRefresh) this.clearTokenCache();
      const token = await this.getAccessToken(init.signal);
      const headers = new Headers(init.headers);
      headers.set("authorization", `Bearer ${token}`);
      headers.set("accept", "application/json");
      if (immutableId) addImmutableIdPreference(headers);
      return this.fetchWithRetry(url, { ...init, headers });
    };

    let response = await request(false);
    if (response.status === 401) {
      await discardResponse(response);
      response = await request(true);
    }
    const body = await readResponseBody(response);
    if (!response.ok) {
      throwMicrosoftGraphError(response, body, method, this.now());
    }
    return body as T;
  }

  private usersPageUrl(options: ListMicrosoftUsersOptions): URL {
    const url = this.graphUrl("users");
    url.searchParams.set("$select", normalizeSelect(options.select).join(","));
    url.searchParams.set(
      "$top",
      String(boundedInteger(options.top, 999, 1, 999, "top")),
    );
    return url;
  }

  private async listUsersPage(
    url: URL,
    signal?: AbortSignal,
  ): Promise<MicrosoftGraphCollection<MicrosoftGraphUser>> {
    const page = await this.authorizedRequest<MicrosoftGraphCollection<MicrosoftGraphUser>>(
      url,
      { signal },
    );
    if (!isRecord(page) || !Array.isArray(page.value)) {
      throw new MicrosoftGraphError(
        "Microsoft Graph returned an invalid users collection.",
        { status: 200, code: "invalid_directory_response" },
      );
    }
    for (const user of page.value) {
      if (!isRecord(user) || typeof user.id !== "string" || !user.id) {
        throw new MicrosoftGraphError(
          "Microsoft Graph returned an invalid user entry.",
          { status: 200, code: "invalid_directory_response" },
        );
      }
    }
    return page;
  }

  /** Returns all Entra users, following and validating every Graph next link. */
  async listAllUsers(
    options: ListMicrosoftUsersOptions = {},
  ): Promise<MicrosoftGraphUser[]> {
    const users: MicrosoftGraphUser[] = [];
    const seenLinks = new Set<string>();
    let url: URL | undefined = this.usersPageUrl(options);
    let pageNumber = 0;

    while (url) {
      const page = await this.listUsersPage(url, options.signal);
      pageNumber += 1;
      const pageUsers = page.value ?? [];
      for (const user of pageUsers) {
        if (options.includeDisabled !== false || user.accountEnabled !== false) {
          users.push(user);
        }
      }
      const rawNextLink = page["@odata.nextLink"];
      if (rawNextLink !== undefined && typeof rawNextLink !== "string") {
        throw new MicrosoftGraphError(
          "Microsoft Graph returned an invalid user pagination link.",
          { status: 200, code: "invalid_next_link" },
        );
      }
      const nextLink = rawNextLink;
      await options.onPage?.({
        pageNumber,
        pageItems: pageUsers.length,
        accumulatedItems: users.length,
        hasNextPage: Boolean(nextLink),
      });
      if (!nextLink) {
        url = undefined;
      } else {
        if (seenLinks.has(nextLink)) {
          throw new MicrosoftGraphError(
            "Microsoft Graph returned the same pagination link more than once.",
            { status: 200, code: "repeated_next_link" },
          );
        }
        seenLinks.add(nextLink);
        url = this.validateNextLink(nextLink, /^\/v1\.0\/users\/?$/);
      }
    }
    return users;
  }

  async getPrimaryCalendar(
    userId: string,
    options: MicrosoftCalendarOptions = {},
  ): Promise<MicrosoftCalendar> {
    const url = this.graphUrl("users", normalizeResourceId(userId, "userId"), "calendar");
    url.searchParams.set(
      "$select",
      "id,name,color,changeKey,canEdit,canShare,canViewPrivateItems,isDefaultCalendar",
    );
    return this.authorizedRequest<MicrosoftCalendar>(url, {
      signal: options.signal,
    });
  }

  async listCalendars(
    userId: string,
    options: MicrosoftCalendarOptions = {},
  ): Promise<MicrosoftCalendar[]> {
    const normalizedUserId = normalizeResourceId(userId, "userId");
    const calendars: MicrosoftCalendar[] = [];
    const seenLinks = new Set<string>();
    let url: URL | undefined = this.graphUrl("users", normalizedUserId, "calendars");
    url.searchParams.set(
      "$select",
      "id,name,color,changeKey,canEdit,canShare,canViewPrivateItems,isDefaultCalendar",
    );
    const expectedPath = new RegExp(
      `^/v1\\.0/users/${encodeURIComponent(normalizedUserId).replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}/calendars/?$`,
      "i",
    );
    while (url) {
      const page = await this.authorizedRequest<MicrosoftGraphCollection<MicrosoftCalendar>>(
        url,
        { signal: options.signal },
      );
      if (!isRecord(page) || !Array.isArray(page.value)) {
        throw new MicrosoftGraphError(
          "Microsoft Graph returned an invalid calendars collection.",
          { status: 200, code: "invalid_calendar_response" },
        );
      }
      calendars.push(...page.value);
      const rawNextLink = page["@odata.nextLink"];
      if (rawNextLink !== undefined && typeof rawNextLink !== "string") {
        throw new MicrosoftGraphError(
          "Microsoft Graph returned an invalid calendar pagination link.",
          { status: 200, code: "invalid_next_link" },
        );
      }
      const nextLink = rawNextLink;
      if (!nextLink) {
        url = undefined;
      } else {
        if (seenLinks.has(nextLink)) {
          throw new MicrosoftGraphError(
            "Microsoft Graph returned the same pagination link more than once.",
            { status: 200, code: "repeated_next_link" },
          );
        }
        seenLinks.add(nextLink);
        url = this.validateNextLink(nextLink, expectedPath);
      }
    }
    return calendars;
  }

  /** Creates a secondary calendar in the selected mailbox. */
  async createCalendar(
    userId: string,
    calendar: MicrosoftCalendarInput,
    options: MicrosoftCalendarOptions = {},
  ): Promise<MicrosoftCalendar> {
    if (!calendar.name.trim()) throw new TypeError("A secondary calendar name is required.");
    return this.authorizedRequest<MicrosoftCalendar>(
      this.graphUrl("users", normalizeResourceId(userId, "userId"), "calendars"),
      {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify(calendar),
        signal: options.signal,
      },
    );
  }

  /** Updates administrator-defined properties on a secondary calendar. */
  async updateCalendar(
    userId: string,
    calendarId: string,
    calendar: MicrosoftCalendarInput,
    options: MicrosoftCalendarOptions = {},
  ): Promise<MicrosoftCalendar> {
    if (!calendar.name.trim()) throw new TypeError("A secondary calendar name is required.");
    const headers = new Headers({ "content-type": "application/json" });
    if (options.ifMatchChangeKey) headers.set("if-match", options.ifMatchChangeKey);
    return this.authorizedRequest<MicrosoftCalendar>(
      this.graphUrl(
        "users",
        normalizeResourceId(userId, "userId"),
        "calendars",
        normalizeResourceId(calendarId, "calendarId"),
      ),
      {
        method: "PATCH",
        headers,
        body: JSON.stringify(calendar),
        signal: options.signal,
      },
    );
  }

  /** Permanently deletes a secondary calendar. */
  async deleteCalendar(
    userId: string,
    calendarId: string,
    options: MicrosoftCalendarOptions = {},
  ): Promise<void> {
    const headers = new Headers();
    if (options.ifMatchChangeKey) headers.set("if-match", options.ifMatchChangeKey);
    await this.authorizedRequest<void>(
      this.graphUrl(
        "users",
        normalizeResourceId(userId, "userId"),
        "calendars",
        normalizeResourceId(calendarId, "calendarId"),
      ),
      { method: "DELETE", headers, signal: options.signal },
    );
  }

  private eventCollectionUrl(
    userId: string,
    calendarId?: string,
    eventId?: string,
  ): URL {
    const segments = ["users", normalizeResourceId(userId, "userId")];
    if (calendarId?.trim()) {
      segments.push("calendars", normalizeResourceId(calendarId, "calendarId"));
    } else {
      segments.push("calendar");
    }
    segments.push("events");
    if (eventId !== undefined) {
      segments.push(normalizeResourceId(eventId, "eventId"));
    }
    return this.graphUrl(...segments);
  }

  async insertEvent(
    userId: string,
    event: MicrosoftEventInput,
    options: MicrosoftEventMutationOptions = {},
  ): Promise<MicrosoftEvent> {
    validateEventInput(event);
    return this.authorizedRequest<MicrosoftEvent>(
      this.eventCollectionUrl(userId, options.calendarId),
      {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify(event),
        signal: options.signal,
      },
      true,
    );
  }

  async updateEvent(
    userId: string,
    eventId: string,
    event: MicrosoftEventInput,
    options: MicrosoftEventMutationOptions = {},
  ): Promise<MicrosoftEvent> {
    validateEventInput(event);
    const headers = new Headers({ "content-type": "application/json" });
    if (options.ifMatchChangeKey) headers.set("if-match", options.ifMatchChangeKey);
    const update = { ...event };
    delete update.transactionId;
    return this.authorizedRequest<MicrosoftEvent>(
      this.eventCollectionUrl(userId, options.calendarId, eventId),
      {
        method: "PATCH",
        headers,
        body: JSON.stringify(update),
        signal: options.signal,
      },
      true,
    );
  }

  async deleteEvent(
    userId: string,
    eventId: string,
    options: MicrosoftEventMutationOptions = {},
  ): Promise<void> {
    const headers = new Headers();
    if (options.ifMatchChangeKey) headers.set("if-match", options.ifMatchChangeKey);
    await this.authorizedRequest<void>(
      this.eventCollectionUrl(userId, options.calendarId, eventId),
      { method: "DELETE", headers, signal: options.signal },
      true,
    );
  }

  async testConnection(
    options: MicrosoftConnectionTestOptions = {},
  ): Promise<MicrosoftConnectionTestResult> {
    const verifyWriteAccess = options.verifyWriteAccess !== false;
    const requestedTargetUserId = options.targetUserId?.trim();
    if (verifyWriteAccess && !requestedTargetUserId) {
      throw new MicrosoftConfigurationError(
        "targetUserId is required to verify Microsoft calendar write access.",
      );
    }
    const directoryUrl = this.usersPageUrl({ top: 1, signal: options.signal });
    let page: MicrosoftGraphCollection<MicrosoftGraphUser>;
    try {
      page = await this.listUsersPage(directoryUrl, options.signal);
    } catch (error) {
      if (isMicrosoftGraphForbidden(error)) {
        throw new MicrosoftConfigurationError(
          "Microsoft Graph denied Entra directory discovery. In the Entra app registration, open API permissions and add User.Read.All and Calendars.ReadWrite under Microsoft Graph → Application permissions (not Delegated permissions). Select Grant admin consent for your organisation and confirm both rows show Granted before trying again. Relay cannot add missing API permissions to the app registration.",
        );
      }
      throw error;
    }
    const sampleUsers = page.value ?? [];
    const targetUserId = requestedTargetUserId || sampleUsers[0]?.id;
    if (!targetUserId) {
      throw new MicrosoftConfigurationError(
        "No Entra user is available for the mailbox calendar test.",
      );
    }
    let calendar: MicrosoftCalendar;
    try {
      calendar = await this.getPrimaryCalendar(targetUserId, {
        signal: options.signal,
      });
    } catch (error) {
      if (isMicrosoftGraphForbidden(error)) {
        throw new MicrosoftConfigurationError(
          "Microsoft Graph authenticated but denied calendar access to the test mailbox. Confirm Calendars.ReadWrite is an Application permission with tenant admin consent. If it is already granted, ensure the mailbox is licensed for Exchange Online and is included in any Exchange Application RBAC or Application Access Policy scope.",
        );
      }
      throw error;
    }
    let secondaryCalendarManagement = false;
    if (verifyWriteAccess) {
      let temporaryCalendarId: string | undefined;
      try {
        let created: MicrosoftCalendar;
        try {
          created = await this.createCalendar(targetUserId, {
            name: "Relay connection test - safe to delete",
          }, { signal: options.signal });
        } catch (error) {
          if (isMicrosoftGraphForbidden(error)) {
            throw new MicrosoftConfigurationError(
              "Microsoft Graph can read the test mailbox but cannot create calendars. Confirm Calendars.ReadWrite is an Application permission with tenant admin consent, and ensure any Exchange Application RBAC or Application Access Policy includes this mailbox.",
            );
          }
          throw error;
        }
        if (!created.id?.trim()) {
          throw new MicrosoftGraphError(
            "Microsoft Graph created a connection-test calendar without returning its identifier.",
            { status: 200, code: "invalid_calendar_response" },
          );
        }
        temporaryCalendarId = normalizeResourceId(created.id, "calendarId");
      } finally {
        if (temporaryCalendarId) {
          try {
            // Cleanup must still run if the caller aborts after creation.
            // Keep that independent cleanup bounded so a broken Graph request
            // cannot leave the diagnostic itself running indefinitely.
            await this.deleteCalendar(targetUserId, temporaryCalendarId, {
              signal: AbortSignal.timeout(30_000),
            });
          } catch (error) {
            throw temporaryCalendarCleanupError(temporaryCalendarId, error);
          }
        }
      }
      secondaryCalendarManagement = true;
    }
    return {
      ok: true,
      tenantId: this.tenantId,
      clientId: this.clientId,
      directory: {
        ok: true,
        sampleUsers: sampleUsers.length,
        hasMoreUsers: Boolean(page["@odata.nextLink"]),
      },
      calendar: {
        ok: true,
        targetUserId,
        primaryCalendarId: calendar.id,
        secondaryCalendarManagement,
      },
    };
  }
}

export function createMicrosoftGraphClient(
  credentials: MicrosoftEntraCredentials | string | unknown,
  options?: MicrosoftGraphClientOptions,
): MicrosoftGraphClient {
  return new MicrosoftGraphClient(credentials, options);
}
