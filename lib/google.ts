/**
 * Dependency-free Google Workspace client for the Relay server.
 *
 * Authentication uses a service-account PKCS#8 private key and OAuth 2.0
 * domain-wide delegation. Relay encrypts the service-account JSON before
 * storing it in SQLite; callers must never write credentials to logs.
 */

export const GOOGLE_OAUTH_TOKEN_URL = "https://oauth2.googleapis.com/token";
export const GOOGLE_DIRECTORY_USERS_URL =
  "https://admin.googleapis.com/admin/directory/v1/users";
export const GOOGLE_CALENDAR_API_BASE_URL =
  "https://www.googleapis.com/calendar/v3";

export const GOOGLE_DIRECTORY_USER_READONLY_SCOPE =
  "https://www.googleapis.com/auth/admin.directory.user.readonly";
export const GOOGLE_CALENDAR_EVENTS_OWNED_SCOPE =
  "https://www.googleapis.com/auth/calendar.events.owned";
export const GOOGLE_WORKSPACE_SCOPES = [
  GOOGLE_DIRECTORY_USER_READONLY_SCOPE,
  GOOGLE_CALENDAR_EVENTS_OWNED_SCOPE,
] as const;

export interface GoogleServiceAccountCredentials {
  type: "service_account";
  project_id: string;
  private_key_id: string;
  private_key: string;
  client_email: string;
  client_id: string;
  token_uri: string;
  auth_uri?: string;
  auth_provider_x509_cert_url?: string;
  client_x509_cert_url?: string;
  universe_domain?: string;
}

export interface GoogleAccessToken {
  accessToken: string;
  expiresAtMs: number;
  tokenType: string;
  scopes: readonly string[];
  subject: string;
}

export interface GoogleDirectoryUserName {
  givenName?: string;
  familyName?: string;
  fullName?: string;
  displayName?: string;
  [key: string]: unknown;
}

export interface GoogleDirectoryUser {
  id: string;
  primaryEmail: string;
  name?: GoogleDirectoryUserName;
  suspended?: boolean;
  archived?: boolean;
  isAdmin?: boolean;
  isDelegatedAdmin?: boolean;
  isMailboxSetup?: boolean;
  customerId?: string;
  orgUnitPath?: string;
  aliases?: string[];
  nonEditableAliases?: string[];
  emails?: unknown;
  externalIds?: unknown;
  [key: string]: unknown;
}

export interface GoogleDirectoryUsersPage {
  kind?: string;
  etag?: string;
  users?: GoogleDirectoryUser[];
  nextPageToken?: string;
}

export type GoogleDirectoryProjection = "basic" | "full";

export interface ListGoogleUsersOptions {
  /** `my_customer` includes every domain in the Workspace customer. */
  customer?: string;
  maxResultsPerPage?: number;
  projection?: GoogleDirectoryProjection;
  orderBy?: "email" | "familyName" | "givenName";
  query?: string;
  includeSuspended?: boolean;
  signal?: AbortSignal;
}

export interface GoogleCalendarEventDateTime {
  date?: string;
  dateTime?: string;
  timeZone?: string;
}

export interface GoogleCalendarEventAttendee {
  id?: string;
  email?: string;
  displayName?: string;
  organizer?: boolean;
  self?: boolean;
  resource?: boolean;
  optional?: boolean;
  responseStatus?:
    | "needsAction"
    | "declined"
    | "tentative"
    | "accepted";
  comment?: string;
  additionalGuests?: number;
  [key: string]: unknown;
}

export interface GoogleCalendarEvent {
  kind?: string;
  etag?: string;
  id?: string;
  status?: "confirmed" | "tentative" | "cancelled";
  htmlLink?: string;
  created?: string;
  updated?: string;
  summary?: string;
  description?: string;
  location?: string;
  colorId?: string;
  creator?: Record<string, unknown>;
  organizer?: Record<string, unknown>;
  start?: GoogleCalendarEventDateTime;
  end?: GoogleCalendarEventDateTime;
  endTimeUnspecified?: boolean;
  recurrence?: string[];
  recurringEventId?: string;
  originalStartTime?: GoogleCalendarEventDateTime;
  transparency?: "opaque" | "transparent";
  visibility?: "default" | "public" | "private" | "confidential";
  iCalUID?: string;
  sequence?: number;
  attendees?: GoogleCalendarEventAttendee[];
  attendeesOmitted?: boolean;
  extendedProperties?: {
    private?: Record<string, string>;
    shared?: Record<string, string>;
  };
  reminders?: Record<string, unknown>;
  source?: { url?: string; title?: string };
  eventType?: string;
  [key: string]: unknown;
}

export interface GoogleCalendarEventInput extends GoogleCalendarEvent {
  start: GoogleCalendarEventDateTime;
  end: GoogleCalendarEventDateTime;
}

export interface GoogleCalendarEventsPage {
  kind?: string;
  etag?: string;
  summary?: string;
  description?: string;
  updated?: string;
  timeZone?: string;
  accessRole?: string;
  defaultReminders?: unknown[];
  nextPageToken?: string;
  nextSyncToken?: string;
  items?: GoogleCalendarEvent[];
}

export type GoogleCalendarSendUpdates = "all" | "externalOnly" | "none";

export interface GoogleCalendarMutationOptions {
  calendarId?: string;
  /** Defaults to a stable per-user value and is always sent to Google. */
  quotaUser?: string;
  sendUpdates?: GoogleCalendarSendUpdates;
  conferenceDataVersion?: 0 | 1;
  supportsAttachments?: boolean;
  maxAttendees?: number;
  signal?: AbortSignal;
  /** Adds an `If-Match` header for safe optimistic updates/deletes. */
  ifMatchEtag?: string;
}

export interface ListGoogleCalendarEventsOptions {
  calendarId?: string;
  quotaUser?: string;
  maxResults?: number;
  pageToken?: string;
  syncToken?: string;
  timeMin?: string;
  timeMax?: string;
  updatedMin?: string;
  privateExtendedProperty?: string | string[];
  showDeleted?: boolean;
  singleEvents?: boolean;
  signal?: AbortSignal;
}

export interface GoogleConnectionTestOptions {
  /** A super admin, used as the delegated subject for Directory API access. */
  adminSubject: string;
  /** User whose primary calendar should be checked. Defaults to a listed user. */
  targetUserEmail?: string;
  signal?: AbortSignal;
}

export interface GoogleConnectionTestResult {
  ok: true;
  serviceAccountEmail: string;
  directory: {
    ok: true;
    adminSubject: string;
    sampleUsers: number;
    hasMoreUsers: boolean;
  };
  calendar: {
    ok: true;
    targetUserEmail: string;
    readableEvents: number;
  };
}

export type GoogleFetch = (
  input: RequestInfo | URL,
  init?: RequestInit,
) => Promise<Response>;

export interface GoogleWorkspaceClientOptions {
  fetch?: GoogleFetch;
  crypto?: Pick<Crypto, "subtle">;
  now?: () => number;
  /** A cached token is refreshed this many seconds before its expiry. */
  tokenExpirySkewSeconds?: number;
}

interface GoogleOAuthTokenResponse {
  access_token?: unknown;
  expires_in?: unknown;
  token_type?: unknown;
  scope?: unknown;
  error?: unknown;
  error_description?: unknown;
}

interface GoogleErrorOptions {
  status: number;
  reason?: string;
  details?: unknown;
  url?: string;
  method?: string;
}

/** An HTTP/OAuth error returned by Google, safe to surface in diagnostics. */
export class GoogleApiError extends Error {
  readonly status: number;
  readonly reason: string;
  readonly details?: unknown;
  readonly url?: string;
  readonly method?: string;

  constructor(message: string, options: GoogleErrorOptions) {
    super(message);
    this.name = "GoogleApiError";
    this.status = options.status;
    this.reason = options.reason ?? "unknown_error";
    this.details = options.details;
    this.url = options.url;
    this.method = options.method;
  }
}

export class GoogleConfigurationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "GoogleConfigurationError";
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function requiredString(
  object: Record<string, unknown>,
  field: string,
): string {
  const value = object[field];
  if (typeof value !== "string" || value.trim() === "") {
    throw new GoogleConfigurationError(
      `Service-account JSON is missing a valid ${field} field.`,
    );
  }
  return value;
}

function optionalString(
  object: Record<string, unknown>,
  field: string,
): string | undefined {
  const value = object[field];
  if (value === undefined) return undefined;
  if (typeof value !== "string") {
    throw new GoogleConfigurationError(
      `Service-account JSON field ${field} must be a string.`,
    );
  }
  return value;
}

/** Parses and validates the relevant fields of a downloaded service-account key. */
export function parseServiceAccountJson(
  input: string | unknown,
): GoogleServiceAccountCredentials {
  let value: unknown = input;
  if (typeof input === "string") {
    try {
      value = JSON.parse(input) as unknown;
    } catch {
      throw new GoogleConfigurationError(
        "The Google service-account secret is not valid JSON.",
      );
    }
  }

  if (!isRecord(value)) {
    throw new GoogleConfigurationError(
      "The Google service-account JSON must be an object.",
    );
  }

  const type = requiredString(value, "type");
  if (type !== "service_account") {
    throw new GoogleConfigurationError(
      'The Google credential must have type "service_account".',
    );
  }

  const privateKey = requiredString(value, "private_key").replace(/\\n/g, "\n");
  if (
    !privateKey.includes("-----BEGIN PRIVATE KEY-----") ||
    !privateKey.includes("-----END PRIVATE KEY-----")
  ) {
    throw new GoogleConfigurationError(
      "The service-account private_key must be a PKCS#8 PEM private key.",
    );
  }

  const tokenUri = optionalString(value, "token_uri") ?? GOOGLE_OAUTH_TOKEN_URL;
  let parsedTokenUri: URL;
  try {
    parsedTokenUri = new URL(tokenUri);
  } catch {
    throw new GoogleConfigurationError(
      "The service-account token_uri is not a valid URL.",
    );
  }
  if (parsedTokenUri.protocol !== "https:") {
    throw new GoogleConfigurationError(
      "The service-account token_uri must use HTTPS.",
    );
  }

  return {
    type: "service_account",
    project_id: requiredString(value, "project_id"),
    private_key_id: requiredString(value, "private_key_id"),
    private_key: privateKey,
    client_email: requiredString(value, "client_email"),
    client_id: requiredString(value, "client_id"),
    token_uri: tokenUri,
    auth_uri: optionalString(value, "auth_uri"),
    auth_provider_x509_cert_url: optionalString(
      value,
      "auth_provider_x509_cert_url",
    ),
    client_x509_cert_url: optionalString(value, "client_x509_cert_url"),
    universe_domain: optionalString(value, "universe_domain"),
  };
}

/** Alias retained for callers that refer to the parsed value as credentials. */
export const parseServiceAccountCredentials = parseServiceAccountJson;

function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    const chunk = bytes.subarray(offset, offset + chunkSize);
    for (let index = 0; index < chunk.length; index += 1) {
      binary += String.fromCharCode(chunk[index]);
    }
  }
  return btoa(binary);
}

function base64UrlEncodeBytes(bytes: Uint8Array): string {
  return bytesToBase64(bytes)
    .replace(/=/g, "")
    .replace(/\+/g, "-")
    .replace(/\//g, "_");
}

function base64UrlEncodeText(value: string): string {
  return base64UrlEncodeBytes(new TextEncoder().encode(value));
}

function pemToPkcs8(pem: string): ArrayBuffer {
  const encoded = pem
    .replace("-----BEGIN PRIVATE KEY-----", "")
    .replace("-----END PRIVATE KEY-----", "")
    .replace(/\s/g, "");

  let binary: string;
  try {
    binary = atob(encoded);
  } catch {
    throw new GoogleConfigurationError(
      "The service-account private key contains invalid base64 data.",
    );
  }

  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes.buffer;
}

async function importPrivateKey(
  credentials: GoogleServiceAccountCredentials,
  cryptoProvider: Pick<Crypto, "subtle">,
): Promise<CryptoKey> {
  try {
    return await cryptoProvider.subtle.importKey(
      "pkcs8",
      pemToPkcs8(credentials.private_key),
      { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" },
      false,
      ["sign"],
    );
  } catch (error) {
    if (error instanceof GoogleConfigurationError) throw error;
    throw new GoogleConfigurationError(
      "Google's service-account private key could not be imported as an RS256 PKCS#8 key.",
    );
  }
}

function normalizeSubject(subject: string): string {
  const normalized = subject.trim().toLowerCase();
  if (normalized === "") {
    throw new GoogleConfigurationError(
      "A Google Workspace user must be supplied for domain-wide delegation.",
    );
  }
  return normalized;
}

function normalizeScopes(scopes: readonly string[]): string[] {
  const normalized = [...new Set(scopes.map((scope) => scope.trim()))]
    .filter((scope) => scope !== "")
    .sort();
  if (normalized.length === 0) {
    throw new GoogleConfigurationError("At least one OAuth scope is required.");
  }
  return normalized;
}

async function signJwtAssertion(
  credentials: GoogleServiceAccountCredentials,
  privateKey: CryptoKey,
  subject: string,
  scopes: readonly string[],
  issuedAtMs: number,
  cryptoProvider: Pick<Crypto, "subtle">,
): Promise<string> {
  const issuedAt = Math.floor(issuedAtMs / 1000);
  const header = {
    alg: "RS256",
    typ: "JWT",
    kid: credentials.private_key_id,
  };
  const claims = {
    iss: credentials.client_email,
    sub: normalizeSubject(subject),
    scope: normalizeScopes(scopes).join(" "),
    aud: GOOGLE_OAUTH_TOKEN_URL,
    iat: issuedAt,
    exp: issuedAt + 3600,
  };
  const unsigned = `${base64UrlEncodeText(JSON.stringify(header))}.${base64UrlEncodeText(JSON.stringify(claims))}`;
  const signature = await cryptoProvider.subtle.sign(
    "RSASSA-PKCS1-v1_5",
    privateKey,
    new TextEncoder().encode(unsigned),
  );
  return `${unsigned}.${base64UrlEncodeBytes(new Uint8Array(signature))}`;
}

/** Creates a signed RS256 assertion, primarily useful for setup diagnostics. */
export async function createGoogleJwtAssertion(
  credentialsInput: GoogleServiceAccountCredentials | string | unknown,
  subject: string,
  scopes: readonly string[] = GOOGLE_WORKSPACE_SCOPES,
  nowMs = Date.now(),
  cryptoProvider: Pick<Crypto, "subtle"> = globalThis.crypto,
): Promise<string> {
  const credentials = parseServiceAccountJson(credentialsInput);
  const key = await importPrivateKey(credentials, cryptoProvider);
  return signJwtAssertion(
    credentials,
    key,
    subject,
    scopes,
    nowMs,
    cryptoProvider,
  );
}

function inputBytes(input: string | Uint8Array | ArrayBuffer): ArrayBuffer {
  if (typeof input === "string") {
    return new TextEncoder().encode(input).buffer as ArrayBuffer;
  }
  if (input instanceof Uint8Array) {
    return input.slice().buffer as ArrayBuffer;
  }
  return input;
}

function bytesToHex(bytes: Uint8Array): string {
  let output = "";
  for (const byte of bytes) output += byte.toString(16).padStart(2, "0");
  return output;
}

export async function sha256Hex(
  input: string | Uint8Array | ArrayBuffer,
  cryptoProvider: Pick<Crypto, "subtle"> = globalThis.crypto,
): Promise<string> {
  const digest = await cryptoProvider.subtle.digest("SHA-256", inputBytes(input));
  return bytesToHex(new Uint8Array(digest));
}

/** RFC 4648 base32hex without padding, using Google's required lowercase form. */
export function base32HexEncode(bytes: Uint8Array): string {
  const alphabet = "0123456789abcdefghijklmnopqrstuv";
  let output = "";
  let buffer = 0;
  let bits = 0;

  for (const byte of bytes) {
    buffer = (buffer << 8) | byte;
    bits += 8;
    while (bits >= 5) {
      bits -= 5;
      output += alphabet[(buffer >>> bits) & 31];
    }
    buffer &= (1 << bits) - 1;
  }
  if (bits > 0) output += alphabet[(buffer << (5 - bits)) & 31];
  return output;
}

function stableJsonValue(
  value: unknown,
  seen: Set<object>,
  arrayElement: boolean,
): string | undefined {
  if (value === null) return "null";
  if (typeof value === "string" || typeof value === "boolean") {
    return JSON.stringify(value);
  }
  if (typeof value === "number") {
    return Number.isFinite(value) ? JSON.stringify(value) : "null";
  }
  if (typeof value === "bigint") {
    throw new TypeError("BigInt values cannot be content-hashed as JSON.");
  }
  if (
    typeof value === "undefined" ||
    typeof value === "function" ||
    typeof value === "symbol"
  ) {
    return arrayElement ? "null" : undefined;
  }
  if (typeof value !== "object") return undefined;
  if (seen.has(value)) {
    throw new TypeError("Circular values cannot be content-hashed as JSON.");
  }

  seen.add(value);
  try {
    if (Array.isArray(value)) {
      return `[${value
        .map((item) => stableJsonValue(item, seen, true) ?? "null")
        .join(",")}]`;
    }

    const entries: string[] = [];
    for (const key of Object.keys(value).sort()) {
      const encoded = stableJsonValue(
        (value as Record<string, unknown>)[key],
        seen,
        false,
      );
      if (encoded !== undefined) {
        entries.push(`${JSON.stringify(key)}:${encoded}`);
      }
    }
    return `{${entries.join(",")}}`;
  } finally {
    seen.delete(value);
  }
}

/** Produces canonical JSON with object keys sorted at every depth. */
export function stableJsonStringify(value: unknown): string {
  const encoded = stableJsonValue(value, new Set<object>(), false);
  if (encoded === undefined) {
    throw new TypeError("The value cannot be represented as JSON.");
  }
  return encoded;
}

/** Stable SHA-256 hex for detecting whether a source event's content changed. */
export async function createContentHash(
  value: unknown,
  cryptoProvider: Pick<Crypto, "subtle"> = globalThis.crypto,
): Promise<string> {
  return sha256Hex(stableJsonStringify(value), cryptoProvider);
}

/**
 * Stable Google event ID. SHA-256 hex is a valid lowercase base32hex subset
 * and its 64-character length is within Calendar's 5-1024 character limit.
 */
export async function createDeterministicEventId(
  sourceIdentity: string,
  cryptoProvider: Pick<Crypto, "subtle"> = globalThis.crypto,
): Promise<string> {
  if (sourceIdentity === "") {
    throw new TypeError("A source identity is required for a Google event ID.");
  }
  return sha256Hex(sourceIdentity, cryptoProvider);
}

export const contentHash = createContentHash;
export const deterministicEventId = createDeterministicEventId;

function googleErrorDetails(body: unknown): {
  message?: string;
  reason?: string;
} {
  if (!isRecord(body)) return {};
  const topError = body.error;
  if (typeof topError === "string") {
    return {
      reason: topError,
      message:
        typeof body.error_description === "string"
          ? body.error_description
          : topError,
    };
  }
  if (!isRecord(topError)) return {};

  let reason = typeof topError.status === "string" ? topError.status : undefined;
  const errors = topError.errors;
  if (Array.isArray(errors) && isRecord(errors[0])) {
    if (typeof errors[0].reason === "string") reason = errors[0].reason;
  }
  return {
    reason,
    message: typeof topError.message === "string" ? topError.message : undefined,
  };
}

async function readResponseBody(response: Response): Promise<unknown> {
  const text = await response.text();
  if (text === "") return undefined;
  try {
    return JSON.parse(text) as unknown;
  } catch {
    return text;
  }
}

function throwGoogleApiError(
  response: Response,
  body: unknown,
  method: string,
): never {
  const parsed = googleErrorDetails(body);
  throw new GoogleApiError(
    parsed.message ??
      `Google API request failed with HTTP ${response.status} ${response.statusText}.`,
    {
      status: response.status,
      reason: parsed.reason ?? `http_${response.status}`,
      details: body,
      url: response.url || undefined,
      method,
    },
  );
}

function setOptionalSearchParam(
  url: URL,
  name: string,
  value: string | number | boolean | undefined,
): void {
  if (value !== undefined) url.searchParams.set(name, String(value));
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

export class GoogleWorkspaceClient {
  readonly credentials: GoogleServiceAccountCredentials;

  private readonly fetchFn: GoogleFetch;
  private readonly cryptoProvider: Pick<Crypto, "subtle">;
  private readonly now: () => number;
  private readonly tokenExpirySkewMs: number;
  private readonly tokenCache = new Map<string, GoogleAccessToken>();
  private readonly pendingTokens = new Map<string, Promise<GoogleAccessToken>>();
  private privateKeyPromise?: Promise<CryptoKey>;

  constructor(
    credentialsInput: GoogleServiceAccountCredentials | string | unknown,
    options: GoogleWorkspaceClientOptions = {},
  ) {
    this.credentials = parseServiceAccountJson(credentialsInput);
    this.fetchFn = options.fetch ?? globalThis.fetch;
    this.cryptoProvider = options.crypto ?? globalThis.crypto;
    this.now = options.now ?? Date.now;
    const skewSeconds = options.tokenExpirySkewSeconds ?? 60;
    if (!Number.isFinite(skewSeconds) || skewSeconds < 0 || skewSeconds >= 3600) {
      throw new RangeError(
        "tokenExpirySkewSeconds must be at least 0 and less than 3600.",
      );
    }
    this.tokenExpirySkewMs = skewSeconds * 1000;
  }

  private tokenKey(subject: string, scopes: readonly string[]): string {
    return `${normalizeSubject(subject)}\u0000${normalizeScopes(scopes).join(" ")}`;
  }

  private getPrivateKey(): Promise<CryptoKey> {
    this.privateKeyPromise ??= importPrivateKey(
      this.credentials,
      this.cryptoProvider,
    );
    return this.privateKeyPromise;
  }

  private async exchangeAccessToken(
    subject: string,
    scopes: readonly string[],
  ): Promise<GoogleAccessToken> {
    const normalizedSubject = normalizeSubject(subject);
    const normalizedScopes = normalizeScopes(scopes);
    const assertion = await signJwtAssertion(
      this.credentials,
      await this.getPrivateKey(),
      normalizedSubject,
      normalizedScopes,
      this.now(),
      this.cryptoProvider,
    );
    const form = new URLSearchParams({
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion,
    });
    const response = await this.fetchFn(GOOGLE_OAUTH_TOKEN_URL, {
      method: "POST",
      headers: { "content-type": "application/x-www-form-urlencoded" },
      body: form.toString(),
    });
    const body = (await readResponseBody(response)) as GoogleOAuthTokenResponse;
    if (!response.ok) throwGoogleApiError(response, body, "POST");

    if (
      !isRecord(body) ||
      typeof body.access_token !== "string" ||
      body.access_token === ""
    ) {
      throw new GoogleApiError(
        "Google's OAuth token response did not contain an access token.",
        {
          status: response.status,
          reason: "invalid_token_response",
          details: body,
          url: GOOGLE_OAUTH_TOKEN_URL,
          method: "POST",
        },
      );
    }

    const expiresIn =
      typeof body.expires_in === "number" && Number.isFinite(body.expires_in)
        ? body.expires_in
        : 3600;
    return {
      accessToken: body.access_token,
      expiresAtMs: this.now() + Math.max(1, expiresIn) * 1000,
      tokenType:
        typeof body.token_type === "string" ? body.token_type : "Bearer",
      scopes: normalizedScopes,
      subject: normalizedSubject,
    };
  }

  async getAccessTokenInfo(
    subject: string,
    scopes: readonly string[] = GOOGLE_WORKSPACE_SCOPES,
  ): Promise<GoogleAccessToken> {
    const key = this.tokenKey(subject, scopes);
    const cached = this.tokenCache.get(key);
    if (
      cached &&
      cached.expiresAtMs - this.now() > this.tokenExpirySkewMs
    ) {
      return cached;
    }

    const pending = this.pendingTokens.get(key);
    if (pending) return pending;

    const exchange = this.exchangeAccessToken(subject, scopes)
      .then((token) => {
        this.tokenCache.set(key, token);
        return token;
      })
      .finally(() => {
        this.pendingTokens.delete(key);
      });
    this.pendingTokens.set(key, exchange);
    return exchange;
  }

  async getAccessToken(
    subject: string,
    scopes: readonly string[] = GOOGLE_WORKSPACE_SCOPES,
  ): Promise<string> {
    return (await this.getAccessTokenInfo(subject, scopes)).accessToken;
  }

  clearTokenCache(subject?: string): void {
    if (subject === undefined) {
      this.tokenCache.clear();
      this.pendingTokens.clear();
      return;
    }

    const prefix = `${normalizeSubject(subject)}\u0000`;
    for (const key of this.tokenCache.keys()) {
      if (key.startsWith(prefix)) this.tokenCache.delete(key);
    }
  }

  private async authorizedRequest<T>(
    subject: string,
    scopes: readonly string[],
    url: URL,
    init: RequestInit = {},
  ): Promise<T> {
    const method = init.method ?? "GET";
    const request = async (forceRefresh: boolean): Promise<Response> => {
      if (forceRefresh) {
        this.tokenCache.delete(this.tokenKey(subject, scopes));
      }
      const token = await this.getAccessToken(subject, scopes);
      const headers = new Headers(init.headers);
      headers.set("authorization", `Bearer ${token}`);
      headers.set("accept", "application/json");
      return this.fetchFn(url, { ...init, headers });
    };

    let response = await request(false);
    if (response.status === 401) response = await request(true);
    const body = await readResponseBody(response);
    if (!response.ok) throwGoogleApiError(response, body, method);
    return body as T;
  }

  private async listUsersPage(
    adminSubject: string,
    options: ListGoogleUsersOptions,
    pageToken?: string,
  ): Promise<GoogleDirectoryUsersPage> {
    const url = new URL(GOOGLE_DIRECTORY_USERS_URL);
    url.searchParams.set("customer", options.customer ?? "my_customer");
    url.searchParams.set(
      "maxResults",
      String(
        boundedInteger(
          options.maxResultsPerPage,
          500,
          1,
          500,
          "maxResultsPerPage",
        ),
      ),
    );
    url.searchParams.set("projection", options.projection ?? "basic");
    url.searchParams.set("orderBy", options.orderBy ?? "email");
    setOptionalSearchParam(url, "query", options.query);
    setOptionalSearchParam(url, "pageToken", pageToken);
    return this.authorizedRequest<GoogleDirectoryUsersPage>(
      adminSubject,
      [GOOGLE_DIRECTORY_USER_READONLY_SCOPE],
      url,
      { signal: options.signal },
    );
  }

  /** Returns all users across every domain in the Workspace customer. */
  async listUsers(
    adminSubject: string,
    options: ListGoogleUsersOptions = {},
  ): Promise<GoogleDirectoryUser[]> {
    const users: GoogleDirectoryUser[] = [];
    const seenPageTokens = new Set<string>();
    let pageToken: string | undefined;

    do {
      const page = await this.listUsersPage(adminSubject, options, pageToken);
      if (page.users !== undefined && !Array.isArray(page.users)) {
        throw new GoogleApiError(
          "Google Directory returned an invalid users collection.",
          { status: 200, reason: "invalid_directory_response", details: page },
        );
      }
      for (const user of page.users ?? []) {
        if (options.includeSuspended !== false || !user.suspended) users.push(user);
      }

      pageToken = page.nextPageToken || undefined;
      if (pageToken) {
        if (seenPageTokens.has(pageToken)) {
          throw new GoogleApiError(
            "Google Directory returned the same page token more than once.",
            {
              status: 200,
              reason: "repeated_page_token",
              details: { pageToken },
            },
          );
        }
        seenPageTokens.add(pageToken);
      }
    } while (pageToken);

    return users;
  }

  async listAllUsers(
    adminSubject: string,
    options: ListGoogleUsersOptions = {},
  ): Promise<GoogleDirectoryUser[]> {
    return this.listUsers(adminSubject, options);
  }

  private async quotaUser(
    userEmail: string,
    requested?: string,
  ): Promise<string> {
    const value = requested?.trim() || normalizeSubject(userEmail);
    if (value.length <= 40) return value;
    if (requested !== undefined) {
      throw new RangeError("quotaUser must not exceed 40 characters.");
    }
    return (await sha256Hex(value, this.cryptoProvider)).slice(0, 40);
  }

  private async calendarUrl(
    userEmail: string,
    calendarId: string | undefined,
    suffix: string,
    quotaUser: string | undefined,
  ): Promise<URL> {
    const resolvedCalendarId = calendarId?.trim() || "primary";
    const url = new URL(
      `${GOOGLE_CALENDAR_API_BASE_URL}/calendars/${encodeURIComponent(resolvedCalendarId)}/${suffix}`,
    );
    url.searchParams.set("quotaUser", await this.quotaUser(userEmail, quotaUser));
    return url;
  }

  async listCalendarEvents(
    userEmail: string,
    options: ListGoogleCalendarEventsOptions = {},
  ): Promise<GoogleCalendarEventsPage> {
    const url = await this.calendarUrl(
      userEmail,
      options.calendarId,
      "events",
      options.quotaUser,
    );
    setOptionalSearchParam(
      url,
      "maxResults",
      options.maxResults === undefined
        ? undefined
        : boundedInteger(options.maxResults, 250, 1, 2500, "maxResults"),
    );
    setOptionalSearchParam(url, "pageToken", options.pageToken);
    setOptionalSearchParam(url, "syncToken", options.syncToken);
    setOptionalSearchParam(url, "timeMin", options.timeMin);
    setOptionalSearchParam(url, "timeMax", options.timeMax);
    setOptionalSearchParam(url, "updatedMin", options.updatedMin);
    setOptionalSearchParam(url, "showDeleted", options.showDeleted);
    setOptionalSearchParam(url, "singleEvents", options.singleEvents);
    const extendedProperties = Array.isArray(options.privateExtendedProperty)
      ? options.privateExtendedProperty
      : options.privateExtendedProperty === undefined
        ? []
        : [options.privateExtendedProperty];
    for (const property of extendedProperties) {
      url.searchParams.append("privateExtendedProperty", property);
    }

    return this.authorizedRequest<GoogleCalendarEventsPage>(
      userEmail,
      [GOOGLE_CALENDAR_EVENTS_OWNED_SCOPE],
      url,
      { signal: options.signal },
    );
  }

  async insertEvent(
    userEmail: string,
    event: GoogleCalendarEventInput,
    options: GoogleCalendarMutationOptions = {},
  ): Promise<GoogleCalendarEvent> {
    const url = await this.calendarUrl(
      userEmail,
      options.calendarId,
      "events",
      options.quotaUser,
    );
    setOptionalSearchParam(url, "sendUpdates", options.sendUpdates);
    setOptionalSearchParam(
      url,
      "conferenceDataVersion",
      options.conferenceDataVersion,
    );
    setOptionalSearchParam(url, "supportsAttachments", options.supportsAttachments);
    setOptionalSearchParam(url, "maxAttendees", options.maxAttendees);

    return this.authorizedRequest<GoogleCalendarEvent>(
      userEmail,
      [GOOGLE_CALENDAR_EVENTS_OWNED_SCOPE],
      url,
      {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify(event),
        signal: options.signal,
      },
    );
  }

  async updateEvent(
    userEmail: string,
    eventId: string,
    event: GoogleCalendarEventInput,
    options: GoogleCalendarMutationOptions = {},
  ): Promise<GoogleCalendarEvent> {
    if (eventId.trim() === "") throw new TypeError("eventId is required.");
    const url = await this.calendarUrl(
      userEmail,
      options.calendarId,
      `events/${encodeURIComponent(eventId)}`,
      options.quotaUser,
    );
    setOptionalSearchParam(url, "sendUpdates", options.sendUpdates);
    setOptionalSearchParam(
      url,
      "conferenceDataVersion",
      options.conferenceDataVersion,
    );
    setOptionalSearchParam(url, "supportsAttachments", options.supportsAttachments);
    setOptionalSearchParam(url, "maxAttendees", options.maxAttendees);
    const headers = new Headers({ "content-type": "application/json" });
    if (options.ifMatchEtag) headers.set("if-match", options.ifMatchEtag);

    return this.authorizedRequest<GoogleCalendarEvent>(
      userEmail,
      [GOOGLE_CALENDAR_EVENTS_OWNED_SCOPE],
      url,
      {
        method: "PUT",
        headers,
        body: JSON.stringify(event),
        signal: options.signal,
      },
    );
  }

  async deleteEvent(
    userEmail: string,
    eventId: string,
    options: GoogleCalendarMutationOptions = {},
  ): Promise<void> {
    if (eventId.trim() === "") throw new TypeError("eventId is required.");
    const url = await this.calendarUrl(
      userEmail,
      options.calendarId,
      `events/${encodeURIComponent(eventId)}`,
      options.quotaUser,
    );
    setOptionalSearchParam(url, "sendUpdates", options.sendUpdates);
    const headers = new Headers();
    if (options.ifMatchEtag) headers.set("if-match", options.ifMatchEtag);
    await this.authorizedRequest<void>(
      userEmail,
      [GOOGLE_CALENDAR_EVENTS_OWNED_SCOPE],
      url,
      { method: "DELETE", headers, signal: options.signal },
    );
  }

  async testConnection(
    adminSubjectOrOptions: string | GoogleConnectionTestOptions,
    targetUserEmail?: string,
  ): Promise<GoogleConnectionTestResult> {
    const options: GoogleConnectionTestOptions =
      typeof adminSubjectOrOptions === "string"
        ? { adminSubject: adminSubjectOrOptions, targetUserEmail }
        : adminSubjectOrOptions;
    const page = await this.listUsersPage(
      options.adminSubject,
      { maxResultsPerPage: 1, signal: options.signal },
      undefined,
    );
    const sampleUsers = page.users ?? [];
    const calendarUser =
      options.targetUserEmail ?? sampleUsers[0]?.primaryEmail ?? options.adminSubject;
    const calendar = await this.listCalendarEvents(calendarUser, {
      maxResults: 1,
      showDeleted: false,
      signal: options.signal,
    });

    return {
      ok: true,
      serviceAccountEmail: this.credentials.client_email,
      directory: {
        ok: true,
        adminSubject: normalizeSubject(options.adminSubject),
        sampleUsers: sampleUsers.length,
        hasMoreUsers: Boolean(page.nextPageToken),
      },
      calendar: {
        ok: true,
        targetUserEmail: normalizeSubject(calendarUser),
        readableEvents: calendar.items?.length ?? 0,
      },
    };
  }
}

export function createGoogleWorkspaceClient(
  credentials: GoogleServiceAccountCredentials | string | unknown,
  options?: GoogleWorkspaceClientOptions,
): GoogleWorkspaceClient {
  return new GoogleWorkspaceClient(credentials, options);
}
