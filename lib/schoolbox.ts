/**
 * A small, dependency-free client for the Schoolbox HTTP API.
 *
 * This module deliberately uses only Web Platform APIs and runs in the Relay
 * Node.js server runtime.
 */

export type SchoolboxRoleType = "staff" | "student" | "parent" | "guest";

export interface SchoolboxCampus {
  id: number;
  name?: string;
  code?: string;
  days?: string;
}

export interface SchoolboxRole {
  id: number;
  name?: string;
  type?: SchoolboxRoleType;
}

export interface SchoolboxYearLevel {
  id: number;
  name?: string;
}

/** Fields returned by Schoolbox's GET /api/user endpoint. */
export interface SchoolboxUser {
  id: number;
  externalId?: string | null;
  username?: string;
  email?: string;
  altEmail?: string | null;
  title?: string | null;
  firstName?: string;
  preferredName?: string | null;
  givenName?: string;
  lastName?: string | null;
  fullName?: string | null;
  positionTitle?: string | null;
  superuser?: boolean;
  enabled?: boolean;
  isDeleted?: boolean;
  campus?: SchoolboxCampus[];
  role?: SchoolboxRole;
  yearLevel?: SchoolboxYearLevel | null;
  [field: string]: unknown;
}

export interface SchoolboxListMetadata {
  count?: number;
  cursor?: {
    current: string | null;
    next: string | null;
  };
}

export interface SchoolboxUsersPage {
  data: SchoolboxUser[];
  metadata: SchoolboxListMetadata;
}

export interface RawSchoolboxCalendarEventMeta {
  editable?: boolean;
  eventId?: number | string | null;
  time?: string;
  detail?: string | null;
  location?: string | null;
  eventTypeId?: number;
  eventType?: string;
  variant?: string;
  type?: string | null;
  author?: string;
  authorId?: number;
  level?: string;
  completed?: boolean;
  seriesId?: number | string | null;
  [field: string]: unknown;
}

export interface RawSchoolboxCalendarEvent {
  resourceId?: number | string | null;
  title: string;
  start: string;
  end: string;
  editable: boolean;
  allDay: boolean;
  color?: string;
  className?: string;
  data?: {
    meta?: RawSchoolboxCalendarEventMeta;
    links?: {
      category?: {
        id?: number;
        name?: string;
      };
      path?: string | null;
      [field: string]: unknown;
    };
    styles?: Record<string, unknown>;
    attendance?: Record<string, unknown>;
    classAttendance?: {
      url?: string | null;
      target?: string | null;
      [field: string]: unknown;
    };
    [field: string]: unknown;
  };
  [field: string]: unknown;
}

/** Alternate noun order retained for discoverability. */
export type SchoolboxRawCalendarEvent = RawSchoolboxCalendarEvent;

export interface NormalizedSchoolboxCalendarEvent {
  /** Stable within one Schoolbox instance and suitable for sync bookkeeping. */
  sourceKey: string;
  /** Schoolbox's event identifier, or null when the feed did not provide one. */
  eventId: string | null;
  userId: number;
  title: string;
  /** Plain text. HTML returned in `data.meta.detail` has been removed. */
  description: string;
  location: string | null;
  type: string | null;
  typeCode: string | null;
  typeId: number | null;
  category: EventCategory;
  completed: boolean;
  author: string | null;
  sourceUrl: string | null;
  /** A YYYY-MM-DD value for all-day events, otherwise a date-time string. */
  start: string;
  /** Exclusive YYYY-MM-DD for all-day events, otherwise a date-time string. */
  end: string;
  allDay: boolean;
  editable: boolean;
  raw: RawSchoolboxCalendarEvent;
}

function classifyEventCategory(raw: RawSchoolboxCalendarEvent): EventCategory {
  const meta = raw.data?.meta;
  const classification = [meta?.type, meta?.eventType, meta?.variant, meta?.level, raw.className]
    .filter((value): value is string => typeof value === "string")
    .join(" ")
    .toLocaleLowerCase("en-AU");
  if (/\b(resource|booking)\b/.test(classification)) {
    return "resource_booking";
  }
  if (raw.data?.classAttendance?.url || /\b(timetable|lesson|class)\b/.test(classification)) return "timetable";
  if (/\b(individual|personal)\b/.test(classification)) return "individual_event";
  if (/\bschool[ _-]?event\b/.test(classification)) return "school_event";
  return "other";
}

/** Concise alias for consumers that deal only with normalized events. */
export type SchoolboxCalendarEvent = NormalizedSchoolboxCalendarEvent;

export interface SchoolboxProblemDetails {
  type: string;
  title: string;
  status: number;
  detail?: string;
  instance?: string;
  [extension: string]: unknown;
}

export interface SchoolboxClientOptions {
  baseUrl: string;
  /** A Schoolbox JWT generated for a superuser. */
  jwt: string;
  /** Default window before `now` used by calendar requests. */
  pastDays?: number;
  /** Default window after `now` used by calendar requests. */
  futureDays?: number;
  userPageLimit?: number;
  /** Number of retries after the initial request. */
  maxRetries?: number;
  retryBaseDelayMs?: number;
  retryMaxDelayMs?: number;
  fetchImpl?: typeof fetch;
  /** Injectable for fast deterministic tests. */
  sleep?: (milliseconds: number) => Promise<void>;
}

export interface SchoolboxUsersPageOptions {
  cursor?: string;
  limit?: number;
  signal?: AbortSignal;
}

export interface SchoolboxGetAllUsersOptions {
  pageLimit?: number;
  signal?: AbortSignal;
}

export type SchoolboxDateInput = Date | string | number;

export interface SchoolboxCalendarRange {
  /** Explicit inclusive range start. Overrides `pastDays`. */
  start?: SchoolboxDateInput;
  /** Explicit exclusive range end. Overrides `futureDays`. */
  end?: SchoolboxDateInput;
  pastDays?: number;
  futureDays?: number;
  /** Primarily useful for deterministic jobs and tests. */
  now?: SchoolboxDateInput;
  signal?: AbortSignal;
}

export interface SchoolboxConnectionTestOptions {
  /** Explicit user whose calendar should be used for the delegation test. */
  userId?: number;
  now?: SchoolboxDateInput;
  signal?: AbortSignal;
}

export interface SchoolboxConnectionTestResult {
  ok: true;
  baseUrl: string;
  usersVisible: number;
  usersTotal: number | null;
  calendarUser: Pick<SchoolboxUser, "id" | "fullName" | "email" | "username">;
  calendarEventsFound: number;
  range: {
    start: string;
    end: string;
  };
}

interface SchoolboxRequestOptions {
  signal?: AbortSignal;
}

interface ResolvedCalendarRange {
  start: Date;
  end: Date;
  signal?: AbortSignal;
}

const DAY_MS = 86_400_000;
const DEFAULT_PAST_DAYS = 30;
const DEFAULT_FUTURE_DAYS = 365;
const DEFAULT_USER_PAGE_LIMIT = 100;
const DEFAULT_MAX_RETRIES = 3;
const DEFAULT_RETRY_BASE_DELAY_MS = 250;
const DEFAULT_RETRY_MAX_DELAY_MS = 5_000;
const MAX_PAGINATION_PAGES = 100_000;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function asNonEmptyString(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }

  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

function asFiniteNumber(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

function asBoolean(value: unknown): boolean {
  return value === true || value === 1 || value === "1" || value === "true";
}

function positiveInteger(value: number, name: string): number {
  if (!Number.isInteger(value) || value < 1) {
    throw new RangeError(`${name} must be a positive integer.`);
  }
  return value;
}

function nonNegativeFinite(value: number, name: string): number {
  if (!Number.isFinite(value) || value < 0) {
    throw new RangeError(`${name} must be a non-negative finite number.`);
  }
  return value;
}

function defaultSleep(milliseconds: number): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, milliseconds);
  });
}

function normaliseBaseUrl(baseUrl: string): string {
  const trimmed = baseUrl.trim().replace(/\/+$/, "");
  if (trimmed.length === 0) {
    throw new TypeError("Schoolbox baseUrl is required.");
  }

  let parsed: URL;
  try {
    parsed = new URL(trimmed);
  } catch {
    throw new TypeError("Schoolbox baseUrl must be an absolute HTTP(S) URL.");
  }

  if (parsed.protocol !== "https:") {
    throw new TypeError("Schoolbox baseUrl must use HTTPS.");
  }
  if (parsed.search || parsed.hash) {
    throw new TypeError("Schoolbox baseUrl must not contain a query string or fragment.");
  }

  return trimmed;
}

function normaliseJwt(jwt: string): string {
  const trimmed = jwt.trim().replace(/^Bearer\s+/i, "");
  if (trimmed.length === 0) {
    throw new TypeError("Schoolbox jwt is required.");
  }
  return trimmed;
}

function toDate(value: SchoolboxDateInput, name: string): Date {
  let date: Date;
  if (value instanceof Date) {
    date = new Date(value.getTime());
  } else if (typeof value === "number") {
    // Accommodate both UNIX seconds and JavaScript milliseconds.
    date = new Date(Math.abs(value) < 100_000_000_000 ? value * 1_000 : value);
  } else {
    date = new Date(value);
  }

  if (!Number.isFinite(date.getTime())) {
    throw new RangeError(`${name} must be a valid date.`);
  }
  return date;
}

function toUnixSeconds(date: Date): string {
  return String(Math.floor(date.getTime() / 1_000));
}

function startOfFollowingUtcMonth(date: Date): Date {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + 1, 1));
}

function addUtcDays(dateString: string, days: number): string {
  const date = new Date(`${dateString}T00:00:00.000Z`);
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

function dateOnly(value: string, field: string): string {
  const directDate = /^(\d{4}-\d{2}-\d{2})(?:$|T|\s)/.exec(value.trim());
  if (directDate) {
    return directDate[1];
  }

  const parsed = new Date(value);
  if (!Number.isFinite(parsed.getTime())) {
    throw new TypeError(`Schoolbox event ${field} is not a valid date.`);
  }
  return parsed.toISOString().slice(0, 10);
}

function timedDate(value: string, field: string): string {
  const trimmed = value.trim();
  if (trimmed.length === 0 || !Number.isFinite(Date.parse(trimmed))) {
    throw new TypeError(`Schoolbox event ${field} is not a valid date-time.`);
  }
  return /^\d{4}-\d{2}-\d{2}$/.test(trimmed)
    ? `${trimmed}T00:00:00`
    : trimmed;
}

function shiftTimedDate(value: string, minutes: number): string {
  const timestamp = Date.parse(value);
  if (!Number.isFinite(timestamp)) throw new TypeError("Schoolbox event date-time cannot be repaired.");

  const shifted = new Date(timestamp + minutes * 60_000);
  const offset = value.match(/(Z|[+-]\d{2}:\d{2})$/i)?.[1];
  if (!offset) {
    // A date-time without an offset is interpreted later using the configured
    // Google calendar time zone, so retain the same wall-clock representation.
    const wallClock = value.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?/);
    if (!wallClock) return shifted.toISOString();
    const wallClockTimestamp = Date.UTC(
      Number(wallClock[1]),
      Number(wallClock[2]) - 1,
      Number(wallClock[3]),
      Number(wallClock[4]),
      Number(wallClock[5]),
      Number(wallClock[6] ?? 0),
    ) + minutes * 60_000;
    return new Date(wallClockTimestamp).toISOString().slice(0, 19);
  }
  if (offset.toUpperCase() === "Z") return shifted.toISOString();

  const sign = offset.startsWith("-") ? -1 : 1;
  const offsetMinutes = sign * (Number(offset.slice(1, 3)) * 60 + Number(offset.slice(4, 6)));
  const localTime = new Date(shifted.getTime() + offsetMinutes * 60_000).toISOString().slice(0, 19);
  return `${localTime}${offset}`;
}

function decodeHtmlEntities(input: string): string {
  const namedEntities: Record<string, string> = {
    amp: "&",
    apos: "'",
    gt: ">",
    hellip: "…",
    laquo: "«",
    ldquo: "“",
    lsquo: "‘",
    lt: "<",
    mdash: "—",
    nbsp: " ",
    ndash: "–",
    quot: '"',
    raquo: "»",
    rdquo: "”",
    rsquo: "’",
  };

  return input.replace(/&(#(?:x[0-9a-f]+|\d+)|[a-z][a-z0-9]+);/gi, (entity, code: string) => {
    if (code.startsWith("#")) {
      const hexadecimal = code[1]?.toLowerCase() === "x";
      const numericPart = hexadecimal ? code.slice(2) : code.slice(1);
      const point = Number.parseInt(numericPart, hexadecimal ? 16 : 10);
      if (Number.isFinite(point) && point >= 0 && point <= 0x10ffff) {
        try {
          return String.fromCodePoint(point);
        } catch {
          return entity;
        }
      }
      return entity;
    }

    return namedEntities[code.toLowerCase()] ?? entity;
  });
}

/** Converts Schoolbox event detail HTML into compact, readable plain text. */
export function schoolboxHtmlToText(html: string | null | undefined): string {
  if (!html) {
    return "";
  }

  const withLineBreaks = html
    .replace(/<\s*(?:script|style)\b[^>]*>[\s\S]*?<\s*\/\s*(?:script|style)\s*>/gi, "")
    .replace(/<\s*br\s*\/?\s*>/gi, "\n")
    .replace(/<\s*\/\s*(?:p|div|li|h[1-6]|tr)\s*>/gi, "\n")
    .replace(/<\s*li\b[^>]*>/gi, "• ")
    .replace(/<[^>]+>/g, "");

  return decodeHtmlEntities(withLineBreaks)
    .replace(/\r\n?/g, "\n")
    .replace(/[\t\f\v ]+/g, " ")
    .replace(/ *\n */g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function stableHash(input: string): string {
  // Two differently-seeded FNV-1a passes give a compact deterministic 64-bit-ish
  // identifier without relying on Node crypto or an asynchronous Web Crypto call.
  let high = 0x811c9dc5;
  let low = 0x9e3779b9;
  for (let index = 0; index < input.length; index += 1) {
    const code = input.charCodeAt(index);
    high = Math.imul(high ^ code, 0x01000193);
    low = Math.imul(low ^ code, 0x85ebca6b);
  }
  return `${(high >>> 0).toString(16).padStart(8, "0")}${(low >>> 0)
    .toString(16)
    .padStart(8, "0")}`;
}

function makeAbsoluteUrl(baseUrl: string | undefined, link: string | null | undefined): string | null {
  const value = asNonEmptyString(link);
  if (!value) {
    return null;
  }

  try {
    if (baseUrl) {
      return new URL(value, `${baseUrl}/`).toString();
    }
    return new URL(value).toString();
  } catch {
    // Preserve a relative source link if no base URL was available to resolve it.
    return value;
  }
}

/** Normalize a single item returned by /calendar/ajax/full. */
export function normalizeSchoolboxCalendarEvent(
  raw: RawSchoolboxCalendarEvent,
  userId: number,
  baseUrl?: string,
): NormalizedSchoolboxCalendarEvent {
  positiveInteger(userId, "userId");

  const meta = raw.data?.meta;
  const eventIdValue = meta?.eventId;
  const eventId =
    (typeof eventIdValue === "string" && eventIdValue.trim().length > 0) ||
    (typeof eventIdValue === "number" && Number.isFinite(eventIdValue))
      ? String(eventIdValue).trim()
      : null;
  const title = asNonEmptyString(raw.title) ?? "(Untitled Schoolbox event)";
  const description = schoolboxHtmlToText(meta?.detail);
  const location = asNonEmptyString(meta?.location) ?? null;
  const eventType =
    asNonEmptyString(meta?.type) ??
    asNonEmptyString(meta?.eventType) ??
    asNonEmptyString(meta?.variant) ??
    asNonEmptyString(raw.className) ??
    null;
  const typeCode = asNonEmptyString(meta?.eventType) ?? null;
  const typeId = typeof meta?.eventTypeId === "number" && Number.isFinite(meta.eventTypeId)
    ? meta.eventTypeId
    : null;
  const sourceUrl = makeAbsoluteUrl(
    baseUrl,
    raw.data?.links?.path ?? raw.data?.classAttendance?.url,
  );

  let start: string;
  let end: string;
  if (raw.allDay) {
    const rawStart = raw.start.trim();
    const rawEnd = raw.end.trim();
    if (!rawStart && !rawEnd) {
      throw new TypeError("Schoolbox event has neither a start nor an end date.");
    }
    if (rawStart) {
      start = dateOnly(rawStart, "start");
      end = rawEnd ? dateOnly(rawEnd, "end") : addUtcDays(start, 1);
    } else {
      end = dateOnly(rawEnd, "end");
      start = addUtcDays(end, -1);
    }
    // FullCalendar and Google Calendar both use an exclusive all-day end date.
    if (end <= start) {
      end = addUtcDays(start, 1);
    }
  } else {
    const rawStart = raw.start.trim();
    const rawEnd = raw.end.trim();
    if (!rawStart && !rawEnd) {
      throw new TypeError("Schoolbox event has neither a start nor an end date-time.");
    }
    if (rawStart) {
      start = timedDate(rawStart, "start");
      end = rawEnd ? timedDate(rawEnd, "end") : shiftTimedDate(start, 30);
    } else {
      end = timedDate(rawEnd, "end");
      start = shiftTimedDate(end, -30);
    }
  }

  const fallbackIdentity = JSON.stringify([
    userId,
    title,
    start,
    end,
    raw.allDay,
    location,
    eventType,
    sourceUrl,
    raw.resourceId ?? null,
  ]);

  return {
    sourceKey: eventId
      ? `schoolbox:event:${eventId}`
      : `schoolbox:fallback:${stableHash(fallbackIdentity)}`,
    eventId,
    userId,
    title,
    description,
    location,
    type: eventType,
    typeCode,
    typeId,
    category: classifyEventCategory(raw),
    completed: Boolean(meta?.completed),
    author: asNonEmptyString(meta?.author) ?? null,
    sourceUrl,
    start,
    end,
    allDay: Boolean(raw.allDay),
    editable: Boolean(raw.editable),
    raw,
  };
}

/**
 * Error raised for unsuccessful Schoolbox responses.
 *
 * RFC 7807 fields are exposed directly, while `problem` retains extension
 * members supplied by a particular Schoolbox installation.
 */
export class SchoolboxApiError extends Error {
  readonly status: number;
  readonly type: string;
  readonly title: string;
  readonly detail?: string;
  readonly instance?: string;
  readonly method: string;
  readonly url: string;
  readonly problem: SchoolboxProblemDetails;
  readonly responseBody: unknown;
  readonly requestId: string | null;
  readonly retryAfterMs: number | null;

  constructor(options: {
    method: string;
    url: string;
    status: number;
    type?: string;
    title?: string;
    detail?: string;
    instance?: string;
    problem?: Record<string, unknown>;
    responseBody?: unknown;
    requestId?: string | null;
    retryAfterMs?: number | null;
  }) {
    const title = options.title ?? "Schoolbox API request failed";
    const detail = options.detail?.trim();
    super(`${title} (${options.status})${detail ? `: ${detail}` : ""}`);
    this.name = "SchoolboxApiError";
    this.status = options.status;
    this.type = options.type ?? "about:blank";
    this.title = title;
    this.detail = detail;
    this.instance = options.instance;
    this.method = options.method;
    this.url = options.url;
    this.responseBody = options.responseBody;
    this.requestId = options.requestId ?? null;
    this.retryAfterMs = options.retryAfterMs ?? null;
    this.problem = {
      ...(options.problem ?? {}),
      type: this.type,
      title: this.title,
      status: this.status,
      ...(this.detail ? { detail: this.detail } : {}),
      ...(this.instance ? { instance: this.instance } : {}),
    };

    // Required by a few older JavaScript runtimes when extending built-ins.
    Object.setPrototypeOf(this, new.target.prototype);
  }

  get retryable(): boolean {
    return this.status === 429 || (this.status >= 500 && this.status <= 599);
  }
}

function parseRetryAfter(value: string | null, now = Date.now()): number | null {
  if (!value) {
    return null;
  }

  const seconds = Number(value);
  if (Number.isFinite(seconds) && seconds >= 0) {
    return Math.round(seconds * 1_000);
  }

  const date = Date.parse(value);
  return Number.isFinite(date) ? Math.max(0, date - now) : null;
}

async function responseToApiError(
  response: Response,
  method: string,
  url: string,
): Promise<SchoolboxApiError> {
  let body: unknown;
  let rawText = "";
  try {
    rawText = await response.text();
    if (rawText.trim().length > 0) {
      try {
        body = JSON.parse(rawText) as unknown;
      } catch {
        body = rawText;
      }
    }
  } catch {
    body = undefined;
  }

  const problem = isRecord(body) ? body : undefined;
  const problemStatus = asFiniteNumber(problem?.status);
  const status = response.status || problemStatus || 500;
  const fallbackTitle = response.statusText
    ? `Schoolbox API returned ${response.statusText}`
    : "Schoolbox API request failed";

  return new SchoolboxApiError({
    method,
    url,
    status,
    type: asNonEmptyString(problem?.type),
    title: asNonEmptyString(problem?.title) ?? fallbackTitle,
    detail:
      asNonEmptyString(problem?.detail) ??
      (typeof body === "string" ? body.slice(0, 2_000) : undefined),
    instance: asNonEmptyString(problem?.instance),
    problem,
    responseBody: body,
    requestId:
      response.headers.get("x-request-id") ??
      response.headers.get("x-correlation-id") ??
      response.headers.get("cf-ray"),
    retryAfterMs: parseRetryAfter(response.headers.get("retry-after")),
  });
}

function validateRawCalendarEvent(value: unknown, index: number): RawSchoolboxCalendarEvent {
  if (!isRecord(value)) {
    throw new TypeError(`Schoolbox calendar event at index ${index} is not an object.`);
  }

  if ((value.start != null && typeof value.start !== "string") || (value.end != null && typeof value.end !== "string")) {
    throw new TypeError(`Schoolbox calendar event at index ${index} has an invalid start/end type.`);
  }

  return {
    ...value,
    title: typeof value.title === "string" ? value.title : "",
    start: typeof value.start === "string" ? value.start : "",
    end: typeof value.end === "string" ? value.end : "",
    editable: asBoolean(value.editable),
    allDay: asBoolean(value.allDay),
  } as RawSchoolboxCalendarEvent;
}

function validateUser(value: unknown, index: number): SchoolboxUser {
  if (!isRecord(value) || !Number.isInteger(value.id) || Number(value.id) < 1) {
    throw new TypeError(`Schoolbox user at index ${index} has no valid id.`);
  }
  return value as SchoolboxUser;
}

export class SchoolboxClient {
  readonly baseUrl: string;

  private readonly jwt: string;
  private readonly pastDays: number;
  private readonly futureDays: number;
  private readonly userPageLimit: number;
  private readonly maxRetries: number;
  private readonly retryBaseDelayMs: number;
  private readonly retryMaxDelayMs: number;
  private readonly fetchImpl: typeof fetch;
  private readonly sleep: (milliseconds: number) => Promise<void>;

  constructor(options: SchoolboxClientOptions) {
    this.baseUrl = normaliseBaseUrl(options.baseUrl);
    this.jwt = normaliseJwt(options.jwt);
    this.pastDays = nonNegativeFinite(options.pastDays ?? DEFAULT_PAST_DAYS, "pastDays");
    this.futureDays = nonNegativeFinite(options.futureDays ?? DEFAULT_FUTURE_DAYS, "futureDays");
    this.userPageLimit = positiveInteger(
      options.userPageLimit ?? DEFAULT_USER_PAGE_LIMIT,
      "userPageLimit",
    );
    this.maxRetries = nonNegativeFinite(
      options.maxRetries ?? DEFAULT_MAX_RETRIES,
      "maxRetries",
    );
    if (!Number.isInteger(this.maxRetries)) {
      throw new RangeError("maxRetries must be an integer.");
    }
    this.retryBaseDelayMs = nonNegativeFinite(
      options.retryBaseDelayMs ?? DEFAULT_RETRY_BASE_DELAY_MS,
      "retryBaseDelayMs",
    );
    this.retryMaxDelayMs = nonNegativeFinite(
      options.retryMaxDelayMs ?? DEFAULT_RETRY_MAX_DELAY_MS,
      "retryMaxDelayMs",
    );
    if (this.retryMaxDelayMs < this.retryBaseDelayMs) {
      throw new RangeError("retryMaxDelayMs must be greater than or equal to retryBaseDelayMs.");
    }

    const implementation = options.fetchImpl ?? globalThis.fetch;
    if (typeof implementation !== "function") {
      throw new TypeError("A standards-compatible fetch implementation is required.");
    }
    this.fetchImpl = implementation.bind(globalThis);
    this.sleep = options.sleep ?? defaultSleep;
  }

  async getUsersPage(options: SchoolboxUsersPageOptions = {}): Promise<SchoolboxUsersPage> {
    const limit = positiveInteger(options.limit ?? this.userPageLimit, "limit");
    const query = new URLSearchParams({ limit: String(limit) });
    if (options.cursor) {
      query.set("cursor", options.cursor);
    }

    const value = await this.requestJson<unknown>(`/api/user?${query.toString()}`, {
      signal: options.signal,
    });

    // Current Schoolbox versions return { data, metadata }. Accepting a bare
    // array also makes upgrades from older installations less brittle.
    if (Array.isArray(value)) {
      return {
        data: value.map(validateUser),
        metadata: {},
      };
    }
    if (!isRecord(value) || !Array.isArray(value.data)) {
      throw this.invalidResponseError("GET", "/api/user", "Expected a user list object.", value);
    }

    const metadataValue = isRecord(value.metadata) ? value.metadata : {};
    const rawCursor = metadataValue.cursor;
    const cursorValue = isRecord(rawCursor) ? rawCursor : undefined;
    const scalarCursor =
      typeof rawCursor === "string" || typeof rawCursor === "number"
        ? String(rawCursor)
        : undefined;
    const current = cursorValue?.current ?? (scalarCursor !== undefined ? options.cursor : undefined);
    // Newer Schoolbox releases document `metadata.cursor.next`. Some older
    // installations instead return the next cursor directly as a string or
    // number in `metadata.cursor`. Support both forms so a full directory is
    // not silently reduced to the first page.
    const next = cursorValue?.next ?? scalarCursor ?? metadataValue.next ?? value.next;
    const hasCursorMetadata =
      Object.prototype.hasOwnProperty.call(metadataValue, "cursor") ||
      typeof next === "string" ||
      typeof next === "number";
    const metadata: SchoolboxListMetadata = {
      ...(asFiniteNumber(metadataValue.count) !== undefined
        ? { count: asFiniteNumber(metadataValue.count) }
        : {}),
      ...(hasCursorMetadata
        ? {
            cursor: {
              current:
                typeof current === "string" || typeof current === "number"
                  ? String(current)
                  : null,
              next:
                (typeof next === "string" || typeof next === "number") && String(next).length > 0
                  ? String(next)
                  : null,
            },
          }
        : {}),
    };

    return {
      data: value.data.map(validateUser),
      metadata,
    };
  }

  async getAllUsers(options: SchoolboxGetAllUsersOptions = {}): Promise<SchoolboxUser[]> {
    const pageLimit = positiveInteger(options.pageLimit ?? this.userPageLimit, "pageLimit");
    const users: SchoolboxUser[] = [];
    const seenCursors = new Set<string>();
    let cursor: string | undefined;

    for (let pageNumber = 0; pageNumber < MAX_PAGINATION_PAGES; pageNumber += 1) {
      const page = await this.getUsersPage({
        cursor,
        limit: pageLimit,
        signal: options.signal,
      });
      users.push(...page.data);

      // Older Schoolbox installations may always emit a scalar cursor, even
      // on their final non-empty page. The total lets us finish without an
      // unnecessary empty request or a false repeated-cursor error.
      if (page.metadata.count !== undefined && users.length >= page.metadata.count) {
        return users;
      }

      const nextCursor = page.metadata.cursor?.next ?? undefined;
      if (!nextCursor) {
        return users;
      }
      if (seenCursors.has(nextCursor) || nextCursor === cursor) {
        throw this.invalidResponseError(
          "GET",
          "/api/user",
          "Schoolbox returned a repeated pagination cursor.",
          page.metadata,
        );
      }
      seenCursors.add(nextCursor);
      cursor = nextCursor;
    }

    throw this.invalidResponseError(
      "GET",
      "/api/user",
      "Schoolbox user pagination exceeded its safety limit.",
      { pages: MAX_PAGINATION_PAGES },
    );
  }

  async getRawCalendarEvents(
    userId: number,
    range: SchoolboxCalendarRange = {},
  ): Promise<RawSchoolboxCalendarEvent[]> {
    positiveInteger(userId, "userId");
    const resolved = this.resolveCalendarRange(range);
    const events: RawSchoolboxCalendarEvent[] = [];
    let cursor = resolved.start;

    while (cursor.getTime() < resolved.end.getTime()) {
      const nextMonth = startOfFollowingUtcMonth(cursor);
      const chunkEnd = new Date(Math.min(nextMonth.getTime(), resolved.end.getTime()));
      const query = new URLSearchParams({
        start: toUnixSeconds(cursor),
        end: toUnixSeconds(chunkEnd),
        userId: String(userId),
      });
      const value = await this.requestJson<unknown>(`/calendar/ajax/full?${query.toString()}`, {
        signal: resolved.signal,
      });
      if (!Array.isArray(value)) {
        throw this.invalidResponseError(
          "GET",
          "/calendar/ajax/full",
          "Expected an array of calendar events.",
          value,
        );
      }
      events.push(...value.map(validateRawCalendarEvent));
      cursor = chunkEnd;
    }

    return events;
  }

  async getCalendarEvents(
    userId: number,
    range: SchoolboxCalendarRange = {},
  ): Promise<NormalizedSchoolboxCalendarEvent[]> {
    const events = await this.getRawCalendarEvents(userId, range);
    return events.map((event) => normalizeSchoolboxCalendarEvent(event, userId, this.baseUrl));
  }

  /**
   * Verify both superuser user-list access and delegated calendar access.
   * A short eight-day window is used so this remains inexpensive.
   */
  async testConnection(
    options: SchoolboxConnectionTestOptions = {},
  ): Promise<SchoolboxConnectionTestResult> {
    const page = await this.getUsersPage({ limit: 10, signal: options.signal });
    if (page.data.length === 0 && options.userId === undefined) {
      throw this.invalidResponseError(
        "GET",
        "/api/user",
        "The JWT can list users, but Schoolbox returned no user to test calendar delegation with.",
        page,
      );
    }

    let calendarUser: SchoolboxUser | undefined;
    if (options.userId !== undefined) {
      positiveInteger(options.userId, "userId");
      calendarUser = page.data.find((user) => user.id === options.userId) ?? {
        id: options.userId,
      };
    } else {
      // Prefer a normal active user: this demonstrates access outside the
      // superuser account for which the JWT was generated.
      calendarUser =
        page.data.find((user) => !user.superuser && user.enabled !== false && !user.isDeleted) ??
        page.data.find((user) => user.enabled !== false && !user.isDeleted) ??
        page.data[0];
    }

    if (!calendarUser) {
      throw this.invalidResponseError(
        "GET",
        "/api/user",
        "No Schoolbox user was available for the delegated calendar test.",
        page,
      );
    }

    const now = options.now === undefined ? new Date() : toDate(options.now, "now");
    const start = new Date(now.getTime() - DAY_MS);
    const end = new Date(now.getTime() + 7 * DAY_MS);
    const events = await this.getCalendarEvents(calendarUser.id, {
      start,
      end,
      signal: options.signal,
    });

    return {
      ok: true,
      baseUrl: this.baseUrl,
      usersVisible: page.data.length,
      usersTotal: page.metadata.count ?? null,
      calendarUser: {
        id: calendarUser.id,
        fullName: calendarUser.fullName,
        email: calendarUser.email,
        username: calendarUser.username,
      },
      calendarEventsFound: events.length,
      range: {
        start: start.toISOString(),
        end: end.toISOString(),
      },
    };
  }

  private resolveCalendarRange(range: SchoolboxCalendarRange): ResolvedCalendarRange {
    const now = range.now === undefined ? new Date() : toDate(range.now, "now");
    const pastDays = nonNegativeFinite(range.pastDays ?? this.pastDays, "pastDays");
    const futureDays = nonNegativeFinite(range.futureDays ?? this.futureDays, "futureDays");
    const start =
      range.start === undefined
        ? new Date(now.getTime() - pastDays * DAY_MS)
        : toDate(range.start, "start");
    const end =
      range.end === undefined
        ? new Date(now.getTime() + futureDays * DAY_MS)
        : toDate(range.end, "end");

    if (start.getTime() >= end.getTime()) {
      throw new RangeError("Schoolbox calendar range start must be before end.");
    }
    return { start, end, signal: range.signal };
  }

  private async requestJson<T>(path: string, options: SchoolboxRequestOptions): Promise<T> {
    const url = `${this.baseUrl}${path.startsWith("/") ? path : `/${path}`}`;
    const method = "GET";

    for (let attempt = 0; attempt <= this.maxRetries; attempt += 1) {
      let response: Response;
      try {
        response = await this.fetchImpl(url, {
          method,
          redirect: "error",
          headers: {
            Accept: "application/json",
            Authorization: `Bearer ${this.jwt}`,
          },
          signal: options.signal,
        });
      } catch (error) {
        if (options.signal?.aborted || attempt >= this.maxRetries) {
          throw error;
        }
        await this.sleep(this.exponentialDelay(attempt));
        continue;
      }

      if (response.ok) {
        const text = await response.text();
        if (text.trim().length === 0) {
          return undefined as T;
        }
        try {
          return JSON.parse(text) as T;
        } catch {
          throw this.invalidResponseError(
            method,
            path,
            "Schoolbox returned a successful response that was not valid JSON.",
            text.slice(0, 2_000),
          );
        }
      }

      const retryable =
        response.status === 429 || (response.status >= 500 && response.status <= 599);
      if (retryable && attempt < this.maxRetries) {
        const retryAfter = parseRetryAfter(response.headers.get("retry-after"));
        try {
          await response.body?.cancel();
        } catch {
          // Ignored: consuming an error body is not required before retrying in
          // all fetch implementations.
        }
        const delay = Math.min(
          this.retryMaxDelayMs,
          retryAfter ?? this.exponentialDelay(attempt),
        );
        await this.sleep(delay);
        continue;
      }

      throw await responseToApiError(response, method, url);
    }

    // The loop always returns, continues, or throws. This is a safety net for
    // future changes to the retry logic.
    throw new SchoolboxApiError({
      method,
      url,
      status: 500,
      type: "urn:schoolbox-client:retry-exhausted",
      title: "Schoolbox request retries were exhausted",
    });
  }

  private exponentialDelay(attempt: number): number {
    return Math.min(this.retryMaxDelayMs, this.retryBaseDelayMs * Math.pow(2, attempt));
  }

  private invalidResponseError(
    method: string,
    path: string,
    detail: string,
    body: unknown,
  ): SchoolboxApiError {
    return new SchoolboxApiError({
      method,
      url: `${this.baseUrl}${path}`,
      status: 502,
      type: "urn:schoolbox-client:invalid-response",
      title: "Invalid response from Schoolbox",
      detail,
      responseBody: body,
    });
  }
}
import type { EventCategory } from "./policy";
