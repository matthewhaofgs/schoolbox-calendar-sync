export const EVENT_CATEGORIES = [
  "timetable",
  "resource_booking",
  "school_event",
  "individual_event",
  "other",
] as const;

export type EventCategory = (typeof EVENT_CATEGORIES)[number];
export type EventTypeFilterMode = "all" | "include" | "exclude";
export type GoogleEventVisibility = "default" | "private" | "public";
export type GoogleEventTransparency = "opaque" | "transparent";
export type ReminderMode = "calendar_default" | "none" | "custom";
export type ReminderMethod = "popup" | "email";

export type ManagedCalendarDefinition = {
  /** Stable Relay identifier. The Google calendar ID is stored per user. */
  id: string;
  name: string;
  description: string;
};

export type GoogleEventRuleOverride = {
  /** Exact-type rules can override category coverage. */
  enabled?: boolean;
  destinationId?: string;
  visibility?: GoogleEventVisibility;
  transparency?: GoogleEventTransparency;
  /** Empty string explicitly selects the destination calendar's default colour. */
  colorId?: string;
  reminderMode?: ReminderMode;
  reminderMethod?: ReminderMethod;
  reminderMinutes?: number;
};

export type ResolvedGoogleEventRule = {
  destinationId: string;
  visibility: GoogleEventVisibility;
  transparency: GoogleEventTransparency;
  colorId: string;
  reminderMode: ReminderMode;
  reminderMethod: ReminderMethod;
  reminderMinutes: number;
};

export type SyncPolicy = {
  categories: Record<EventCategory, boolean>;
  eventTypeMode: EventTypeFilterMode;
  eventTypes: string[];
  includeAllDayEvents: boolean;
  includeTimedEvents: boolean;
  includeCompletedEvents: boolean;
  includeDescription: boolean;
  includeLocation: boolean;
  includeSchoolboxLink: boolean;
  includeEventTypeInDescription: boolean;
  includeAuthorInDescription: boolean;
  titlePrefix: string;
  visibility: GoogleEventVisibility;
  transparency: GoogleEventTransparency;
  colorId: string;
  reminderMode: ReminderMode;
  reminderMethod: ReminderMethod;
  reminderMinutes: number;
  defaultDestinationId: string;
  secondaryCalendars: ManagedCalendarDefinition[];
  categoryOverrides: Partial<Record<EventCategory, GoogleEventRuleOverride>>;
  eventTypeOverrides: Record<string, GoogleEventRuleOverride>;
  deleteMissingEvents: boolean;
  deleteExcludedEvents: boolean;
};

export type SyncPolicyInput = Partial<Omit<SyncPolicy, "categories">> & {
  categories?: Partial<Record<EventCategory, boolean>>;
};

export const DEFAULT_SYNC_POLICY: SyncPolicy = {
  categories: {
    timetable: true,
    resource_booking: true,
    school_event: true,
    individual_event: true,
    other: true,
  },
  eventTypeMode: "all",
  eventTypes: [],
  includeAllDayEvents: true,
  includeTimedEvents: true,
  includeCompletedEvents: true,
  includeDescription: true,
  includeLocation: true,
  includeSchoolboxLink: true,
  includeEventTypeInDescription: false,
  includeAuthorInDescription: false,
  titlePrefix: "",
  visibility: "default",
  transparency: "opaque",
  colorId: "",
  reminderMode: "calendar_default",
  reminderMethod: "popup",
  reminderMinutes: 10,
  defaultDestinationId: "primary",
  secondaryCalendars: [],
  categoryOverrides: {},
  eventTypeOverrides: {},
  deleteMissingEvents: true,
  deleteExcludedEvents: true,
};

function boolean(value: unknown, fallback: boolean): boolean {
  return typeof value === "boolean" ? value : fallback;
}

function enumValue<T extends string>(value: unknown, allowed: readonly T[], fallback: T): T {
  return typeof value === "string" && allowed.includes(value as T) ? value as T : fallback;
}

function integer(value: unknown, fallback: number, min: number, max: number): number {
  return typeof value === "number" && Number.isFinite(value)
    ? Math.max(min, Math.min(max, Math.round(value)))
    : fallback;
}

export function normalizeEventTypeLabel(value: unknown): string {
  return typeof value === "string" ? value.trim().replace(/\s+/g, " ").slice(0, 120) : "";
}

export function eventTypeKey(value: unknown): string {
  return normalizeEventTypeLabel(value).toLocaleLowerCase("en-AU");
}

function normalizedCalendarDefinitions(value: unknown, fallback: ManagedCalendarDefinition[]): ManagedCalendarDefinition[] {
  if (value === undefined) return fallback.map((calendar) => ({ ...calendar }));
  if (!Array.isArray(value)) return [];
  const calendars: ManagedCalendarDefinition[] = [];
  const seenIds = new Set<string>();
  const seenNames = new Set<string>();
  for (const candidate of value) {
    if (!candidate || typeof candidate !== "object") continue;
    const record = candidate as Record<string, unknown>;
    const id = typeof record.id === "string" ? record.id.trim().toLowerCase() : "";
    const name = typeof record.name === "string" ? record.name.trim().replace(/\s+/g, " ").slice(0, 100) : "";
    const description = typeof record.description === "string" ? record.description.trim().slice(0, 500) : "";
    const nameKey = name.toLocaleLowerCase("en-AU");
    if (!/^[a-z0-9][a-z0-9_-]{0,39}$/.test(id) || seenIds.has(id) || (nameKey && seenNames.has(nameKey))) continue;
    seenIds.add(id);
    if (nameKey) seenNames.add(nameKey);
    calendars.push({ id, name, description });
    if (calendars.length >= 20) break;
  }
  return calendars;
}

function normalizedDestinationId(value: unknown, allowed: Set<string>, fallback: string): string {
  return typeof value === "string" && allowed.has(value) ? value : fallback;
}

function normalizedRuleOverride(value: unknown, allowedDestinations: Set<string>): GoogleEventRuleOverride | null {
  if (!value || typeof value !== "object") return null;
  const record = value as Record<string, unknown>;
  const result: GoogleEventRuleOverride = {};
  if (typeof record.enabled === "boolean") result.enabled = record.enabled;
  if (typeof record.destinationId === "string" && allowedDestinations.has(record.destinationId)) {
    result.destinationId = record.destinationId;
  }
  if (["default", "private", "public"].includes(String(record.visibility))) {
    result.visibility = record.visibility as GoogleEventVisibility;
  }
  if (["opaque", "transparent"].includes(String(record.transparency))) {
    result.transparency = record.transparency as GoogleEventTransparency;
  }
  if (typeof record.colorId === "string" && /^(?:|[1-9]|1[01])$/.test(record.colorId)) {
    result.colorId = record.colorId;
  }
  if (["calendar_default", "none", "custom"].includes(String(record.reminderMode))) {
    result.reminderMode = record.reminderMode as ReminderMode;
  }
  if (["popup", "email"].includes(String(record.reminderMethod))) {
    result.reminderMethod = record.reminderMethod as ReminderMethod;
  }
  if (typeof record.reminderMinutes === "number" && Number.isFinite(record.reminderMinutes)) {
    result.reminderMinutes = integer(record.reminderMinutes, 10, 0, 40_320);
  }
  return Object.keys(result).length > 0 ? result : null;
}

function normalizedCategoryOverrides(
  value: unknown,
  fallback: Partial<Record<EventCategory, GoogleEventRuleOverride>>,
  allowedDestinations: Set<string>,
): Partial<Record<EventCategory, GoogleEventRuleOverride>> {
  const source = value === undefined ? fallback : value;
  if (!source || typeof source !== "object") return {};
  const record = source as Record<string, unknown>;
  const result: Partial<Record<EventCategory, GoogleEventRuleOverride>> = {};
  for (const category of EVENT_CATEGORIES) {
    const rule = normalizedRuleOverride(record[category], allowedDestinations);
    if (rule) result[category] = rule;
  }
  return result;
}

function normalizedEventTypeOverrides(
  value: unknown,
  fallback: Record<string, GoogleEventRuleOverride>,
  allowedDestinations: Set<string>,
): Record<string, GoogleEventRuleOverride> {
  const source = value === undefined ? fallback : value;
  if (!source || typeof source !== "object") return {};
  const result: Record<string, GoogleEventRuleOverride> = {};
  for (const [candidateKey, candidateRule] of Object.entries(source as Record<string, unknown>)) {
    const key = eventTypeKey(candidateKey);
    const rule = normalizedRuleOverride(candidateRule, allowedDestinations);
    if (!key || !rule || result[key]) continue;
    result[key] = rule;
    if (Object.keys(result).length >= 200) break;
  }
  return result;
}

export function normalizeSyncPolicy(input: unknown, fallback: SyncPolicy = DEFAULT_SYNC_POLICY): SyncPolicy {
  const value = input && typeof input === "object" ? input as Record<string, unknown> : {};
  const categories = value.categories && typeof value.categories === "object"
    ? value.categories as Record<string, unknown>
    : {};
  const typeValues = Array.isArray(value.eventTypes) ? value.eventTypes : fallback.eventTypes;
  const eventTypes: string[] = [];
  const seenTypes = new Set<string>();
  for (const candidate of typeValues) {
    const label = normalizeEventTypeLabel(candidate);
    const key = eventTypeKey(label);
    if (!key || seenTypes.has(key)) continue;
    seenTypes.add(key);
    eventTypes.push(label);
    if (eventTypes.length >= 100) break;
  }
  const colorId = typeof value.colorId === "string" && /^(?:|[1-9]|1[01])$/.test(value.colorId)
    ? value.colorId
    : fallback.colorId;
  const secondaryCalendars = normalizedCalendarDefinitions(value.secondaryCalendars, fallback.secondaryCalendars);
  const allowedDestinations = new Set(["primary", ...secondaryCalendars.map((calendar) => calendar.id)]);
  const fallbackDestination = allowedDestinations.has(fallback.defaultDestinationId)
    ? fallback.defaultDestinationId
    : "primary";
  const defaultDestinationId = normalizedDestinationId(value.defaultDestinationId, allowedDestinations, fallbackDestination);

  return {
    categories: Object.fromEntries(EVENT_CATEGORIES.map((category) => [
      category,
      boolean(categories[category], fallback.categories[category]),
    ])) as Record<EventCategory, boolean>,
    eventTypeMode: enumValue(value.eventTypeMode, ["all", "include", "exclude"] as const, fallback.eventTypeMode),
    eventTypes,
    includeAllDayEvents: boolean(value.includeAllDayEvents, fallback.includeAllDayEvents),
    includeTimedEvents: boolean(value.includeTimedEvents, fallback.includeTimedEvents),
    includeCompletedEvents: boolean(value.includeCompletedEvents, fallback.includeCompletedEvents),
    includeDescription: boolean(value.includeDescription, fallback.includeDescription),
    includeLocation: boolean(value.includeLocation, fallback.includeLocation),
    includeSchoolboxLink: boolean(value.includeSchoolboxLink, fallback.includeSchoolboxLink),
    includeEventTypeInDescription: boolean(value.includeEventTypeInDescription, fallback.includeEventTypeInDescription),
    includeAuthorInDescription: boolean(value.includeAuthorInDescription, fallback.includeAuthorInDescription),
    titlePrefix: typeof value.titlePrefix === "string" ? value.titlePrefix.trim().slice(0, 40) : fallback.titlePrefix,
    visibility: enumValue(value.visibility, ["default", "private", "public"] as const, fallback.visibility),
    transparency: enumValue(value.transparency, ["opaque", "transparent"] as const, fallback.transparency),
    colorId,
    reminderMode: enumValue(value.reminderMode, ["calendar_default", "none", "custom"] as const, fallback.reminderMode),
    reminderMethod: enumValue(value.reminderMethod, ["popup", "email"] as const, fallback.reminderMethod),
    reminderMinutes: integer(value.reminderMinutes, fallback.reminderMinutes, 0, 40_320),
    defaultDestinationId,
    secondaryCalendars,
    categoryOverrides: normalizedCategoryOverrides(value.categoryOverrides, fallback.categoryOverrides, allowedDestinations),
    eventTypeOverrides: normalizedEventTypeOverrides(value.eventTypeOverrides, fallback.eventTypeOverrides, allowedDestinations),
    deleteMissingEvents: boolean(value.deleteMissingEvents, fallback.deleteMissingEvents),
    deleteExcludedEvents: boolean(value.deleteExcludedEvents, fallback.deleteExcludedEvents),
  };
}

export type PolicyEvent = {
  category: EventCategory;
  type: string | null;
  allDay: boolean;
  completed: boolean;
};

export function eventIncludedByPolicy(event: PolicyEvent, policy: SyncPolicy): boolean {
  if (event.allDay ? !policy.includeAllDayEvents : !policy.includeTimedEvents) return false;
  if (event.completed && !policy.includeCompletedEvents) return false;
  const typeOverride = policy.eventTypeOverrides[eventTypeKey(event.type)];
  if (typeOverride?.enabled !== undefined) return typeOverride.enabled;
  if (!policy.categories[event.category]) return false;
  if (policy.eventTypeMode === "all") return true;
  const configured = new Set(policy.eventTypes.map(eventTypeKey));
  const listed = configured.has(eventTypeKey(event.type));
  return policy.eventTypeMode === "include" ? listed : !listed;
}

function applyGoogleRuleOverride(
  current: ResolvedGoogleEventRule,
  override: GoogleEventRuleOverride | undefined,
): ResolvedGoogleEventRule {
  if (!override) return current;
  return {
    destinationId: override.destinationId ?? current.destinationId,
    visibility: override.visibility ?? current.visibility,
    transparency: override.transparency ?? current.transparency,
    colorId: override.colorId ?? current.colorId,
    reminderMode: override.reminderMode ?? current.reminderMode,
    reminderMethod: override.reminderMethod ?? current.reminderMethod,
    reminderMinutes: override.reminderMinutes ?? current.reminderMinutes,
  };
}

/** Resolves default, category, then exact Schoolbox type settings. */
export function resolveGoogleEventRule(event: Pick<PolicyEvent, "category" | "type">, policy: SyncPolicy): ResolvedGoogleEventRule {
  const base: ResolvedGoogleEventRule = {
    destinationId: policy.defaultDestinationId,
    visibility: policy.visibility,
    transparency: policy.transparency,
    colorId: policy.colorId,
    reminderMode: policy.reminderMode,
    reminderMethod: policy.reminderMethod,
    reminderMinutes: policy.reminderMinutes,
  };
  const categoryRule = applyGoogleRuleOverride(base, policy.categoryOverrides[event.category]);
  return applyGoogleRuleOverride(categoryRule, policy.eventTypeOverrides[eventTypeKey(event.type)]);
}
