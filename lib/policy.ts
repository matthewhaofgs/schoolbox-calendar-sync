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
  if (!policy.categories[event.category]) return false;
  if (event.allDay ? !policy.includeAllDayEvents : !policy.includeTimedEvents) return false;
  if (event.completed && !policy.includeCompletedEvents) return false;
  if (policy.eventTypeMode === "all") return true;
  const configured = new Set(policy.eventTypes.map(eventTypeKey));
  const listed = configured.has(eventTypeKey(event.type));
  return policy.eventTypeMode === "include" ? listed : !listed;
}
