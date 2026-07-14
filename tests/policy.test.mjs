import assert from "node:assert/strict";
import test from "node:test";

import {
  DEFAULT_SYNC_POLICY,
  eventIncludedByPolicy,
  normalizeSyncPolicy,
} from "../lib/policy.ts";
import { normalizeSchoolboxCalendarEvent } from "../lib/schoolbox.ts";
import { eventBody } from "../lib/sync.ts";

test("sync policy normalizes exact type rules and applies every event filter", () => {
  const policy = normalizeSyncPolicy({
    categories: { resource_booking: false },
    eventTypeMode: "include",
    eventTypes: [" Excursion ", "excursion", "Timetable"],
    includeAllDayEvents: false,
    includeCompletedEvents: false,
    colorId: "99",
    reminderMinutes: 99_999,
  });
  assert.deepEqual(policy.eventTypes, ["Excursion", "Timetable"]);
  assert.equal(policy.colorId, "");
  assert.equal(policy.reminderMinutes, 40_320);
  assert.equal(eventIncludedByPolicy({ category: "school_event", type: "excursion", allDay: false, completed: false }, policy), true);
  assert.equal(eventIncludedByPolicy({ category: "resource_booking", type: "Excursion", allDay: false, completed: false }, policy), false);
  assert.equal(eventIncludedByPolicy({ category: "school_event", type: "Other", allDay: false, completed: false }, policy), false);
  assert.equal(eventIncludedByPolicy({ category: "school_event", type: "Excursion", allDay: true, completed: false }, policy), false);
  assert.equal(eventIncludedByPolicy({ category: "school_event", type: "Excursion", allDay: false, completed: true }, policy), false);
});

test("Schoolbox normalization classifies documented calendar sources", () => {
  const common = {
    title: "Calendar item",
    start: "2026-07-14T09:00:00+10:00",
    end: "2026-07-14T10:00:00+10:00",
    editable: false,
    allDay: false,
  };
  assert.equal(normalizeSchoolboxCalendarEvent({ ...common, className: "timetable source1", data: { meta: { type: "Timetable" } } }, 1).category, "timetable");
  assert.equal(normalizeSchoolboxCalendarEvent({ ...common, resourceId: 12, data: { meta: { type: "Room booking" } } }, 1).category, "resource_booking");
  assert.equal(normalizeSchoolboxCalendarEvent({ ...common, resourceId: 12, data: { meta: { type: "Custom category" } } }, 1).category, "other", "a calendar view resource id alone must not imply a booking");
  assert.equal(normalizeSchoolboxCalendarEvent({ ...common, data: { meta: { type: "School Event" } } }, 1).category, "school_event");
  assert.equal(normalizeSchoolboxCalendarEvent({ ...common, data: { meta: { type: "Personal appointment" } } }, 1).category, "individual_event");
  assert.equal(normalizeSchoolboxCalendarEvent({ ...common, data: { meta: { type: "Custom category" } } }, 1).category, "other");
});

test("Google event rendering respects content, privacy, colour and reminder settings", async () => {
  const event = {
    sourceKey: "schoolbox:event:42",
    eventId: "42",
    userId: 1,
    title: "Excursion",
    description: "Bring lunch",
    location: "Museum",
    type: "Excursion",
    typeCode: "type3",
    typeId: 3,
    category: "school_event",
    completed: false,
    author: "Calendar Coordinator",
    sourceUrl: "https://school.example.edu/calendar/event/42",
    start: "2026-07-14T09:00:00+10:00",
    end: "2026-07-14T15:00:00+10:00",
    allDay: false,
    editable: false,
    raw: {},
  };
  const policy = normalizeSyncPolicy({
    includeDescription: false,
    includeLocation: false,
    includeSchoolboxLink: false,
    includeEventTypeInDescription: true,
    includeAuthorInDescription: true,
    titlePrefix: "[Relay]",
    visibility: "private",
    transparency: "transparent",
    colorId: "9",
    reminderMode: "custom",
    reminderMethod: "email",
    reminderMinutes: 30,
  });
  const rendered = await eventBody(event, "google-user", "Australia/Sydney", "source-occurrence", policy);
  assert.equal(rendered.summary, "[Relay] Excursion");
  assert.equal(rendered.description, "Schoolbox type: Excursion\n\nSchoolbox author: Calendar Coordinator");
  assert.equal(rendered.location, undefined);
  assert.equal(rendered.source, undefined);
  assert.equal(rendered.visibility, "private");
  assert.equal(rendered.transparency, "transparent");
  assert.equal(rendered.colorId, "9");
  assert.deepEqual(rendered.reminders, { useDefault: false, overrides: [{ method: "email", minutes: 30 }] });

  const defaults = await eventBody(event, "google-user", "Australia/Sydney", "source-occurrence", normalizeSyncPolicy({}, DEFAULT_SYNC_POLICY));
  assert.match(defaults.description ?? "", /Bring lunch/);
  assert.ok(defaults.source);
  assert.equal(defaults.visibility, undefined);
  assert.equal(defaults.transparency, undefined);
  assert.equal(defaults.colorId, undefined);
  assert.equal(defaults.reminders, undefined);
});
