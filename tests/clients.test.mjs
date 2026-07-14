import assert from "node:assert/strict";
import test from "node:test";

import {
  normalizeSchoolboxCalendarEvent,
  schoolboxHtmlToText,
  SchoolboxClient,
} from "../lib/schoolbox.ts";
import {
  createContentHash,
  createDeterministicEventId,
} from "../lib/google.ts";

test("normalizes Schoolbox calendar HTML, dates, links and stable identities", () => {
  const raw = {
    title: "Year 9 Mathematics",
    start: "2026-07-14T09:00:00+10:00",
    end: "2026-07-14T09:50:00+10:00",
    editable: false,
    allDay: false,
    className: "timetable source1",
    data: {
      meta: {
        eventId: 412,
        detail: "<p>Chapter 8<br><strong>Bring calculator</strong></p>",
        location: "B14",
        type: "Timetable",
      },
      links: { path: "/calendar/event/412" },
    },
  };

  const event = normalizeSchoolboxCalendarEvent(raw, 88, "https://school.example.edu");
  assert.equal(event.sourceKey, "schoolbox:event:412");
  assert.equal(event.description, "Chapter 8\nBring calculator");
  assert.equal(event.location, "B14");
  assert.equal(event.sourceUrl, "https://school.example.edu/calendar/event/412");
  assert.equal(event.allDay, false);
});

test("repairs an invalid all-day exclusive end date", () => {
  const event = normalizeSchoolboxCalendarEvent(
    { title: "Pupil-free day", start: "2026-08-03", end: "2026-08-03", editable: false, allDay: true },
    2,
  );
  assert.equal(event.start, "2026-08-03");
  assert.equal(event.end, "2026-08-04");
});

test("repairs a missing timed boundary with a 30 minute duration", () => {
  const missingEnd = normalizeSchoolboxCalendarEvent(
    { title: "Due item", start: "2026-08-14T09:00:00+10:00", end: "", editable: false, allDay: false },
    2,
  );
  assert.equal(missingEnd.start, "2026-08-14T09:00:00+10:00");
  assert.equal(missingEnd.end, "2026-08-14T09:30:00+10:00");

  const missingStart = normalizeSchoolboxCalendarEvent(
    { title: "Due item", start: "", end: "2026-08-14T09:00:00+10:00", editable: false, allDay: false },
    2,
  );
  assert.equal(missingStart.start, "2026-08-14T08:30:00+10:00");
  assert.equal(missingStart.end, "2026-08-14T09:00:00+10:00");
});

test("repairs a missing all-day boundary as one calendar day", () => {
  const missingEnd = normalizeSchoolboxCalendarEvent(
    { title: "Closure", start: "2026-08-14", end: "", editable: false, allDay: true },
    2,
  );
  assert.equal(missingEnd.start, "2026-08-14");
  assert.equal(missingEnd.end, "2026-08-15");

  const missingStart = normalizeSchoolboxCalendarEvent(
    { title: "Closure", start: "", end: "2026-08-15", editable: false, allDay: true },
    2,
  );
  assert.equal(missingStart.start, "2026-08-14");
  assert.equal(missingStart.end, "2026-08-15");
});

test("rejects an event with no date anchor", () => {
  assert.throws(
    () => normalizeSchoolboxCalendarEvent(
      { title: "Undated", start: "", end: "", editable: false, allDay: false },
      2,
    ),
    /neither a start nor an end date-time/,
  );
});

test("Schoolbox user pagination follows cursor metadata", async () => {
  const requests = [];
  const client = new SchoolboxClient({
    baseUrl: "https://school.example.edu/",
    jwt: "Bearer test-jwt",
    fetchImpl: async (input, init) => {
      const url = new URL(String(input));
      requests.push({ url, authorization: new Headers(init?.headers).get("authorization") });
      const page = url.searchParams.get("cursor")
        ? { data: [{ id: 2, email: "two@example.edu" }], metadata: { cursor: { current: "next", next: null } } }
        : { data: [{ id: 1, email: "one@example.edu" }], metadata: { cursor: { current: null, next: "next" } } };
      return Response.json(page);
    },
  });

  const users = await client.getAllUsers();
  assert.deepEqual(users.map((user) => user.id), [1, 2]);
  assert.equal(requests.length, 2);
  assert.equal(requests[0].authorization, "Bearer test-jwt");
});

test("Schoolbox user pagination follows legacy scalar cursors", async () => {
  const requestedCursors = [];
  const client = new SchoolboxClient({
    baseUrl: "https://school.example.edu/",
    jwt: "test-jwt",
    fetchImpl: async (input) => {
      const cursor = new URL(String(input)).searchParams.get("cursor");
      requestedCursors.push(cursor);
      if (cursor === null) {
        return Response.json({
          data: [{ id: 1, email: "one@example.edu" }],
          metadata: { count: 2, cursor: 1012 },
        });
      }
      assert.equal(cursor, "1012");
      return Response.json({
        data: [{ id: 2, email: "two@example.edu" }],
        metadata: { count: 2, cursor: null },
      });
    },
  });

  const users = await client.getAllUsers();
  assert.deepEqual(users.map((user) => user.id), [1, 2]);
  assert.deepEqual(requestedCursors, [null, "1012"]);
});

test("Google IDs and hashes are deterministic and Calendar-safe", async () => {
  const firstId = await createDeterministicEventId("tenant:user:schoolbox:event:42");
  const secondId = await createDeterministicEventId("tenant:user:schoolbox:event:42");
  assert.equal(firstId, secondId);
  assert.match(firstId, /^[0-9a-v]{5,1024}$/);

  const a = await createContentHash({ summary: "Assembly", start: { date: "2026-08-01" } });
  const b = await createContentHash({ start: { date: "2026-08-01" }, summary: "Assembly" });
  assert.equal(a, b);
});

test("Schoolbox detail sanitizer removes executable markup", () => {
  assert.equal(
    schoolboxHtmlToText("<script>alert(1)</script><p>Hello&nbsp;<em>world</em></p>"),
    "Hello world",
  );
});
