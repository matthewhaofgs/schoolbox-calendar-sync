import assert from "node:assert/strict";
import test from "node:test";

import {
  isMicrosoftGraphConflict,
  isMicrosoftGraphNotFound,
  MicrosoftConfigurationError,
  MicrosoftGraphClient,
  MicrosoftGraphError,
  parseMicrosoftCredentials,
} from "../lib/microsoft.ts";

const credentials = {
  tenantId: "11111111-1111-4111-8111-111111111111",
  clientId: "22222222-2222-4222-8222-222222222222",
  clientSecret: "unit-test-secret",
};

function tokenResponse(token = "test-access-token") {
  return Response.json({
    access_token: token,
    token_type: "Bearer",
    expires_in: 3600,
  });
}

test("strictly validates Entra credentials without exposing the secret", () => {
  assert.deepEqual(parseMicrosoftCredentials(credentials), credentials);
  assert.throws(
    () => parseMicrosoftCredentials({ ...credentials, tenantId: "organizations" }),
    MicrosoftConfigurationError,
  );
  assert.throws(
    () => parseMicrosoftCredentials({ ...credentials, clientId: "00000000-0000-0000-0000-000000000000" }),
    /non-empty GUID/,
  );
  assert.throws(
    () => parseMicrosoftCredentials({ ...credentials, clientSecret: " secret" }),
    (error) => {
      assert.equal(error instanceof MicrosoftConfigurationError, true);
      assert.equal(error.message.includes("secret"), true);
      assert.equal(error.message.includes("unit-test-secret"), false);
      return true;
    },
  );
});

test("uses client credentials, caches the token, and paginates selected users", async () => {
  const requests = [];
  let tokenRequests = 0;
  const client = new MicrosoftGraphClient(credentials, {
    fetch: async (input, init) => {
      const url = new URL(String(input));
      if (url.hostname === "login.microsoftonline.com") {
        tokenRequests += 1;
        const form = new URLSearchParams(String(init?.body));
        assert.equal(form.get("client_id"), credentials.clientId);
        assert.equal(form.get("client_secret"), credentials.clientSecret);
        assert.equal(form.get("scope"), "https://graph.microsoft.com/.default");
        assert.equal(form.get("grant_type"), "client_credentials");
        return tokenResponse();
      }

      requests.push({ url, headers: new Headers(init?.headers) });
      if (url.searchParams.has("$skiptoken")) {
        return Response.json({
          value: [
            {
              id: "user-2",
              userPrincipalName: "second@example.test",
              accountEnabled: false,
            },
          ],
        });
      }
      return Response.json({
        value: [
          {
            id: "user-1",
            userPrincipalName: "first@example.test",
            accountEnabled: true,
          },
        ],
        "@odata.nextLink":
          "https://graph.microsoft.com/v1.0/users?$select=id&$skiptoken=next",
      });
    },
  });

  const pages = [];
  const users = await client.listAllUsers({
    includeDisabled: false,
    top: 50,
    onPage: (page) => pages.push(page),
  });
  assert.deepEqual(users.map((user) => user.id), ["user-1"]);
  assert.equal(tokenRequests, 1);
  assert.equal(requests.length, 2);
  assert.equal(requests[0].url.pathname, "/v1.0/users");
  assert.equal(requests[0].url.searchParams.get("$top"), "50");
  assert.match(requests[0].url.searchParams.get("$select"), /^id,/);
  assert.match(requests[0].url.searchParams.get("$select"), /userPrincipalName/);
  assert.equal(requests[0].headers.get("authorization"), "Bearer test-access-token");
  assert.deepEqual(pages.map((page) => page.hasNextPage), [true, false]);

  await client.listAllUsers({ top: 1 });
  assert.equal(tokenRequests, 1, "a valid application token should be reused");
});

test("manages primary and secondary calendars and events through app-only user paths", async () => {
  const graphRequests = [];
  const client = new MicrosoftGraphClient(credentials, {
    fetch: async (input, init) => {
      const url = new URL(String(input));
      if (url.hostname === "login.microsoftonline.com") return tokenResponse();
      const request = {
        url,
        method: init?.method ?? "GET",
        headers: new Headers(init?.headers),
        body: init?.body ? JSON.parse(String(init.body)) : undefined,
      };
      graphRequests.push(request);
      if (request.method === "DELETE") return new Response(null, { status: 204 });
      if (url.pathname.endsWith("/calendar")) {
        return Response.json({ id: "primary-calendar", name: "Calendar" });
      }
      if (url.pathname.endsWith("/calendars") && request.method === "POST") {
        return Response.json({ id: "secondary-calendar", ...request.body }, { status: 201 });
      }
      if (url.pathname.includes("/events")) {
        return Response.json({ id: "immutable-event", ...request.body });
      }
      return Response.json({ id: "secondary-calendar", ...request.body });
    },
  });

  await client.getPrimaryCalendar("teacher@example.test");
  await client.createCalendar("teacher@example.test", {
    name: "School events",
    color: "lightBlue",
  });
  await client.updateCalendar(
    "teacher@example.test",
    "secondary/calendar",
    { name: "Renamed events" },
    { ifMatchChangeKey: "calendar-change-key" },
  );

  const event = {
    subject: "Test event",
    start: { dateTime: "2026-08-14T09:00:00", timeZone: "Australia/Sydney" },
    end: { dateTime: "2026-08-14T09:30:00", timeZone: "Australia/Sydney" },
    showAs: "busy",
    transactionId: "relay-source-event-42",
  };
  await client.insertEvent("teacher@example.test", event);
  await client.updateEvent(
    "teacher@example.test",
    "event/id",
    event,
    { calendarId: "secondary/calendar", ifMatchChangeKey: "event-change-key" },
  );
  await client.deleteEvent(
    "teacher@example.test",
    "event/id",
    { calendarId: "secondary/calendar" },
  );
  await client.deleteCalendar("teacher@example.test", "secondary/calendar");

  assert.equal(graphRequests[0].url.pathname, "/v1.0/users/teacher%40example.test/calendar");
  assert.equal(graphRequests[1].url.pathname, "/v1.0/users/teacher%40example.test/calendars");
  assert.equal(graphRequests[2].method, "PATCH");
  assert.match(graphRequests[2].url.pathname, /secondary%2Fcalendar$/);
  assert.equal(graphRequests[2].headers.get("if-match"), "calendar-change-key");

  const inserted = graphRequests[3];
  assert.equal(inserted.url.pathname, "/v1.0/users/teacher%40example.test/calendar/events");
  assert.equal(inserted.headers.get("prefer"), 'IdType="ImmutableId"');
  assert.equal(inserted.body.transactionId, "relay-source-event-42");

  const updated = graphRequests[4];
  assert.equal(updated.method, "PATCH");
  assert.match(updated.url.pathname, /calendars\/secondary%2Fcalendar\/events\/event%2Fid$/);
  assert.equal(updated.headers.get("prefer"), 'IdType="ImmutableId"');
  assert.equal(updated.headers.get("if-match"), "event-change-key");
  assert.equal(updated.body.transactionId, undefined);

  const deleted = graphRequests[5];
  assert.equal(deleted.method, "DELETE");
  assert.equal(deleted.headers.get("prefer"), 'IdType="ImmutableId"');
  assert.match(graphRequests[6].url.pathname, /calendars\/secondary%2Fcalendar$/);
});

test("omits the create-only transactionId from event PATCH requests", async () => {
  let patchBody;
  const client = new MicrosoftGraphClient(credentials, {
    fetch: async (input, init) => {
      const url = new URL(String(input));
      if (url.hostname === "login.microsoftonline.com") return tokenResponse();
      assert.equal(init?.method, "PATCH");
      patchBody = JSON.parse(String(init?.body));
      return Response.json({ id: "updated-event", ...patchBody });
    },
  });
  const event = {
    subject: "Updated event",
    start: { dateTime: "2026-08-14T10:00:00", timeZone: "Australia/Sydney" },
    end: { dateTime: "2026-08-14T10:30:00", timeZone: "Australia/Sydney" },
    transactionId: "create-only-idempotency-key",
  };

  await client.updateEvent("mailbox-id", "event-id", event);

  assert.equal(Object.hasOwn(patchBody, "transactionId"), false);
  assert.equal(patchBody.subject, "Updated event");
  assert.equal(event.transactionId, "create-only-idempotency-key");
});

test("returns structured privacy-safe Graph errors and classifies common outcomes", async () => {
  const client = new MicrosoftGraphClient(credentials, {
    maxRetries: 0,
    fetch: async (input) => {
      const url = new URL(String(input));
      if (url.hostname === "login.microsoftonline.com") return tokenResponse();
      return Response.json(
        {
          error: {
            code: "ErrorItemNotFound",
            message: "private.person@example.test was not found",
            innerError: {
              "request-id": "body-request-id",
              date: "2026-08-14T01:00:00Z",
            },
          },
        },
        {
          status: 404,
          headers: {
            "request-id": "header-request-id",
            date: "Fri, 14 Aug 2026 01:00:00 GMT",
          },
        },
      );
    },
  });

  await assert.rejects(
    client.deleteEvent("user-id", "missing-event"),
    (error) => {
      assert.equal(error instanceof MicrosoftGraphError, true);
      assert.equal(error.status, 404);
      assert.equal(error.code, "ErrorItemNotFound");
      assert.equal(error.requestId, "header-request-id");
      assert.equal(error.date, "Fri, 14 Aug 2026 01:00:00 GMT");
      assert.equal(isMicrosoftGraphNotFound(error), true);
      assert.equal(isMicrosoftGraphConflict(error), false);
      assert.equal(error.message.includes("private.person"), false);
      assert.equal(JSON.stringify(error).includes("private.person"), false);
      return true;
    },
  );

  const conflict = new MicrosoftGraphError("Conflict", {
    status: 409,
    code: "ErrorItemAlreadyExists",
  });
  assert.equal(isMicrosoftGraphConflict(conflict), true);
});

test("honours Retry-After and bounds only exponential fallback retries", async () => {
  let graphAttempts = 0;
  const delays = [];
  const client = new MicrosoftGraphClient(credentials, {
    maxRetries: 2,
    retryBaseDelayMs: 20,
    maxRetryDelayMs: 50,
    sleep: async (delayMs) => delays.push(delayMs),
    fetch: async (input) => {
      const url = new URL(String(input));
      if (url.hostname === "login.microsoftonline.com") return tokenResponse();
      graphAttempts += 1;
      if (graphAttempts === 1) {
        return new Response(null, {
          status: 429,
          headers: { "retry-after": "10" },
        });
      }
      if (graphAttempts === 2) return new Response(null, { status: 503 });
      return Response.json({ value: [] });
    },
  });

  assert.deepEqual(await client.listAllUsers(), []);
  assert.equal(graphAttempts, 3);
  assert.deepEqual(delays, [10_000, 40]);
});

test("connection test checks directory, primary-calendar, and secondary-calendar write access", async () => {
  const paths = [];
  const client = new MicrosoftGraphClient(credentials, {
    fetch: async (input, init) => {
      const url = new URL(String(input));
      if (url.hostname === "login.microsoftonline.com") return tokenResponse();
      paths.push(`${url.pathname}${url.search}`);
      if (init?.method === "DELETE") return new Response(null, { status: 204 });
      if (url.pathname === "/v1.0/users") {
        return Response.json({ value: [{ id: "sample-user" }] });
      }
      if (url.pathname.endsWith("/calendars")) return Response.json({ id: "connection-test-calendar" }, { status: 201 });
      return Response.json({ id: "primary-calendar" });
    },
  });

  const result = await client.testConnection({ targetUserId: "selected-user" });
  assert.equal(result.ok, true);
  assert.equal(result.directory.sampleUsers, 1);
  assert.equal(result.calendar.targetUserId, "selected-user");
  assert.equal(result.calendar.primaryCalendarId, "primary-calendar");
  assert.equal(result.calendar.secondaryCalendarManagement, true);
  assert.match(paths[0], /^\/v1\.0\/users\?/);
  assert.match(paths[0], /%24top=1/);
  assert.match(paths[1], /^\/v1\.0\/users\/selected-user\/calendar\?/);
  assert.equal(paths[2], "/v1.0/users/selected-user/calendars");
  assert.equal(paths[3], "/v1.0/users/selected-user/calendars/connection-test-calendar");
});

test("connection write verification requires an explicit target mailbox", async () => {
  let requestCount = 0;
  const client = new MicrosoftGraphClient(credentials, {
    fetch: async () => {
      requestCount += 1;
      return tokenResponse();
    },
  });

  await assert.rejects(
    client.testConnection(),
    (error) => {
      assert.equal(error instanceof MicrosoftConfigurationError, true);
      assert.match(error.message, /targetUserId is required/);
      return true;
    },
  );
  assert.equal(requestCount, 0, "no arbitrary directory mailbox should be probed");
});

test("connection test explains missing Graph application permissions", async () => {
  const client = new MicrosoftGraphClient(credentials, {
    fetch: async (input) => {
      const url = new URL(String(input));
      if (url.hostname === "login.microsoftonline.com") return tokenResponse();
      return Response.json({
        error: {
          code: "Authorization_RequestDenied",
          message: "private.person@example.test is not authorized",
        },
      }, { status: 403 });
    },
  });

  await assert.rejects(
    client.testConnection({ targetUserId: "selected-mailbox-id" }),
    (error) => {
      assert.equal(error instanceof MicrosoftConfigurationError, true);
      assert.match(error.message, /User\.Read\.All/);
      assert.match(error.message, /Calendars\.ReadWrite/);
      assert.match(error.message, /Application permissions \(not Delegated permissions\)/);
      assert.match(error.message, /Grant admin consent/);
      assert.equal(error.message.includes("private.person"), false);
      assert.equal(error.message.includes("selected-mailbox-id"), false);
      return true;
    },
  );
});

test("connection test distinguishes mailbox policy denial from directory permission denial", async () => {
  const client = new MicrosoftGraphClient(credentials, {
    fetch: async (input) => {
      const url = new URL(String(input));
      if (url.hostname === "login.microsoftonline.com") return tokenResponse();
      if (url.pathname === "/v1.0/users") {
        return Response.json({ value: [{ id: "directory-sample" }] });
      }
      return Response.json({
        error: { code: "ErrorAccessDenied", message: "mailbox policy denied access" },
      }, { status: 403 });
    },
  });

  await assert.rejects(
    client.testConnection({ targetUserId: "selected-mailbox-id" }),
    (error) => {
      assert.equal(error instanceof MicrosoftConfigurationError, true);
      assert.match(error.message, /Calendars\.ReadWrite/);
      assert.match(error.message, /Exchange Application RBAC or Application Access Policy/);
      assert.equal(error.message.includes("selected-mailbox-id"), false);
      return true;
    },
  );
});

test("connection test surfaces actionable privacy-safe temporary calendar cleanup failures", async () => {
  const controller = new AbortController();
  const requests = [];
  const temporaryCalendarId = "opaque-connection-test-calendar-id";
  const client = new MicrosoftGraphClient(credentials, {
    maxRetries: 0,
    fetch: async (input, init) => {
      const url = new URL(String(input));
      if (url.hostname === "login.microsoftonline.com") return tokenResponse();
      requests.push({ path: url.pathname, method: init?.method ?? "GET" });
      if (url.pathname === "/v1.0/users") {
        return Response.json({ value: [{ id: "directory-sample" }] });
      }
      if (url.pathname.endsWith("/calendar")) {
        return Response.json({ id: "primary-calendar" });
      }
      if (init?.method === "POST") {
        controller.abort(new Error("caller stopped waiting"));
        return Response.json({ id: temporaryCalendarId }, { status: 201 });
      }
      if (init?.method === "DELETE") {
        return Response.json(
          {
            error: {
              code: "ErrorAccessDenied",
              message: "private.person@example.test cannot delete this calendar",
            },
          },
          {
            status: 403,
            headers: { "request-id": "cleanup-request-id" },
          },
        );
      }
      throw new Error(`Unexpected Microsoft Graph request: ${url.pathname}`);
    },
  });

  await assert.rejects(
    client.testConnection({
      targetUserId: "selected-mailbox-id",
      signal: controller.signal,
    }),
    (error) => {
      assert.equal(error instanceof MicrosoftGraphError, true);
      assert.equal(error.status, 403);
      assert.equal(error.code, "temporary_calendar_cleanup_failed");
      assert.equal(error.requestId, "cleanup-request-id");
      assert.match(error.message, new RegExp(temporaryCalendarId));
      assert.match(error.message, /Delete the calendar.*then run the connection test again/);
      assert.equal(error.message.includes("private.person"), false);
      assert.equal(error.message.includes("selected-mailbox-id"), false);
      return true;
    },
  );
  assert.deepEqual(requests.at(-1), {
    path: `/v1.0/users/selected-mailbox-id/calendars/${temporaryCalendarId}`,
    method: "DELETE",
  });
});

test("an aborted retry delay stops the operation", async () => {
  const controller = new AbortController();
  const stopped = new Error("stop retrying");
  const client = new MicrosoftGraphClient(credentials, {
    fetch: async (input) => {
      const url = new URL(String(input));
      if (url.hostname === "login.microsoftonline.com") return tokenResponse();
      return new Response(null, { status: 429 });
    },
    sleep: async (_delay, signal) => {
      controller.abort(stopped);
      if (signal?.aborted) throw signal.reason;
    },
  });

  await assert.rejects(
    client.listAllUsers({ signal: controller.signal }),
    (error) => error === stopped,
  );
});
