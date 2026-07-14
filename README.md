# Relay

Relay is a self-hosted, one-way calendar bridge from Schoolbox to Google Workspace. It discovers users by matching each Google primary email to a unique Schoolbox primary or alternate email, reads each enabled user's Schoolbox calendar feed, applies an administrator-defined event policy, and reconciles Relay-owned events into configured primary or app-managed secondary Google calendars.

Relay includes timetable lessons, resource bookings, school events, and individual events. Google calendar access uses a service account with Domain-Wide Delegation, so users do not install an app or complete individual consent.

Relay is an independent project and is not affiliated with or endorsed by Schoolbox or Google. This public repository contains source code and safe examples only; never commit a live environment file, credential, database, backup, or exported calendar data.

## Deployment model

Relay is designed for one internal server and one running application replica:

- The Node server listens on `0.0.0.0:3000` so it can accept requests on every server interface.
- An internal DNS name and HTTPS reverse proxy should be placed in front of port 3000.
- Firewall access should be restricted to the IT network or VPN. Do not publish Relay to the internet.
- Operational state is stored in a named SQLite database. Docker uses the `relay-calendar-data` volume.
- The web process and authenticated scheduler run together under `npm start`.
- Relay makes outbound HTTPS requests only to the configured Schoolbox server and official Google identity, Directory, OAuth, and Calendar endpoints.

Google web OAuth redirect URIs require HTTPS except on localhost. Use a hostname under a domain the school controls, such as `relay.school.edu.au`, even if its DNS record is internal-only.

## Authentication and roles

Relay has two authentication paths:

1. An immutable local break-glass administrator, created interactively on the server. Its password is PBKDF2-HMAC-SHA-256 hashed; plaintext is never stored.
2. Google Workspace OpenID Connect for explicitly approved IT staff. Relay requests only `openid`, `email`, and `profile`; it does not store Google access or refresh tokens.

Workspace membership alone does not grant access. The local administrator or an existing Administrator adds each staff email and selects a role:

| Role | Access |
| --- | --- |
| Viewer | Dashboard, people/mappings, and run history |
| Operator | Viewer access plus diagnostics and manual syncs |
| Administrator | Operator access plus Schoolbox/Google connections, schedule, Google sign-in configuration, and staff/role management |
| Local administrator | All administrator rights plus ownership of the local break-glass password |

All permissions are checked by the server. Sessions use opaque random tokens stored as hashes in SQLite, HTTP-only cookies, an eight-hour maximum life, a 30-minute idle timeout, exact-origin checks, and CSRF tokens. The UI warns five minutes before the earliest session deadline and lets the administrator explicitly extend it; an expired session returns directly to sign-in. Five failed local password attempts temporarily lock the account for 15 minutes.

## Requirements

- A Linux server with Docker Engine and Docker Compose for the container deployment
- Node.js 22.13 or newer, including npm, on the machine used to run the provided environment generator; native deployments also use Node.js to run Relay
- Internal DNS and a trusted TLS certificate for the Relay hostname
- Schoolbox 26.0 or newer
- A Schoolbox superuser JWT that can list users and read another user's calendar
- A Google Cloud project with Admin SDK and Google Calendar API enabled
- A Google service account with Domain-Wide Delegation
- A Google Workspace super admin or delegated administrator account able to list Directory users
- A Google Web OAuth client with an Internal audience for IT staff sign-in

## Docker deployment

### 1. Prepare DNS, TLS, and the firewall

Create an internal DNS record such as `relay.school.edu.au`. Install a trusted certificate on nginx, Caddy, IIS, or the school's existing load balancer. Permit HTTPS only from the IT subnet or VPN. Port `3000` is unencrypted application traffic and bypasses reverse-proxy controls, so the host firewall should allow it only from the reverse proxy (or loopback when the proxy is on the Relay host), never from the general LAN.

An nginx starting point is provided at `deploy/nginx-relay.conf.example`. It includes an HTTP-to-HTTPS redirect and assumes nginx runs on the Relay host. Replace the hostname and certificate paths before enabling it. If the reverse proxy runs elsewhere, replace `127.0.0.1:3000` with Relay's private address and allow port `3000` only from that proxy address.

### 2. Generate the private environment

The provided generator is a host-side Node.js command. Install Node.js 22.13 or newer, including npm, on the machine where the repository is prepared, then run this from the project directory:

```bash
npm run setup:env
```

No npm dependency installation is needed for this generator. Enter the final HTTPS origin. The command creates ignored `.env.production` with independent random values for credential encryption, browser sessions, and scheduler authentication. Back up this file securely; it is required to decrypt stored credentials. The later administrator bootstrap command runs inside the built container and does not require host-installed application dependencies.

### 3. Build and bootstrap the local administrator

```bash
docker compose build
docker compose run --rm relay node scripts/bootstrap-admin.mjs
```

Choose a username (the default is `administrator`) and a password of at least 14 characters. The command is intentionally interactive and refuses to replace an existing owner.

### 4. Start Relay

```bash
docker compose up -d
docker compose ps
```

The container listens on `0.0.0.0:3000`. Open the HTTPS hostname through the reverse proxy; the raw IP/HTTP address is not a valid production login origin.

Check health without authentication:

```bash
curl http://127.0.0.1:3000/api/health
```

Expected result: `{"ok":true}`.

## Native Node deployment

For a server managed by systemd or another process supervisor:

```bash
npm ci
npm run setup:env
npm run auth:bootstrap
npm run build
npm start
```

`npm start` launches both the Next.js server and scheduler. Configure the supervisor with the project directory as its working directory and restart the service on failure. Do not run more than one replica while using SQLite.

For local development only:

```bash
npm run setup:dev-env
npm run auth:bootstrap
npm run dev
```

The development origin is `http://127.0.0.1:3000`; production requires HTTPS.

## First sign-in and IT access

1. Sign in with the local administrator account.
2. Open **IT access**.
3. In Google Cloud, configure the OAuth consent screen with an **Internal** audience.
4. Create an OAuth client of type **Web application**.
5. Copy the exact authorised redirect URI shown by Relay into the Google client.
6. Save the client ID, client secret, and Workspace domain in Relay. The secret is encrypted at rest.
7. Add each IT staff member's primary Google Workspace email and select Viewer, Operator, or Administrator.
8. Staff can now choose **Continue with Google Workspace** on the Relay login screen.

After this bootstrap, any enabled Administrator can add, change, disable, or remove other Google Workspace staff entries. Only the immutable local administrator can change the break-glass password.

The `hd` login hint is not trusted by itself. Relay verifies the signed Google ID token, issuer, audience, expiry, nonce, verified email, Workspace domain, and stable Google subject before creating a session.

## Schoolbox and calendar setup

1. In Schoolbox Admin, edit the superuser Relay will use, scroll to `TOKENS`, and choose `Create token`.
2. Sign in to Relay as an Administrator and open **Setup**.
3. Enter the HTTPS Schoolbox base URL and JWT, then run the connection test.
4. In Google Cloud, enable the Admin SDK and Google Calendar API.
5. Create a dedicated service account, enable Domain-Wide Delegation, and download its JSON key.
6. In Google Admin, open **Security > Access and data control > API controls > Manage Domain Wide Delegation**.
7. Add the service account's numeric Client ID with exactly these scopes:

```text
https://www.googleapis.com/auth/calendar.events.owned,https://www.googleapis.com/auth/calendar.app.created,https://www.googleapis.com/auth/admin.directory.user.readonly
```

8. Paste the service-account JSON and delegated administrator email into Relay, then run the Directory and Calendar tests.
9. Select the schedule, calendar window, and whether newly discovered users should be enabled automatically, then activate Relay.
10. Run the first sync manually and inspect **Runs** and **People**. A fresh installation defaults new users to paused, so this first run can safely populate the user list without writing calendar events.

The Web OAuth client used for IT login and the service account used for calendar synchronization are separate credentials with separate purposes.

After both connections and credentials are saved, **Setup** changes to a completed connection summary. Reopen the wizard only when replacing credentials or delegation; routine sync changes remain under **Settings**.

## Granular sync settings

New and upgraded installations initially include every event category and preserve the original mirror behaviour. Saving settings never starts a sync; test the resulting policy with one enabled pilot user before wider rollout.

- **Schedule** controls intervals from 15 minutes to daily and a rolling calendar window from today through two years ahead. Relay automatically divides Schoolbox calendar requests into the API's recommended month-sized ranges.
- **People** controls whether future Google/Schoolbox matches begin enabled or paused.
- **Event rules** combines source coverage and Google routing on one screen. It defines reusable secondary calendars, global Google defaults, category coverage and overrides, exact-type coverage and overrides, and global timed/all-day/completed safeguards.
- The screen presents the rule order directly: exact Schoolbox type values take priority over category values, which take priority over global defaults. “Inherit” means the setting comes from the preceding level.
- Exact Schoolbox type rules can override source inclusion, destination calendar, visibility, busy/free state, one of Google's event colours, and reminder behaviour. Global timed/all-day/completed safeguards still apply first.
- Type labels observed during normal enabled-user syncs become expandable rule editors; Relay does not scan paused users merely to build the catalogue.
- **Event content** controls descriptions, locations, Schoolbox source links, type and author annotations, and title prefixes.
- **Connections** exposes the saved Schoolbox URL, delegated Google administrator, Directory customer, time zone, service-account identity, encrypted-secret replacement fields, and both connection tests.
- **Reconciliation** separately controls removal when an event disappears from Schoolbox and removal when the current event policy excludes it. Both operations use Relay's managed-event mappings and never select unrelated Google events.
- **Advanced** can pause scheduled runs and tune concurrent per-user calendar work from one to ten. Manual runs remain available while the scheduler is paused.

Schoolbox type labels are installation-defined metadata. Relay combines broad source classification with exact, case-insensitive label rules; unrecognised sources fall under **Other and custom** so they are not silently lost by default.

Secondary calendars are created lazily for each enabled user only when an included event first targets that destination. Relay stores the returned Google calendar ID per user. Changing a type's destination writes the managed event to its new calendar before deleting the old copy; removing a destination from Relay settings does not delete the Google calendar itself. Existing and upgraded policies continue to target the primary calendar until an administrator deliberately changes a rule.

## Pilot and user coverage

Relay discovers every active Google Workspace identity on each run, records its Schoolbox email match, and syncs only users whose **Calendar sync** control is enabled.

- Open **People** to search and filter all discovered users, toggle one user, or select visible rows for a bulk enable/pause action.
- **Remove Relay events** pauses that user and deletes only events recorded in Relay's managed-event mapping table. Other Google Calendar entries are never selected for deletion.
- Open **Settings > People** to choose whether users discovered in future start enabled or paused. Changing this default never changes existing selections.
- Google Workspace accounts without a unique active Schoolbox primary or alternate email match are labelled **Unmatched** as an informational state. This is expected for former staff, service accounts, and other Google-only identities. A unique primary Schoolbox email takes precedence; ambiguous addresses at the same match level are left unmatched to prevent an unsafe identity association.
- Pausing a user stops future updates; it does not remove Relay-created events already present in Google Calendar.
- Viewer and Operator accounts can see coverage. Administrator and local-administrator accounts can change it.
- Existing installations migrate existing users and the new-user default to enabled, preserving the previous sync-all behaviour.

## Persistence, secrets, and backups

The database contains configuration, encrypted credentials, staff access, session hashes, user/event/calendar-target mappings, audit entries, and sync history. Secret fields are AES-256-GCM encrypted with `CONFIG_ENCRYPTION_KEY`.

For Docker, persistent data is in the `relay-calendar-data` volume. Back up the database volume and `.env.production` together. They form one recovery set: restoring the database without its matching encryption key can make saved credentials unrecoverable.

For a consistent backup, stop Relay first. These commands create timestamped files owned by the current host user with restrictive permissions, and mount the database volume read-only while archiving it:

```bash
(
  set -eu
  umask 077
  install -d -m 700 backups
  stamp="$(date -u +%Y%m%dT%H%M%SZ)"
  docker compose stop relay
  trap 'docker compose start relay >/dev/null' EXIT
  docker run --rm \
    -v relay-calendar-data:/data:ro \
    alpine tar -czf - -C /data . > "backups/relay-data-${stamp}.tgz"
  cp .env.production "backups/relay-env-${stamp}"
  test -s "backups/relay-data-${stamp}.tgz"
  test -s "backups/relay-env-${stamp}"
  sha256sum "backups/relay-data-${stamp}.tgz" "backups/relay-env-${stamp}" > "backups/relay-${stamp}.sha256"
  docker compose start relay
  trap - EXIT
  printf 'Backup completed: %s\n' "$stamp"
)
```

Copy the two backup files and their `.sha256` manifest together to the protected backup system. The subshell restarts Relay if an archive or copy step fails.

### Restore a Docker backup

Restoring replaces the current Relay database. Verify the timestamp and volume name before running these commands. The matching `relay-data-...tgz`, `relay-env-...`, and checksum manifest must come from the same backup run. The checksum is verified before the current volume is removed.

```bash
(
  set -eu
  stamp=YYYYMMDDTHHMMSSZ
  test -s "backups/relay-data-${stamp}.tgz"
  test -s "backups/relay-env-${stamp}"
  test -s "backups/relay-${stamp}.sha256"
  sha256sum -c "backups/relay-${stamp}.sha256"
  docker compose down
  docker volume rm relay-calendar-data
  docker volume create relay-calendar-data
  docker run --rm \
    -v relay-calendar-data:/data \
    -v "$PWD/backups:/backup:ro" \
    alpine sh -c "tar -xzf /backup/relay-data-${stamp}.tgz -C /data && chown -R 10001:10001 /data && find /data -type d -exec chmod 700 {} \\; && find /data -type f -exec chmod 600 {} \\;"
  install -m 600 "backups/relay-env-${stamp}" .env.production
  docker compose up -d
  docker compose ps
  curl http://127.0.0.1:3000/api/health
)
```

The container runs as UID/GID `10001:10001`. A fresh Docker-managed volume inherits the correct ownership from the image. If a restored, copied, or manually created volume produces a database permission error, repair it while Relay is stopped:

```bash
docker compose stop relay
docker run --rm -v relay-calendar-data:/data alpine sh -c 'chown -R 10001:10001 /data && find /data -type d -exec chmod 700 {} \; && find /data -type f -exec chmod 600 {} \;'
docker compose start relay
```

Protect backups with the same controls as the production server. Losing `CONFIG_ENCRYPTION_KEY` makes stored Schoolbox, service-account, and OAuth secrets unrecoverable.

## Operations

- Health: `GET /api/health`
- Container status: `docker compose ps`
- Application and scheduler logs: `docker compose logs -f --tail=200 relay`
- Stop: `docker compose down` (the named database volume remains)
- Upgrade: back up first, replace the source, run `docker compose build --pull`, then `docker compose up -d`
- Change local password: **IT access > Local administrator password**
- Revoke a staff session immediately: disable, change, or remove that staff entry; active sessions are deleted

If the local administrator password is lost, recover it only from an interactive server console. Stop Relay first so no session can be created during recovery; the command requires the administrator username, writes an audit entry, and revokes every active session:

```bash
# Docker
docker compose stop relay
docker compose run --rm relay node scripts/reset-admin-password.mjs
docker compose start relay

# Native Node (while the process supervisor is stopped)
npm run auth:reset-password
```

Interrupted syncs without a heartbeat for five minutes are marked failed automatically so the scheduler can recover after a server restart. Manual and scheduled runs share a database lock, preventing concurrent organisation-wide syncs.

## Security checklist

- Keep Relay internal-only and firewall it to the IT network/VPN.
- Use HTTPS with a trusted certificate and keep `APP_ORIGIN` exactly aligned with the browser URL.
- Keep `.env.production`, the SQLite volume, backups, Schoolbox JWT, service-account key, and OAuth client secret out of source control.
- Store the local administrator password in the IT password vault and use Google sign-in for daily administration.
- Assign Viewer or Operator unless configuration access is required.
- Use a dedicated Schoolbox service user and Google service account.
- Review the Runs screen after configuration or credential changes.
- Keep the host OS, Node/Docker, reverse proxy, and Relay dependencies patched.

## Official references

- [Schoolbox API](https://api.schoolbox.com.au/)
- [Google Workspace Domain-Wide Delegation](https://developers.google.com/identity/protocols/oauth2/service-account#delegatingauthority)
- [Google Calendar authorization scopes](https://developers.google.com/workspace/calendar/api/auth)
- [Google Calendar event resource](https://developers.google.com/workspace/calendar/api/v3/reference/events)
- [Google Calendar reminders](https://developers.google.com/workspace/calendar/api/concepts/reminders)
- [Google Directory users.list](https://developers.google.com/workspace/admin/directory/reference/rest/v1/users/list)
- [Google OpenID Connect](https://developers.google.com/identity/openid-connect/openid-connect)
- [Google OAuth web server applications](https://developers.google.com/identity/protocols/oauth2/web-server)
- [Next.js self-hosting](https://nextjs.org/docs/app/guides/self-hosting)
