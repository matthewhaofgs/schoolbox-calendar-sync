# Relay operations

This document contains production maintenance procedures for a Docker deployment. Commands assume the repository root as the working directory and the default `relay-calendar-data` volume name.

## Service status

```bash
docker compose ps
curl http://127.0.0.1:3000/api/health
docker compose logs -f --tail=200 relay
```

## Upgrade

1. Create and verify a backup.
2. Update the application source.
3. Rebuild the container and recreate the service.

```bash
docker compose build --pull
docker compose up -d
docker compose ps
curl http://127.0.0.1:3000/api/health
```

## Microsoft 365 client-secret rotation

Rotate the single-tenant Entra application credential before its configured expiry date. Consent is attached to the application and does not normally need to be granted again when only the secret changes.

1. Create a new client secret for the existing Entra application. Leave the old secret active during validation.
2. In **Settings > Connections**, enter and save the new secret. Relay automatically pauses the Microsoft target and clears its prior verification marker when target credentials change.
3. Run the Microsoft 365 connection diagnostic with a known licensed pilot mailbox. A successful probe records the replacement credential as verified; re-enable the target and run a Microsoft-only pilot sync.
4. After both checks succeed, delete the old client secret in Entra.

If the application client ID, tenant ID, redirect URI, or Graph application permissions change, update the Relay connection and repeat the interactive admin-consent flow. The exact redirect URI remains `${APP_ORIGIN}/api/auth/microsoft/admin-consent/callback`.

## Connection reconfiguration

Use the provider-specific guides under **Connections** for connection changes. Schoolbox, Google Workspace, and Microsoft 365 have independent saved, verified, and completed states. A connection change invalidates and pauses only the affected calendar target; it does not rewrite the other provider's credentials or completion state.

The safe sequence for a changed target is:

1. Open that target's connection guide and save the replacement values.
2. Run the guide's connection diagnostic against the saved version.
3. Complete the target guide and choose whether to activate delivery.
4. Run a target-only pilot sync before widening user coverage.

## Backup

The SQLite database and `.env.production` form one recovery set. A consistent backup requires Relay to be stopped while the volume is archived.

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

Transfer the archive, environment file, and checksum manifest together to protected backup storage.

## Restore

Restore replaces the current Relay database. The data archive, environment file, and checksum manifest must share the same timestamp.

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

The container runs as UID/GID `10001:10001`. Database permission errors after a restore can be repaired while Relay is stopped:

```bash
docker compose stop relay
docker run --rm -v relay-calendar-data:/data alpine sh -c 'chown -R 10001:10001 /data && find /data -type d -exec chmod 700 {} \; && find /data -type f -exec chmod 600 {} \;'
docker compose start relay
```

## Local administrator password recovery

Password recovery requires an interactive server console and a stopped Relay service. The reset command writes an audit entry and revokes all active sessions.

Docker deployment:

```bash
docker compose stop relay
docker compose run --rm relay node scripts/reset-admin-password.mjs
docker compose start relay
```

Native Node.js deployment:

```bash
npm run auth:reset-password
```

The native application process must remain stopped while the reset command runs.
