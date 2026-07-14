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
