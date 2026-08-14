# Relay installation on a fresh Ubuntu server

This guide installs Relay on a new Ubuntu Server using Docker Compose and nginx. It is written for an IT administrator who is comfortable copying commands but does not routinely administer Linux.

The finished deployment has this network path:

```text
IT workstation -> HTTPS on TCP 443 -> nginx -> Relay on 127.0.0.1:3000
```

Relay is not intended to be published directly to the internet. Keep the server and its HTTPS endpoint restricted to the IT network or VPN.

## Before starting

Prepare the following information:

- The server's static IP address.
- The username and password or SSH key for an Ubuntu account with `sudo` access.
- A DNS name for Relay, such as `relay.example.edu`.
- The CIDR range of the IT network or VPN, such as `10.20.30.0/24`.
- A TLS certificate and unencrypted PEM private key for the Relay DNS name. The certificate should be trusted by IT workstations and include its intermediate certificate chain.
- A strong password of at least 14 characters for Relay's local break-glass administrator. This is a Relay application password and is unrelated to the Ubuntu or SSH password.

The DNS record must point the Relay name to the server before Google or Microsoft callback configuration is completed.

This guide assumes a supported 64-bit Ubuntu Server LTS release and a direct installation on the server. Commands beginning with `sudo` will ask for the Ubuntu account password. Linux does not display characters while a password is entered.

## 1. Connect to the server

Open PowerShell, Terminal, or another SSH client on an administrator workstation. Replace `ubuntu` and `192.0.2.10` with the server's Ubuntu username and IP address:

```bash
ssh ubuntu@192.0.2.10
```

Accept the host-key prompt only after confirming its fingerprint through a trusted source. Enter the existing Ubuntu password when prompted.

Do not change the SSH password as part of this installation.

## 2. Patch Ubuntu and install basic tools

Refresh the package list:

```bash
sudo apt update
```

Install all available Ubuntu updates:

```bash
sudo DEBIAN_FRONTEND=noninteractive apt full-upgrade -y
```

Install the tools used by this guide:

```bash
sudo apt install -y ca-certificates curl git nginx openssl ufw
```

Set the correct server time zone. Replace `Australia/Sydney` if the school uses another zone:

```bash
sudo timedatectl set-timezone Australia/Sydney
```

Confirm the time and time zone:

```bash
timedatectl
```

Restart the server so kernel and system-library updates are active:

```bash
sudo reboot
```

The SSH connection will close. Wait approximately one minute, then reconnect:

```bash
ssh ubuntu@192.0.2.10
```

## 3. Record the deployment values

The following variables reduce typing mistakes. Replace all three example values before running the commands:

```bash
RELAY_HOSTNAME="relay.example.edu"
```

```bash
IT_NETWORK="10.20.30.0/24"
```

```bash
SERVER_IP="192.0.2.10"
```

Display the values and check them carefully:

```bash
printf 'Relay name: %s\nIT network: %s\nServer IP: %s\n' "$RELAY_HOSTNAME" "$IT_NETWORK" "$SERVER_IP"
```

These variables last only for the current SSH session. Set them again if the session is disconnected before the installation is complete.

Confirm that DNS returns the expected server address:

```bash
getent hosts "$RELAY_HOSTNAME"
```

If the command returns no result or the wrong address, correct DNS before continuing.

## 4. Install Docker Engine and Docker Compose

Remove packages that can conflict with Docker's official packages. Messages saying a package is not installed are harmless on a fresh server:

```bash
for package in docker.io docker-compose docker-compose-v2 docker-doc docker-buildx podman-docker containerd runc; do sudo apt remove -y "$package"; done
```

Create the directory for repository signing keys:

```bash
sudo install -m 0755 -d /etc/apt/keyrings
```

Download Docker's official signing key:

```bash
sudo curl -fsSL https://download.docker.com/linux/ubuntu/gpg -o /etc/apt/keyrings/docker.asc
```

Make the key readable by the package manager:

```bash
sudo chmod a+r /etc/apt/keyrings/docker.asc
```

Add Docker's official Ubuntu package repository:

```bash
sudo tee /etc/apt/sources.list.d/docker.sources > /dev/null <<EOF
Types: deb
URIs: https://download.docker.com/linux/ubuntu
Suites: $(. /etc/os-release && echo "${UBUNTU_CODENAME:-$VERSION_CODENAME}")
Components: stable
Architectures: $(dpkg --print-architecture)
Signed-By: /etc/apt/keyrings/docker.asc
EOF
```

Refresh the package list:

```bash
sudo apt update
```

Install Docker Engine, Buildx, and the Compose plugin:

```bash
sudo apt install -y docker-ce docker-ce-cli containerd.io docker-buildx-plugin docker-compose-plugin
```

Start Docker now and on future boots:

```bash
sudo systemctl enable --now docker
```

Confirm Docker is running:

```bash
sudo systemctl --no-pager --full status docker
```

Press `q` if the status command opens a scrollable view.

Confirm the Compose plugin is available:

```bash
sudo docker compose version
```

Run Docker's test container:

```bash
sudo docker run --rm hello-world
```

The account is deliberately not added to the `docker` group. Membership in that group is effectively root access, so this guide consistently uses `sudo docker`.

## 5. Download Relay

Create the application directory:

```bash
sudo install -d -o root -g root -m 0755 /opt/relay
```

Clone the public Relay repository into it:

```bash
sudo git clone --branch main --single-branch https://github.com/matthewhaofgs/schoolbox-calendar-sync.git /opt/relay
```

Move into the repository:

```bash
cd /opt/relay
```

Confirm the checked-out branch and commit:

```bash
sudo git status --short --branch
```

The branch should be `main` and the working tree should have no changes.

## 6. Create the private production environment

Set the external HTTPS origin from the DNS name entered earlier:

```bash
APP_ORIGIN="https://${RELAY_HOSTNAME}"
```

Generate three independent random secrets:

```bash
CONFIG_ENCRYPTION_KEY="$(openssl rand -base64 32 | tr -d '\n')"
```

```bash
SESSION_SECRET="$(openssl rand -base64 48 | tr -d '\n')"
```

```bash
SCHEDULER_TOKEN="$(openssl rand -base64 48 | tr -d '\n')"
```

Create `.env.production` without displaying the secrets:

```bash
sudo tee /opt/relay/.env.production > /dev/null <<EOF
NODE_ENV=production
RELAY_HOST=0.0.0.0
PORT=3000
APP_ORIGIN=${APP_ORIGIN}
DATABASE_PATH=data/relay.sqlite
CONFIG_ENCRYPTION_KEY=${CONFIG_ENCRYPTION_KEY}
SESSION_SECRET=${SESSION_SECRET}
SCHEDULER_TOKEN=${SCHEDULER_TOKEN}
EOF
```

Restrict the file to the root account:

```bash
sudo chown root:root /opt/relay/.env.production
```

```bash
sudo chmod 600 /opt/relay/.env.production
```

Remove the secrets from the current shell:

```bash
unset CONFIG_ENCRYPTION_KEY SESSION_SECRET SCHEDULER_TOKEN
```

Confirm the file exists with the correct permissions. Do not display its contents:

```bash
sudo ls -l /opt/relay/.env.production
```

The expected permissions at the start of the line are `-rw-------`.

The environment file and Relay database are a matched recovery set. Back them up together and never commit, email, or paste either one into a ticket or chat.

## 7. Restrict the Relay container to the local proxy

The repository's base Compose file listens on all server interfaces to support different deployment topologies. On this single-server nginx installation, create a local override that publishes Relay only on loopback:

```bash
sudo tee /opt/relay/compose.override.yaml > /dev/null <<'EOF'
services:
  relay:
    ports: !override
      - "127.0.0.1:3000:3000"
EOF
```

Restrict the override file:

```bash
sudo chown root:root /opt/relay/compose.override.yaml
```

```bash
sudo chmod 644 /opt/relay/compose.override.yaml
```

Keep the host-specific override out of the Git working tree:

```bash
printf '/compose.override.yaml\n' | sudo tee -a /opt/relay/.git/info/exclude > /dev/null
```

Validate the merged Compose configuration without printing its environment values:

```bash
sudo docker compose config --quiet
```

The `!override` tag requires Docker Compose 2.24.4 or newer. The current Compose plugin installed from Docker's official repository satisfies this requirement.

## 8. Build Relay

Download the required base images and build the application image:

```bash
sudo docker compose build --pull
```

The first build may take several minutes. It should finish with a successful build message and no errors.

## 9. Create the local Relay administrator

Run the interactive administrator bootstrap:

```bash
sudo docker compose run --rm relay node scripts/bootstrap-admin.mjs
```

At the prompts:

1. Press Enter to use the username `administrator`, or enter another local Relay username.
2. Enter a new Relay password containing at least 14 characters.
3. Enter the same password again.

The password will not be shown. This command creates only a Relay application account; it does not change the Ubuntu or SSH password.

Store the local Relay credentials in the organisation's password manager. This account is the break-glass owner and can manage other Relay administrators.

## 10. Start Relay

Start the web application and scheduler in the background:

```bash
sudo docker compose up -d
```

Show the container state:

```bash
sudo docker compose ps
```

Wait up to 30 seconds, then test Relay directly on the server:

```bash
curl --fail --show-error http://127.0.0.1:3000/api/health
```

The expected response is:

```json
{"ok":true}
```

If the health check fails, view the latest logs:

```bash
sudo docker compose logs --tail=200 relay
```

Do not continue until the local health check succeeds.

Confirm port 3000 is listening only on loopback:

```bash
sudo ss -lntp | grep ':3000'
```

The listening address should be `127.0.0.1:3000`, not `0.0.0.0:3000` or `[::]:3000`.

## 11. Install the TLS certificate

On the administrator workstation, place these two PEM files in the current directory:

- `relay.fullchain.pem`: the server certificate followed by any intermediate certificates.
- `relay.key`: the matching unencrypted private key.

From the workstation, replace the username, server address, and local filenames as necessary, then upload the files:

```bash
scp relay.fullchain.pem relay.key ubuntu@192.0.2.10:/tmp/
```

Return to the server's SSH session and create the nginx certificate directory:

```bash
sudo install -d -o root -g root -m 0700 /etc/nginx/tls
```

Install the certificate chain:

```bash
sudo install -o root -g root -m 0644 /tmp/relay.fullchain.pem /etc/nginx/tls/relay.fullchain.pem
```

Install the private key:

```bash
sudo install -o root -g root -m 0600 /tmp/relay.key /etc/nginx/tls/relay.key
```

Delete the uploaded temporary copies:

```bash
rm -f /tmp/relay.fullchain.pem /tmp/relay.key
```

Confirm that the certificate covers the Relay DNS name:

```bash
openssl x509 -in /etc/nginx/tls/relay.fullchain.pem -noout -subject -issuer -dates -ext subjectAltName
```

Never display or copy the contents of `/etc/nginx/tls/relay.key`.

## 12. Configure nginx

Copy Relay's nginx example into the nginx site directory:

```bash
sudo cp /opt/relay/deploy/nginx-relay.conf.example /etc/nginx/sites-available/relay
```

Replace the example hostname with the real Relay DNS name:

```bash
sudo sed -i "s/relay\.school\.edu\.au/${RELAY_HOSTNAME}/g" /etc/nginx/sites-available/relay
```

Enable the Relay site:

```bash
sudo ln -sfn /etc/nginx/sites-available/relay /etc/nginx/sites-enabled/relay
```

Disable nginx's default site:

```bash
sudo rm -f /etc/nginx/sites-enabled/default
```

Check the nginx configuration:

```bash
sudo nginx -t
```

Do not continue if this command reports an error. Correct the named file and line before retrying.

Enable nginx on boot and reload it:

```bash
sudo systemctl enable --now nginx
```

```bash
sudo systemctl reload nginx
```

Test nginx locally while sending the correct hostname:

```bash
curl --fail --show-error --resolve "${RELAY_HOSTNAME}:443:127.0.0.1" "https://${RELAY_HOSTNAME}/api/health"
```

This command succeeds only if Ubuntu trusts the certificate's issuing CA. If an internal CA is not trusted by Ubuntu, use the browser test in section 14 after installing the organisation's root CA on the workstation.

## 13. Configure the host firewall

Allow existing SSH access before enabling the firewall:

```bash
sudo ufw allow OpenSSH
```

Allow HTTP from the IT network so nginx can redirect it to HTTPS:

```bash
sudo ufw allow from "$IT_NETWORK" to any port 80 proto tcp comment 'Relay HTTP from IT network'
```

Allow HTTPS from the IT network:

```bash
sudo ufw allow from "$IT_NETWORK" to any port 443 proto tcp comment 'Relay HTTPS from IT network'
```

Set the default policies:

```bash
sudo ufw default deny incoming
```

```bash
sudo ufw default allow outgoing
```

Enable the firewall:

```bash
sudo ufw enable
```

Confirm the rules:

```bash
sudo ufw status verbose
```

The `OpenSSH` rule is intentionally added before UFW is enabled to avoid locking out the active administrator. Restrict SSH to an approved management network later if that is part of the organisation's server standard.

Docker-published ports can bypass normal UFW processing. The loopback-only Compose override in section 7 is therefore an essential control, not just a convenience.

## 14. Verify access from an IT workstation

On an IT workstation connected to the allowed network or VPN, open:

```text
https://relay.example.edu
```

Replace the example name with the real Relay name. Confirm all of the following:

- The browser shows a valid trusted HTTPS connection.
- The Relay sign-in page appears.
- The local Relay administrator can sign in.
- `http://relay.example.edu` redirects to HTTPS.
- Connecting to `http://relay.example.edu:3000` fails. Port 3000 must not be reachable from another computer.

If the browser reports a certificate warning, stop and correct the certificate chain, DNS name, or workstation CA trust before entering provider credentials.

## 15. Complete the in-app connection guides

Sign in with the local Relay administrator and open **Connections**. Schoolbox, Google Workspace, and Microsoft 365 have independent setup guides. Relay becomes operational when Schoolbox and at least one calendar target are complete.

### Schoolbox

Prepare:

- The Schoolbox HTTPS base URL.
- A Schoolbox superuser JWT with user-list and delegated calendar access.

Open **Connections > Schoolbox**, save the values, run the stored-connection diagnostic, and complete the guide.

### Google Workspace target

Prepare:

- A Google Cloud service-account JSON key.
- The delegated Google Workspace administrator email.
- The Google Workspace customer ID if the default `my_customer` value is not suitable.

Enable the Admin SDK and Google Calendar API in the Google Cloud project. In Google Admin, grant the service account's numeric client ID these Domain-Wide Delegation scopes as one comma-separated line:

```text
https://www.googleapis.com/auth/calendar.events.owned,https://www.googleapis.com/auth/calendar.app.created,https://www.googleapis.com/auth/admin.directory.user.readonly
```

Open **Connections > Google Workspace**, follow its independent guide, run the diagnostic, and choose whether to activate Google delivery.

### Microsoft 365 target

Create a single-tenant Microsoft Entra app registration for **Accounts in this organizational directory only**.

For the redirect platform, choose **Web**. Do not choose **Single-page application** or **Mobile and desktop applications / public client (native)**. Leave public client flows disabled.

Add the exact redirect URI displayed by Relay. Its form is:

```text
https://relay.example.edu/api/auth/microsoft/admin-consent/callback
```

Add these Microsoft Graph **application permissions**, not delegated permissions:

```text
User.Read.All
Calendars.ReadWrite
```

`User.Read.All` allows Relay to discover directory users. `Calendars.ReadWrite` application permission allows Relay to create, update, and remove Relay-managed events and secondary calendars in user mailboxes across the tenant. Both are required for the full synchronization workflow.

Create a client secret and record its value immediately. In **Connections > Microsoft 365**, enter the tenant ID, application client ID, client secret, and a known licensed pilot mailbox. Complete admin consent with a Privileged Role Administrator or Global Administrator account, then run Relay's stored-connection diagnostic.

See the [Microsoft 365 setup and troubleshooting guide](microsoft-365.md) for the complete Entra portal checklist and Graph error guidance.

### IT staff sign-in

Relay's administrator sign-in uses a separate Google Web OAuth client. It is not the Google service account used for calendar synchronization.

1. Configure a Google OAuth consent screen with an **Internal** audience.
2. Create a **Web application** OAuth client.
3. Add the exact callback URL displayed under **IT access** as an authorised redirect URI.
4. Enter the Workspace domain, client ID, and client secret in Relay.
5. Add each approved IT staff member and assign the required Relay role.

Keep the local administrator as the break-glass account even after Google sign-in works.

## 16. Run a safe pilot

Leave both **Enable new Google users by default** and **Enable new Microsoft users by default** off during initial testing.

Run discovery, then review **People > Google Workspace** and **People > Microsoft 365** separately. Enable only one or a small number of matched test users on the intended target.

Run a manual sync for that target only. Review the run summary, per-user result, and event actions before widening coverage.

Relay can remove only its recorded managed events for a pilot user. It can also delete tracked Relay-created secondary calendars when explicitly requested; it never deletes a primary calendar through that operation.

## 17. Routine status commands

Move to the Relay directory before running Compose commands:

```bash
cd /opt/relay
```

Show service state:

```bash
sudo docker compose ps
```

Check health:

```bash
curl --fail --show-error http://127.0.0.1:3000/api/health
```

Show the latest application logs:

```bash
sudo docker compose logs --tail=200 relay
```

Follow new log messages until `Ctrl+C` is pressed:

```bash
sudo docker compose logs --follow --tail=200 relay
```

Restart Relay:

```bash
sudo docker compose restart relay
```

Stop Relay without deleting its database volume:

```bash
sudo docker compose down
```

Start it again:

```bash
sudo docker compose up -d
```

Never add `-v` to `docker compose down`; that option deletes the named database volume.

## 18. Backups and updates

Before every update, create a matched backup of the database volume and `.env.production`. Follow the [operations guide](operations.md), which contains the verified backup, restore, upgrade, and local-password recovery commands.

A normal source update, after a verified backup, uses these commands:

```bash
cd /opt/relay
```

```bash
sudo git fetch origin
```

```bash
sudo git checkout main
```

```bash
sudo git pull --ff-only origin main
```

```bash
sudo docker compose build --pull
```

```bash
sudo docker compose up -d
```

```bash
sudo docker compose ps
```

```bash
curl --fail --show-error http://127.0.0.1:3000/api/health
```

The host-specific `.env.production` and `compose.override.yaml` files must remain in place through updates.

## Troubleshooting

### Relay does not start

```bash
cd /opt/relay
```

```bash
sudo docker compose ps
```

```bash
sudo docker compose logs --tail=200 relay
```

### nginx configuration fails

```bash
sudo nginx -t
```

```bash
sudo journalctl -u nginx --no-pager --lines=100
```

### The browser cannot connect

Check DNS:

```bash
getent hosts "$RELAY_HOSTNAME"
```

Check listening ports:

```bash
sudo ss -lntp | grep -E ':(80|443|3000)'
```

Check the firewall:

```bash
sudo ufw status verbose
```

Check nginx locally:

```bash
curl --resolve "${RELAY_HOSTNAME}:443:127.0.0.1" "https://${RELAY_HOSTNAME}/api/health"
```

### A local administrator password is lost

Do not change the Ubuntu or SSH password. Use Relay's offline password-reset procedure in the [operations guide](operations.md).

## Upstream installation references

- [Install Docker Engine on Ubuntu](https://docs.docker.com/engine/install/ubuntu/)
- [Install the Docker Compose plugin](https://docs.docker.com/compose/install/linux/)
- [Docker Compose file merging and `!override`](https://docs.docker.com/reference/compose-file/merge/)
- [Docker packet filtering and UFW behavior](https://docs.docker.com/engine/network/packet-filtering-firewalls/)
