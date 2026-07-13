# Security policy

Relay handles privileged Google Workspace credentials, Schoolbox credentials,
identity data, and calendar data. Deploy it only on an internally accessible
server as described in the README.

## Reporting a vulnerability

Use this repository's **Security > Report a vulnerability** option to submit a
private report. Do not open a public issue containing credentials, user data,
calendar data, internal hostnames, or exploit details.

Include the affected version or commit, impact, reproduction steps, and any
suggested mitigation. Replace all real credentials and personal information
with clearly fictional values before attaching logs or screenshots.

## Operational incidents

If a deployed Relay instance may have been compromised, take it offline from
the network, preserve relevant logs, and rotate the Schoolbox JWT, Google
service-account key, Google OAuth client secret, local administrator password,
`CONFIG_ENCRYPTION_KEY`, `SESSION_SECRET`, and `SCHEDULER_TOKEN` as applicable.
Revoking or replacing secrets is an operator responsibility and should not wait
for a source-code vulnerability report.
