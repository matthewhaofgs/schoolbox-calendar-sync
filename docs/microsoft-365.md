# Microsoft 365 setup and troubleshooting

Relay synchronizes Schoolbox calendar data to Outlook Calendar by using a confidential, single-tenant Microsoft Entra application and Microsoft Graph application permissions. This is app-only access: no end user installs Relay, signs in to it, or grants consent.

## Entra app registration checklist

Create the registration under **Microsoft Entra admin center > App registrations > New registration** with these values:

| Setting | Required value |
| --- | --- |
| Supported account types | Accounts in this organizational directory only |
| Authentication platform | **Web** |
| Redirect URI | `${APP_ORIGIN}/api/auth/microsoft/admin-consent/callback` |
| Public client flows | Disabled |

Do not choose **Single-page application** or **Mobile and desktop applications / public client (native)**. The redirect URI must match the value displayed by Relay exactly, including scheme, hostname, port, path, and trailing-slash form.

Create a client secret under **Certificates & secrets** and store the secret value in Relay. Record its expiry date in the organisation's credential-management system.

## Required Microsoft Graph permissions

Under **API permissions > Add a permission > Microsoft Graph > Application permissions**, add:

| Permission | Purpose |
| --- | --- |
| `User.Read.All` | Discover active Entra directory users and their current mail identities |
| `Calendars.ReadWrite` | Read, create, update, and delete Relay-managed events and secondary calendars in user mailboxes |

Both rows must show permission type **Application** and **Granted for your organisation**. Do not use the similarly named delegated permissions. `Calendars.ReadWrite` application permission is sufficient for Relay's organisation-wide calendar operations; Relay's per-target user switches determine which matched mailboxes are actually changed.

An Exchange Application RBAC assignment or legacy Application Access Policy can further restrict the mailboxes available to the application. If one is configured, its scope must include every mailbox Relay should synchronize and the explicit test mailbox used by the connection diagnostic.

## Relay connection workflow

1. Open **Connections > Microsoft 365**. This guide is independent of the Schoolbox and Google Workspace setup tracks.
2. Enter the directory tenant ID, application client ID, client secret, and an active licensed pilot mailbox.
3. Save the Microsoft draft.
4. Use **Grant or renew admin consent** with a Privileged Role Administrator or Global Administrator.
5. Verify that Entra shows both required application permissions as granted.
6. Run the Microsoft connection diagnostic. Relay tests application authentication, directory discovery, primary-calendar access, and secondary-calendar writes in the configured pilot mailbox. The temporary diagnostic calendar is deleted immediately.
7. Complete the Microsoft setup and enable delivery only when the diagnostic succeeds.

The Google Workspace connection and scheduler preference are not changed by this workflow. Changing the Microsoft tenant, client ID, client secret, or test mailbox invalidates only the Microsoft verification state; save and test the replacement before re-enabling that target.

## Pilot rollout

Keep **Enable newly matched Microsoft accounts** off during initial validation. Run Microsoft-only discovery, review **People > Microsoft 365**, and enable one or a small number of matched mailboxes. Run a Microsoft-only sync and inspect its target, user, and event diagnostics before widening coverage.

Google and Microsoft coverage are independent. The same Schoolbox person can be enabled for Google, Microsoft, both, or neither. Per-person Schoolbox event exclusions apply consistently to both targets, while event policy, calendar routing, cleanup, and enablement remain target-specific.

## Directory identity and address changes

Microsoft target users are stored by stable Entra object ID, not by email address. During each Microsoft directory discovery Relay refreshes the account's current display name, primary address, aliases, enabled state, and Schoolbox match.

Relay chooses the current address in this order:

1. The uppercase `SMTP:` entry in `proxyAddresses`, which Microsoft uses to denote the primary SMTP address.
2. The Entra `mail` value.
3. `userPrincipalName`.

Lowercase `smtp:` entries and the remaining mail values are retained as aliases for matching. When an address changes in Microsoft 365, allow time for Entra/Exchange directory propagation, then run a Microsoft-only sync or diagnostic discovery and refresh the People page. The stable Entra object ID preserves the account's Relay enablement and managed-event mappings while the displayed email and Schoolbox match are recalculated.

If the old address remains visible after propagation, verify the primary address in Exchange/Entra rather than relying only on the sign-in name. If the account becomes **Unmatched**, confirm that Schoolbox contains one unambiguous active primary or alternate email corresponding to the new address.

## Troubleshooting

### `Authorization_RequestDenied` or HTTP 403 during directory discovery

Confirm that `User.Read.All` was added under **Application permissions**, not Delegated permissions, and that tenant-wide admin consent shows as granted. Relay's consent action approves permissions already configured on the app registration; it cannot add a missing permission.

### HTTP 403 during mailbox or calendar access

Confirm that `Calendars.ReadWrite` is an **Application** permission with admin consent, the test user has an active Exchange Online mailbox, and any Exchange Application RBAC or Application Access Policy includes that mailbox. Retry the Relay connection diagnostic after correcting the tenant configuration.

### Consent callback or redirect mismatch

Register the exact callback shown by Relay under the Entra **Web** platform. Confirm that production `APP_ORIGIN` is the externally accessed HTTPS origin and that the scheme, hostname, port, path, and trailing-slash form match.

### Users are discovered but no events are written

Open **People > Microsoft 365** and confirm that the account is matched and its **Calendar sync** switch is enabled for Microsoft. An account enabled for Google is not automatically enabled for Microsoft. Then review the Microsoft event rules and run a Microsoft-only pilot sync.

### Throttling or timeouts

Relay honors Microsoft Graph `Retry-After` responses and applies discovery, per-account, and whole-run deadlines. Review the target and user outcomes in **Runs**. If throttling persists, lower **Settings > Advanced > Concurrent target accounts** and retry a Microsoft-only run.

## Permission and security boundaries

- App-only Graph authorization is tenant-wide unless Exchange scopes it more narrowly.
- Relay writes only to matched Microsoft accounts whose Microsoft target switch is enabled.
- Event deletion is limited to provider event IDs recorded in Relay's Microsoft mapping tables.
- Secondary-calendar deletion is limited to Microsoft calendar IDs recorded as Relay-created; primary calendars are never eligible.
- The Microsoft client secret is encrypted in the Relay database by `CONFIG_ENCRYPTION_KEY`.
- Microsoft admin consent does not grant access to the Relay administration interface; Relay administrator access remains local or Google OpenID Connect allowlist-based.

## Related documentation

- [Relay README](../README.md)
- [Production operations](operations.md)
- [Microsoft identity platform client-credentials flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-client-creds-grant-flow)
- [Microsoft Graph permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference)
- [Microsoft Graph throttling guidance](https://learn.microsoft.com/en-us/graph/throttling)
