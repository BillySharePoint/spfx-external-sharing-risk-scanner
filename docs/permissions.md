# Permissions

## Overview

The External Sharing Risk Scanner uses **only delegated permissions** — it accesses data as the currently logged-in SharePoint user. It does NOT use application permissions and does NOT require admin consent.

## Required Permissions

### SharePoint REST API (Delegated)

| Endpoint | Permission Needed | Purpose |
|---|---|---|
| `/_api/web` | Site Read | Read site title, URL, web ID |
| `/_api/web/currentuser` | Site Read | Read current user info |
| `/_api/web/siteusers` | Site Read | List users visible to current user |
| `/_api/web/sitegroups` | Site Read | List SharePoint groups |
| `/_api/web/sitegroups/getbyid({id})/users` | Site Read / Group Membership visibility | List group members |
| `/_api/web/lists` | Site Read | List document libraries |
| `/_api/web/lists(guid'{id}')` | Site Read | Check library permission inheritance |
| `/_api/web` (HasUniqueRoleAssignments) | Site Read | Check site permission inheritance |

### What Permission Level Works Best?

| User Role | What They'll See |
|---|---|
| **Site Collection Administrator** | Full visibility — all users, all groups, all permissions |
| **Site Owner** | Most data — may miss some system groups |
| **Site Member** | Partial — may not see all group memberships |
| **Site Visitor** | Limited — basic site info, possibly restricted user/group visibility |

### What Happens When Permissions Are Insufficient?

The web part handles permission errors gracefully:
- If user enumeration fails → Shows a warning, continues with available data
- If group member enumeration fails → Shows "Unable to read group members" note
- If library permission check fails → Shows "Error" status for that indicator
- If site info fails → Shows a fatal error with retry option

## Microsoft Graph

**Not used.** The public MVP does not require any Microsoft Graph API permissions.

## Admin Consent

**Not required.** The solution works with standard delegated SharePoint permissions already available to site users.

## API Permission Requests

The `package-solution.json` does NOT include any `webApiPermissionRequests`. This means:
- No tenant-admin approval needed
- No Azure AD app registration needed (beyond what SPFx provides)
- Instant deployment without permission bottlenecks

## Security Model

```
User → SPFx Web Part → SharePoint REST API → Returns only data the user can already see
```

The web part NEVER:
- Elevates permissions
- Accesses data the user wouldn't normally be able to see
- Stores or transmits scan results outside the tenant
- Modifies any SharePoint data (read-only)
