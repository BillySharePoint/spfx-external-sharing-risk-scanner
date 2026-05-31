# SPFx External Sharing Risk Scanner

A read-only SharePoint Framework (SPFx) web part that helps site owners and Microsoft 365 admins identify basic external sharing and permission risk signals on a SharePoint Online site.

![SPFx 1.20+](https://img.shields.io/badge/SPFx-1.20+-green.svg)
![Node.js v22](https://img.shields.io/badge/Node.js-v22%20LTS-green.svg)
![SharePoint Online](https://img.shields.io/badge/SharePoint-Online-blue.svg)
![React 17](https://img.shields.io/badge/React-17-blue.svg)

---

## What It Does

The External Sharing Risk Scanner provides a **quick governance snapshot** of the current SharePoint site:

- **Site Summary** — Shows site title, URL, current user, and scan timestamp
- **External User Detection** — Flags possible external/guest users using multiple heuristic signals (login patterns, email domain checks)
- **SharePoint Group Review** — Lists groups with member counts and external user indicators
- **Permission Inheritance Check** — Identifies the site and document libraries with unique (broken) permissions
- **Risk Score** — Calculates a transparent 0–100 risk score with labeled categories (Low / Medium / High / Review Recommended)
- **Recommended Actions** — Provides contextual governance guidance based on findings

---

## Screenshots

> *Screenshots to be added after deployment*

---

## Prerequisites

| Requirement | Version |
|---|---|
| Node.js | v22 LTS |
| SharePoint Online | Required |
| SPFx | 1.20.0+ |
| PnP PowerShell | Latest (for test setup only) |

---

## Setup & Development

### 1. Clone the repository

```bash
git clone https://github.com/BillySharePoint/spfx-external-sharing-risk-scanner.git
cd spfx-external-sharing-risk-scanner
```

### 2. Install dependencies

```bash
npm install
```

### 3. Configure serve target

Edit `config/serve.json` and set your SharePoint site URL:

```json
{
  "initialPage": "https://your-tenant.sharepoint.com/sites/your-site/_layouts/workbench.aspx"
}
```

### 4. Local development

```bash
gulp serve
```

This opens the SharePoint Workbench where you can add the web part.

### 5. Build for production

```bash
gulp bundle --ship
gulp package-solution --ship
```

The `.sppkg` file will be generated in `sharepoint/solution/`.

---

## Deployment

1. Build the package:
   ```bash
   gulp bundle --ship
   gulp package-solution --ship
   ```

2. Upload `sharepoint/solution/spfx-external-sharing-risk-scanner.sppkg` to your tenant or site app catalog.

3. Approve the solution (no admin consent required for API permissions — the solution uses only delegated SharePoint REST calls).

4. Add the web part to any modern page: **Edit page → Add web part → Search "External Sharing Risk Scanner"**

---

## Configuration

### Property Pane Settings

| Property | Description |
|---|---|
| **Internal Domains** | Comma-separated list of your organization's email domains (e.g., `contoso.com, contoso.onmicrosoft.com`). Users with email domains NOT in this list will be flagged as possible external users. |

If no internal domains are configured, external user detection is limited to login name pattern checks only (e.g., `#EXT#`).

---

## Test Environment Setup

PnP PowerShell scripts are provided in the `scripts/` folder to set up a test site with realistic data:

```powershell
# Connect to your test site
.\scripts\Connect-TestSite.ps1

# Create test libraries, groups, and permission scenarios
.\scripts\Setup-TestEnvironment.ps1

# Verify test data is in place
.\scripts\Verify-TestData.ps1

# Clean up when done
.\scripts\Remove-TestEnvironment.ps1
```

---

## Permissions

This web part uses **only delegated permissions** already available to the current SharePoint user:

| API | Endpoint | Purpose |
|---|---|---|
| SharePoint REST | `/_api/web` | Read site properties |
| SharePoint REST | `/_api/web/currentuser` | Read current user info |
| SharePoint REST | `/_api/web/siteusers` | List visible site users |
| SharePoint REST | `/_api/web/sitegroups` | List SharePoint groups |
| SharePoint REST | `/_api/web/sitegroups/getbyid({id})/users` | List group members |
| SharePoint REST | `/_api/web/lists` | List document libraries and check permissions |

**No Microsoft Graph permissions are required.**  
**No admin consent is required.**  
**No application permissions are used.**

---

## Security

This solution is **read-only** and:

- Does NOT modify permissions, users, groups, or sharing links
- Does NOT send data outside the tenant
- Does NOT call third-party APIs
- Does NOT store scan results externally
- Does NOT log sensitive data in production mode

---

## Known Limitations

> **Important:** This is a basic visibility tool, not a full security audit.

- External user detection uses heuristic signals — results are indicators, not definitive
- Results depend on the current user's SharePoint permissions
- Does not scan the full Microsoft 365 tenant — only the current site
- Does not deeply scan every file/folder for item-level permissions
- Does not enumerate sharing links at the file level
- Does not replace Microsoft Purview, Defender, audit logs, or formal governance reviews
- Limited to the first 20 document libraries for permission checks
- Group member enumeration may fail if the current user lacks permission to view specific groups

---

## Architecture

```
src/webparts/externalSharingRiskScanner/
├── ExternalSharingRiskScannerWebPart.ts    # Web part class (PnPjs init, property pane)
├── components/
│   ├── ExternalSharingRiskScanner.tsx      # Main orchestrator component
│   ├── SiteSummaryCard.tsx                 # Site info display
│   ├── RiskSummaryCard.tsx                 # Risk score visualization
│   ├── ExternalUsersTable.tsx              # External user list
│   ├── GroupsSummaryTable.tsx              # SharePoint groups list
│   ├── PermissionIndicators.tsx            # Permission inheritance display
│   ├── RecommendedActions.tsx              # Governance recommendations
│   ├── LoadingState.tsx                    # Loading spinner
│   ├── ErrorState.tsx                      # Error display
│   └── EmptyState.tsx                      # No-data display
├── models/                                 # TypeScript interfaces
├── services/                               # Data access services
│   ├── SharePointSiteService.ts            # Site/web data
│   ├── SharePointUserService.ts            # User enumeration
│   ├── SharePointGroupService.ts           # Group enumeration
│   ├── PermissionService.ts                # Permission inheritance checks
│   ├── RiskScoringService.ts               # Risk score calculation
│   └── LoggerService.ts                    # Minimal logging
├── constants/                              # Configuration constants
├── utils/                                  # Utility functions
└── loc/                                    # Localization strings
```

---

## Disclaimer

This tool provides **basic governance indicators** for educational and administrative awareness purposes. It is not a substitute for professional security audits, Microsoft Purview compliance tools, or formal governance reviews.

The risk score is a simplified indicator based on visible signals and should not be interpreted as a definitive security rating.

Use at your own risk. The authors are not responsible for any decisions made based on scan results.

---

## Contributing

Contributions are welcome! Please open an issue first to discuss proposed changes.

---

## License

MIT

---

## Pro / Custom Version

Need advanced capabilities like tenant-wide scanning, deep file-level permission analysis, sharing link enumeration, scheduled scans, or export features? Contact us for custom governance solutions.
