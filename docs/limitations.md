# Known Limitations

## Public MVP Scope

This is a **basic governance visibility tool** designed to surface high-level risk indicators. It is intentionally limited in scope.

## Detection Limitations

### External User Detection

- Detection relies on **heuristic signals** — login name patterns (`#EXT#`, guest URNs) and email domain comparison
- **False positives** are possible: some internal service accounts may match external patterns
- **False negatives** are possible: some external users may not follow standard Azure AD guest patterns
- If no internal domains are configured in the property pane, detection is limited to login name patterns only
- External users added via Microsoft 365 Groups (Teams) may not appear in site user lists depending on access patterns

### Permission Inheritance

- Only checks whether objects **have** unique permissions — does not enumerate **what** those unique permissions are
- Limited to the first 20 document libraries (to prevent performance issues)
- Does not check folder-level or item-level permissions (too expensive for an in-page scan)
- "Unique permissions" doesn't necessarily mean "insecure" — it just means permissions are managed separately

### Risk Score

- The risk score is a **simplified numeric indicator**, not a compliance rating
- Scoring weights are transparent and deterministic but may not reflect your organization's specific risk tolerance
- A "Low" score does NOT mean the site is fully secure
- A "High" score does NOT mean there is an active security breach

## Scope Limitations

### What This Tool Does NOT Do

- ❌ Scan the entire Microsoft 365 tenant
- ❌ Enumerate sharing links at the file/item level
- ❌ Detect anonymous sharing links
- ❌ Check tenant-level sharing policies
- ❌ Perform deep recursive item-level permission scans
- ❌ Store historical scan results
- ❌ Schedule automated scans
- ❌ Send notifications or alerts
- ❌ Export results to CSV/Excel
- ❌ Compare multiple sites
- ❌ Provide automated remediation
- ❌ Replace Microsoft Purview, Defender, or audit logs
- ❌ Replace formal security/compliance audits

### Performance

- Designed for typical team sites (< 1000 users, < 50 libraries)
- May be slow on very large sites with many groups
- Fetches data on every scan — no caching between page loads
- Limited to 20 document libraries for permission checks

### Environment

- SharePoint Online only — does not support SharePoint on-premises
- Requires a modern page — will not work on classic pages
- Requires JavaScript to be enabled in the browser

## What Should Be Used Instead For…

| Need | Recommended Tool |
|---|---|
| Tenant-wide sharing audit | SharePoint Admin Center, Microsoft Purview |
| Sharing link analysis | Microsoft 365 Admin Center, Graph API reports |
| Compliance monitoring | Microsoft Purview Compliance Portal |
| Security incident detection | Microsoft Defender for Cloud Apps |
| Audit log analysis | Unified Audit Log, Purview Audit |
| Automated governance | Azure Automation, Power Automate, third-party tools |
| File-level permission analysis | SharePoint admin reports, PnP PowerShell scripts |
