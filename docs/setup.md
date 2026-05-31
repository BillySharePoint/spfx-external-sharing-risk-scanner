# Development Environment Setup

## Prerequisites

1. **Node.js v22 LTS** — Download from [nodejs.org](https://nodejs.org/)
2. **npm** — Comes with Node.js
3. **Gulp CLI** — Install globally:
   ```bash
   npm install gulp-cli --global
   ```
4. **Yeoman & SPFx Generator** (optional, for re-scaffolding):
   ```bash
   npm install yo @microsoft/generator-sharepoint --global
   ```
5. **PnP PowerShell** (for test environment setup):
   ```powershell
   Install-Module PnP.PowerShell -Scope CurrentUser
   ```

## Initial Setup

```bash
# Clone the repository
git clone https://github.com/BillySharePoint/spfx-external-sharing-risk-scanner.git
cd spfx-external-sharing-risk-scanner

# Install dependencies
npm install

# Verify the build works
gulp build
```

## Local Development

```bash
# Start the local workbench server
gulp serve
```

This will:
1. Build the project
2. Start a local HTTPS server on port 4321
3. Open the SharePoint Workbench (configured in `config/serve.json`)

### Serving on your SharePoint site

Edit `config/serve.json` to point to your test site:

```json
{
  "initialPage": "https://your-tenant.sharepoint.com/sites/your-site/_layouts/workbench.aspx"
}
```

Or navigate directly to `https://your-tenant.sharepoint.com/sites/your-site/_layouts/workbench.aspx` while `gulp serve` is running.

## Building

```bash
# Development build
gulp build

# Production build
gulp bundle --ship

# Create .sppkg package
gulp package-solution --ship
```

## Deploying to App Catalog

1. Build the production package:
   ```bash
   gulp bundle --ship
   gulp package-solution --ship
   ```

2. Go to your SharePoint tenant App Catalog (or site-level app catalog)

3. Upload `sharepoint/solution/spfx-external-sharing-risk-scanner.sppkg`

4. Trust the solution when prompted

5. Navigate to your target site → Site Contents → Add an App → Find "spfx-external-sharing-risk-scanner-client-side-solution"

6. Add the web part to any modern page

## Test Environment

See the `scripts/` folder for PnP PowerShell scripts that set up realistic test data.

```powershell
# Connect
.\scripts\Connect-TestSite.ps1

# Setup test data
.\scripts\Setup-TestEnvironment.ps1

# Verify
.\scripts\Verify-TestData.ps1
```

## Troubleshooting

### "Gulp is not recognized"
Install gulp-cli globally: `npm install gulp-cli --global`

### Build errors about missing types
Run `npm install` to ensure all dependencies are installed.

### Web part not appearing in workbench
- Check the browser console for errors
- Ensure you're on the correct SharePoint site
- Try clearing browser cache and local storage
- Verify the manifest JSON is valid

### Permission errors during scan
The web part requires at minimum Site Member permissions. For best results, run as a Site Owner or Site Collection Administrator.
