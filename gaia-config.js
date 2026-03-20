// ═══════════════════════════════════════════════════════════════════════════
//  Gaia Configuration File
//
//  Copy this file to your deployment and fill in the values below.
//  Do NOT commit secrets — if your org is concerned about the Client ID
//  being visible, use a private GitHub repo.
//
//  ArcGIS Online Setup:
//  ─────────────────────────────────────────────────────────────────────────
//  1. Sign in to ArcGIS Online as an org admin
//  2. Go to: Organisation → Settings → App Registration (or use ArcGIS Developer Dashboard)
//  3. Register a new application:
//       - Type:         Browser Application (implicit grant)
//       - Redirect URI: https://YOUR-ORG.github.io/gaia/
//                       (must match EXACTLY — trailing slash matters)
//  4. Copy the Client ID shown after registration
//  5. Paste it below as  agol.clientId
//  6. Set agol.portalUrl to your organisation's portal URL, e.g.:
//       - Public AGOL:   https://www.arcgis.com
//       - Enterprise:    https://gis.your-org.com/portal
//  ─────────────────────────────────────────────────────────────────────────
//
//  SAML Note:
//  If your organisation uses SAML/SSO, no extra configuration is needed here.
//  ArcGIS Online handles the SAML redirect automatically — when the user
//  clicks "Sign In", AGOL detects your org's IdP and redirects to your
//  corporate login page. Gaia never sees the SAML credentials.
// ═══════════════════════════════════════════════════════════════════════════

const GAIA_CONFIG = {

  agol: {
    // Required — paste your registered app's Client ID here
    clientId: 'YOUR_CLIENT_ID',

    // Your ArcGIS Online organisation portal URL
    // Typical values:
    //   Public AGOL:  'https://www.arcgis.com'
    //   Enterprise:   'https://gis.your-org.com/portal'
    portalUrl: 'https://www.arcgis.com',

    // Redirect URI — must exactly match what you registered in AGOL
    // Usually your GitHub Pages URL with a trailing slash:
    //   'https://your-org.github.io/gaia/'
    // Leave as-is to auto-detect from the current page URL:
    redirectUri: window.location.href.split('?')[0].split('#')[0],
  },

  // Optional — defaults shown below
  defaults: {
    basemap:    'cartodb-light',   // default basemap key
    crs:        'EPSG:4326',       // default coordinate display CRS
    cataloguePath: './catalogue.csv',
  },
};
