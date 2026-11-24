// Values from the App Registration:
// entra.microsoft.com → App registrations → <your app> → Overview
window.APP_CONFIG = {
  // "NeenOpal Export for Power BI": multi-tenant, publisher verified (the earlier single-tenant
  // test registration "Export to Excel" was 8d911d46-08a2-4fd4-ae8a-651922cbdce3)
  clientId: "873fbd6b-a9e9-4536-94a0-e17749b84b30",
  // Who can sign in:
  //  - "common" → work or school accounts from any organization (personal Microsoft accounts
  //    reach the app too, and are told Power BI needs a work account)
  //  - "organizations" → any work or school account only
  //  - a tenant ID → only that organization (and its guests)
  tenantId: "common",

  // Must exactly match a "Single-page application" redirect URI on the App Registration.
  redirectUri: "http://localhost:5500",

  productName: "NeenOpal Export for Power BI",

  // "trial" adds a "Created with …" footer to every sheet; "pro" removes it.
  // NOT ENFORCED – a client-side flag only. Phase 2 replaces it with a licence check
  // (Azure Function mapping the customer's tenant ID to trial / pro).
  edition: "trial",

  // Viewers (who can't download the report file) get the report's design and Power BI's own
  // totals from the services Power BI's website uses to draw the report. Those services are
  // internal and undocumented. true (default): used. false: Viewers' exports use only
  // Microsoft's documented APIs (correct data, theme colours, totals from the rows).
  viewerRoute: true,
};
