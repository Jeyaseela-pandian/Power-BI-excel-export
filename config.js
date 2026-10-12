// Values from the App Registration:
// entra.microsoft.com → App registrations → <your app> → Overview
window.APP_CONFIG = {
  clientId: "8d911d46-08a2-4fd4-ae8a-651922cbdce3",
  // The tenant you SIGN IN to, not necessarily where the app is registered:
  //  - App registered in your own organization → that tenant's ID
  //  - App registered in another tenant (must be multi-tenant) → YOUR organization's tenant ID,
  //    or "organizations" to allow any work account
  tenantId: "38d5e40c-9bd6-41fb-8951-c387c97018d9", // NeenOpal

  // Must exactly match a "Single-page application" redirect URI on the App Registration.
  redirectUri: "http://localhost:5500",

  productName: "NeenOpal Export for Power BI",

  // "trial" adds a "Created with …" footer to every sheet; "pro" removes it.
  // NOT ENFORCED – a client-side flag only. Phase 2 replaces it with a licence check
  // (Azure Function mapping the customer's tenant ID to trial / pro).
  edition: "trial",
};
