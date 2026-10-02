/* =====================================================================
   MJE Hubs - Microsoft 365 connection, used by BOTH apps
   (Safety Hub at MJE/, Service Hub at MJE/service/).

   To move the Hubs to another tenant, change these four values only
   (Setup-MJE-Hub.ps1 prints TENANT_ID, SP_HOST and SITE_PATH at the end).
   While CLIENT_ID starts with "<<" the apps run in DEMO mode.

   IDs only - this repo is public: never put register data, customer
   names or seed data in this file or anywhere in MJE/.
   ===================================================================== */
const MJE_TENANT = {
  CLIENT_ID: "b159957f-f5c2-40a9-8802-7b3dea3e6b45",   // Entra app "MJE Hub"
  TENANT_ID: "ea9e105c-2df6-47ef-a8e3-b4aa1d788423",
  SP_HOST:   "conceptcompg.sharepoint.com",
  SITE_PATH: "/sites/MJE"
};
