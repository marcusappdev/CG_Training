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

/* =====================================================================
   Which Hubs are switched on.  Each one is "on", "off" or "preview".
     "on"      - everyone can use it (the default)
     "off"     - hidden from the header for everyone; opening its address
                 shows "isn't available yet"
     "preview" - only the emails in MJE_PREVIEW can see and use it;
                 everyone else is treated as "off"
   Save the change in GitHub; it shows within about 10 minutes
   (Ctrl+F5 to see it straight away).
   This hides a Hub - it is not security. SharePoint permissions on the
   MJE lists are still what controls who can read or change data.
   ===================================================================== */
const MJE_APPS_ON = {
  safety:  "on",
  service: "preview"
};
// Sign-in emails that can see a Hub set to "preview" (office staff only)
const MJE_PREVIEW = ["mscully@concept.com.pg"];
