# Create the Function App + deploy from GitHub

This repo builds **.NET 8 isolated** Azure Functions. Use this checklist when you add a **new** Function App and connect **GitHub Actions**.

## 1. Create the Function App (Azure portal)

1. **Azure Portal** → **Create a resource** → **Function App**.
2. **Basics**
   - **Name:** choose a globally unique name (e.g. `streamflo-sharepoint-digest`). You will use this as `AZURE_FUNCTIONAPP_NAME`.
   - **Publish:** Code  
   - **Runtime stack:** .NET  
   - **Version:** 8 (LTS)  
   - **Region:** your choice (often same as storage).
3. **Hosting**
   - **Storage:** create new or select your existing storage account.
   - **Plan:** Consumption, Flex Consumption, or Premium—match what your org uses.
4. Create the app. Wait until it finishes.

## 2. Configure runtime (Application settings)

In **Function App** → **Configuration** → **Application settings**, add at least (see README for full list):

| Setting | Example / note |
|---------|----------------|
| `AZURE_TENANT_ID` | Entra tenant ID |
| `AZURE_CLIENT_ID` | App registration client ID (Graph + SharePoint access) |
| `AZURE_CLIENT_SECRET` | Client secret |
| `CONFIG_SITE_URL` | `https://streamflogroup.sharepoint.com/itsp` |
| `CONFIG_LIST_NAME` | `Digest Subscriptions` |
| `SEND_FROM_USER_ID` | Sender UPN or object ID |
| `WEBSITE_TIME_ZONE` | e.g. `Mountain Standard Time` |
| `FUNCTIONS_WORKER_RUNTIME` | `dotnet-isolated` |
| `WEBSITE_USE_PLACEHOLDER_DOTNETISOLATED` | `1` |

Confirm **AzureWebJobsStorage** (and **DEPLOYMENT_STORAGE_CONNECTION_STRING** if your plan uses it) are valid—see [docs/DEPLOY-TROUBLESHOOTING.md](./DEPLOY-TROUBLESHOOTING.md).

**Save** and **Restart** the Function App if prompted.

## 3. GitHub repository variable (Function App name)

1. GitHub repo → **Settings** → **Secrets and variables** → **Actions** → **Variables** tab.
2. **New repository variable**
   - **Name:** `AZURE_FUNCTIONAPP_NAME`  
   - **Value:** the **exact** Function App name from Azure (not the full URL).

**Push to `main`** runs **main_streamflo-sharepoint-digest.yml** (Streamflo Function App + OIDC secrets). The optional **Deploy to Azure Function** workflow is **manual-only** and uses different (`AZURE_*` / publish profile) secrets—do not point those at a tenant you no longer use.

## 4. GitHub Actions secrets (deploy identity + Azure)

Same page → **Secrets** tab. Configure:

| Secret | Purpose |
|--------|--------|
| `AZURE_TENANT_ID` | Tenant where the **subscription** and Function App live |
| `AZURE_SUBSCRIPTION_ID` | Target subscription ID |
| `AZURE_CLIENT_ID` | **Service principal (app) used only for deployment**—must have rights to deploy to the Function App (e.g. Contributor on RG) |
| `AZURE_CLIENT_SECRET` | Secret for that deploy SP |
| `AZURE_FUNCTIONAPP_PUBLISH_PROFILE` | Contents of the publish profile file (see below) |
| `AZURE_FUNCTIONAPP_RESOURCE_GROUP` | Resource group **name** containing the Function App |

**Publish profile**

1. Azure Portal → your **Function App** → **Get publish profile** (download `.PublishSettings`).
2. Open the file in a text editor, **copy the entire XML/content**.
3. Paste into GitHub secret `AZURE_FUNCTIONAPP_PUBLISH_PROFILE`.

**Deploy service principal**

Create an app registration (or use an existing automation account) and grant it **Contributor** (or narrower custom role) on the resource group or subscription. Use **that** app’s client ID + secret in `AZURE_CLIENT_ID` / `AZURE_CLIENT_SECRET` for the workflow—not necessarily the same app used for Microsoft Graph at runtime.

## 5. Run the workflow

- **Automatic:** push to `main`.
- **Manual:** **Actions** → **Deploy to Azure Function** → **Run workflow**.

The job prints **Function App: &lt;name&gt;** so you can confirm the correct target.

## 6. OIDC (optional, no client secret)

You can switch `azure/login` to **OpenID Connect** (federated credential) so you do not store `AZURE_CLIENT_SECRET`. That requires extra Entra + GitHub configuration; the current workflow uses client ID + secret for simplicity.

## Related docs

- [DEPLOY-TROUBLESHOOTING.md](./DEPLOY-TROUBLESHOOTING.md) — storage / zip deploy errors
