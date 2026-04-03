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

## 3. GitHub: deploy workflow and OIDC

**Push to `main`** (and manual **Run workflow**) use **main_streamflo-sharepoint-digest.yml**. It deploys **`streamflo-sharepoint-digest`** (change the name in that YAML if you fork). Login uses **OIDC**: GitHub secrets named like **`AZUREAPPSERVICE_*`** (tenant, subscription, client ID), usually created from Azure Portal → Function App → **Deployment Center** → **GitHub Actions**. See [Connect GitHub to Azure](https://learn.microsoft.com/en-us/azure/developer/github/connect-from-azure).

You can delete unused legacy GitHub secrets (`AZURE_FUNCTIONAPP_PUBLISH_PROFILE`, deploy-only SP secrets) if you added them for the old manual workflow.

## 4. Run the workflow

- **Automatic:** push to `main`.
- **Manual:** **Actions** → **Build and deploy dotnet core project to Azure Function App - streamflo-sharepoint-digest** → **Run workflow**.

## Related docs

- [DEPLOY-TROUBLESHOOTING.md](./DEPLOY-TROUBLESHOOTING.md) — storage / zip deploy errors
