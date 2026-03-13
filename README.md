# SharePoint Daily Digest

Azure Function that runs once a day at 8:00 AM and emails users about new or changed items in SharePoint lists/libraries they’re subscribed to.

## How it works

1. **Config list** – A SharePoint list (on a site you specify) has two columns:
   - **Title** – Full URL to a list or library (e.g. `https://tenant.sharepoint.com/sites/MySite/Lists/MyList/AllItems.aspx`).
   - **Email** – Email address of the person who should receive the digest for that list/library.

2. **Daily run** – At 8:00 AM (see [Timer](#timer-and-time-zone)), the function:
   - Reads all rows from the config list.
   - For each row, gets items from that list/library that were **modified in the last 24 hours**.
   - Sends one digest email per row (only if there are changes) with links to the changed items.

## App registration (Microsoft Entra)

Create an app registration in your tenant and grant these **application** permissions (admin consent required):

| Permission   | Use                          |
|-------------|------------------------------|
| `Sites.Read.All` | Read SharePoint sites, lists, and list items. |
| `Mail.Send`      | Send email from the specified mailbox.        |

Then create a **client secret** and note:

- **Application (client) ID**
- **Directory (tenant) ID**
- **Client secret value**

### Sending email

Mail is sent via Microsoft Graph. You must send **from** a user or shared mailbox. Set **SEND_FROM_USER_ID** to that mailbox’s **Object ID** or **User Principal Name**.  
That user/mailbox must have **Mail.Send** granted to the app (application permission with admin consent is enough).

## Configuration

Configure the following in **Azure Function App** → **Configuration** → **Application settings** (and in `local.settings.json` for local runs).

| Setting               | Description |
|-----------------------|-------------|
| `AZURE_TENANT_ID`     | Directory (tenant) ID of your app registration. |
| `AZURE_CLIENT_ID`     | Application (client) ID. |
| `AZURE_CLIENT_SECRET` | Client secret value. |
| `CONFIG_SITE_URL`     | SharePoint site that contains the config list (e.g. `https://tenant.sharepoint.com/sites/MySite`). |
| `CONFIG_LIST_NAME`    | Display name of the config list (default: `Digest Subscriptions`). |
| `SEND_FROM_USER_ID`   | Object ID or UPN of the mailbox to send digest emails from. |

`AzureWebJobsStorage` must point to your storage account connection string (required for the function app).

## Config list in SharePoint

1. On the site specified by `CONFIG_SITE_URL`, create a list (e.g. **Digest Subscriptions**).
2. Add two single-line text columns (if not already present):
   - **Title** – used for the list/library URL.
   - **Email** – used for the recipient email.
   - **Brand** (optional) – one of **Streamflo**, **Masterflo**, **Dycor** to style the digest email (case-insensitive).
3. Add one row per subscription:
   - **Title**: full URL to the list or library (e.g. from “Copy link” on the list or library).
   - **Email**: address to receive the daily digest for that list/library.
   - **Brand**: Streamflo, Masterflo, or Dycor (optional).

## Timer and time zone

The function uses a timer schedule: `0 0 8 * * *` (8:00 AM every day in **UTC** by default).

To use your local time (e.g. 8:00 AM Eastern):

- In Azure: **Function App** → **Configuration** → **Application settings** → add **WEBSITE_TIME_ZONE** = `Eastern Standard Time` (or your [Windows time zone id](https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/default-time-zones)).

## Local development

Run the function locally to reproduce and debug startup/runtime errors (e.g. `WorkerProcessExitException`). You’ll see worker output and any `[Worker startup failed]` messages in the same console.

1. **.NET 8 SDK** – Ensure the [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0) is installed (`dotnet --version`).
2. **Azure Functions Core Tools v4** – Install one of:
   - **npm:** `npm install -g azure-functions-core-tools@4`
   - **winget:** `winget install Microsoft.Azure.FunctionsCoreTools`
   - Or [MSI](https://learn.microsoft.com/en-us/azure/azure-functions/functions-run-local#install-the-azure-functions-core-tools).
3. **Local settings** – Copy the template and add your values (do not commit real secrets):
   ```bash
   copy local.settings.json.example local.settings.json
   ```
   Edit `local.settings.json`: set `AzureWebJobsStorage` to a real Azure Storage connection string (or run [Azurite](https://learn.microsoft.com/en-us/azure/storage/common/storage-use-azurite) and keep `UseDevelopmentStorage=true`). Fill in `AZURE_*`, `CONFIG_SITE_URL`, `SEND_FROM_USER_ID` for full runs.  
   **Note:** `local.settings.json` is in `.gitignore` and must never be committed. Optional: copy `.githooks/pre-commit` to `.git/hooks/pre-commit` to block committing it by mistake.
4. **Build and run** – From the project folder:
   ```bash
   dotnet build
   func start
   ```
   Or use the script: `.\run-local.ps1` (PowerShell). The timer runs on its schedule; to trigger immediately you can use the **Run** button in the Azure Functions CLI output or call the function from the portal later.

   **"The listener for function 'DailyDigest' was unable to start" / connection refused to 127.0.0.1:10000:** The host needs storage for the timer trigger. Either (a) set `AzureWebJobsStorage` in `local.settings.json` to a **real Azure Storage connection string** (from Azure Portal → your storage account → Access keys), or (b) run the storage emulator [Azurite](https://learn.microsoft.com/en-us/azure/storage/common/storage-use-azurite) (e.g. `docker run -p 10000:10000 -p 10001:10001 -p 10002:10002 mcr.microsoft.com/azure-storage/azurite` or `npx azurite --silent --location c:\azurite --debug c:\azurite\debug.log`) so `UseDevelopmentStorage=true` can connect.

## Deploy to Azure from GitHub

This repo is set up to deploy to [GitHub - a1ch/sharpoint-notifications-replacement](https://github.com/a1ch/sharpoint-notifications-replacement).

1. **Push this code** to that repository (e.g. clone, copy files, push to `main`).

2. **Get a publish profile** from Azure:
   - **Function App** → **Overview** → **Download publish profile**.

3. **Add GitHub secrets** for the repo (deploy uses RBAC and needs Azure login):
   - `AZURE_FUNCTIONAPP_PUBLISH_PROFILE` – (optional) paste the full contents of the downloaded publish profile.
   - For RBAC deploy, also add: `AZURE_CLIENT_ID`, `AZURE_TENANT_ID`, `AZURE_SUBSCRIPTION_ID`, `AZURE_CLIENT_SECRET` (from an app registration that has **Contributor** or **Website Contributor** on the Function App or its resource group).
   - Optional: `AZURE_FUNCTIONAPP_RESOURCE_GROUP` – if set, the workflow will ensure `FUNCTIONS_WORKER_RUNTIME=dotnet-isolated` and `WEBSITE_USE_PLACEHOLDER_DOTNETISOLATED=1` on the Function App after each deploy (helps avoid worker startup failures).

4. **Set the function app name** in the workflow:
   - Edit `.github/workflows/deploy-azure-function.yml`.
   - Replace `YOUR_FUNCTION_APP_NAME` with your Azure Function App name (in the `env` section).

5. **Application settings** (including `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`, `CONFIG_SITE_URL`, `CONFIG_LIST_NAME`, `SEND_FROM_USER_ID`, and `AzureWebJobsStorage`) must be configured in the Azure Function App; they are not stored in GitHub.

After that, pushes to `main` will build and deploy the function. You can also run the workflow manually (**Actions** → **Deploy to Azure Function** → **Run workflow**).

**If the build fails with errors about `IDictionary`/`IReadOnlyDictionary`, `GetByPath`, or `SendMailPostRequestBody`:** the workflow may be running from a fork or an older clone. Sync with the upstream repo: in your clone run `git fetch https://github.com/a1ch/sharpoint-notifications-replacement.git main` and then `git merge FETCH_HEAD` (or reset to that commit), then push. Ensure the workflow runs from the repo that has the latest `main` (check the "Verify repo and commit" step in the Actions log).

**If the function fails to start with `System.AggregateException` / `WorkerProcessExitException` or "dotnet exited with code 150":** The isolated worker process is exiting before the host can load function metadata—usually because the Function App is not configured for .NET 8 Isolated. In the Azure Portal:

1. Open the Function App → **Configuration** → **General settings**.
2. Set **Stack** to **.NET 8** and **Platform** (or "Platform configuration") to **.NET Isolated**. Save.
3. In **Application settings**, ensure `FUNCTIONS_WORKER_RUNTIME` = `dotnet-isolated`. If you use a **Consumption** plan, add `WEBSITE_USE_PLACEHOLDER_DOTNETISOLATED` = `1`.
4. Save, **restart** the Function App, then redeploy (or run the deploy workflow again).

To see the actual worker exit reason: use **Monitoring** → **Log stream**, restart the app, and look for lines starting with `[Worker startup failed]` or `[Inner]` (the app logs the real exception to stderr). You can also check **Development** → **Advanced Tools (Kudu)** → **Debug console** → **CMD** → `LogFiles` for the latest worker logs.

**"Node.js 20 actions are deprecated" warning:** The workflow sets `FORCE_JAVASCRIPT_ACTIONS_TO_NODE24: 'true'` so actions use Node 24. If you still see the warning, start a **new** workflow run (don’t re-run an old one); re-runs use the workflow from the original commit. The warning is harmless and will disappear when GitHub switches the default in 2026.

## License

MIT.
