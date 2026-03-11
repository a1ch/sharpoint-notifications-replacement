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
3. Add one row per subscription:
   - **Title**: full URL to the list or library (e.g. from “Copy link” on the list or library).
   - **Email**: address to receive the daily digest for that list/library.

## Timer and time zone

The function uses a timer schedule: `0 0 8 * * *` (8:00 AM every day in **UTC** by default).

To use your local time (e.g. 8:00 AM Eastern):

- In Azure: **Function App** → **Configuration** → **Application settings** → add **WEBSITE_TIME_ZONE** = `Eastern Standard Time` (or your [Windows time zone id](https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/default-time-zones)).

## Local development

1. Install [Azure Functions Core Tools](https://learn.microsoft.com/en-us/azure/azure-functions/functions-run-local) and ensure the storage emulator is running (or set `AzureWebJobsStorage` to a real storage connection string).
2. Copy `local.settings.json` and fill in the values (do not commit real secrets).
3. Run: `func start` from the project folder.

## Deploy to Azure from GitHub

This repo is set up to deploy to [GitHub - a1ch/sharpoint-notifications-replacement](https://github.com/a1ch/sharpoint-notifications-replacement).

1. **Push this code** to that repository (e.g. clone, copy files, push to `main`).

2. **Get a publish profile** from Azure:
   - **Function App** → **Overview** → **Download publish profile**.

3. **Add GitHub secrets** for the repo:
   - `AZURE_FUNCTIONAPP_PUBLISH_PROFILE` – paste the full contents of the downloaded publish profile.

4. **Set the function app name** in the workflow:
   - Edit `.github/workflows/deploy-azure-function.yml`.
   - Replace `YOUR_FUNCTION_APP_NAME` with your Azure Function App name (in the `env` section).

5. **Application settings** (including `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`, `CONFIG_SITE_URL`, `CONFIG_LIST_NAME`, `SEND_FROM_USER_ID`, and `AzureWebJobsStorage`) must be configured in the Azure Function App; they are not stored in GitHub.

After that, pushes to `main` will build and deploy the function. You can also run the workflow manually (**Actions** → **Deploy to Azure Function** → **Run workflow**).

## License

MIT.
