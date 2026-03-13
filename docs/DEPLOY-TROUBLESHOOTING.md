# Deploy to Azure – Troubleshooting

## Error: "Failed to upload blob to storage account" / "Server failed to authenticate the request"

This happens when the **Function App’s storage account** (used by Flex Consumption for deployment and runtime) rejects the request. Fix the storage configuration in Azure; the GitHub Action itself is usually fine.

### 1. Check how the Function App connects to storage

In **Azure Portal**: Function App → **Configuration** → **Application settings**.

Check **both** of these if present:

- **`AzureWebJobsStorage`** – runtime (triggers, state). Must be a valid connection string or use managed identity.
- **`DEPLOYMENT_STORAGE_CONNECTION_STRING`** – used by the platform for zip deploy / Kudu. If this is set, it **must** be a valid connection string with a working key; otherwise deploy can fail with "Failed to upload blob".

Then:

- **If you see `AzureWebJobsStorage`** (full connection string):
  - Open the **Storage account** used in that connection string.
  - Go to **Security + networking** → **Access keys**.
  - Regenerate Key1 (or Key2) if needed, then in the Function App **Configuration** update `AzureWebJobsStorage` with the new connection string (same format, new key).
  - Save and restart the Function App, then redeploy from GitHub.
  - Do the same for **`DEPLOYMENT_STORAGE_CONNECTION_STRING`** if it exists: use the storage account’s **Access keys** and set the app setting to a full connection string with a valid key.

- **If you see `AzureWebJobsStorage__accountName`** (managed identity):
  - The app uses **managed identity** to talk to storage.
  - Function App → **Identity** → ensure **System assigned** is **On** (or note the **User assigned** identity).
  - Open the **Storage account** used by that account name.
  - Go to **Access control (IAM)** → **Add role assignment**:
    - Role: **Storage Blob Data Contributor**
    - Also add **Storage Queue Data Contributor** (Functions uses queues).
  - Assign access to the Function App’s **system-assigned** or **user-assigned** managed identity.
  - Wait a few minutes for IAM to apply, then redeploy.

### 2. Storage account firewall

- Storage account → **Networking**.
- If **Public network access** is “Disabled” or restricted by firewall:
  - Either allow **Azure services** (and any required networks), or
  - Use the recommended pattern for [secured storage with Azure Functions](https://learn.microsoft.com/en-us/azure/azure-functions/configure-networking-how-to) (e.g. VNet, service endpoint, private link) so the Function App and deployment path can reach the storage account.

### 3. Redeploy

After fixing storage and saving app settings (and restarting the app if you changed settings), run the GitHub Action again (push to `main` or **Actions** → **Deploy to Azure Function** → **Run workflow**).
