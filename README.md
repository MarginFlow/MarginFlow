# Margin Flow

Margin Flow is a standalone website for freight brokers who want to scan a Microsoft Outlook inbox by lane instead of opening carrier emails one by one.

## What it does

- Signs in with a Microsoft mailbox
- Loads recent Inbox emails from Microsoft Graph
- Filters emails for a lane using pickup and delivery ZIPs
- Accepts only one ZIP if that is all the lane information available
- Extracts carrier name, MC number, and quoted rate
- Hides wrong-lane emails and lane matches with margin at or below `-50%`
- Lets the broker select relevant carriers and send individual reply prices back

## Local run

From this folder:

```powershell
node .\server.js
```

Open:

```text
http://localhost:3000
```

## Vercel deployment

Margin Flow is now structured to work on Vercel:

- static app files are in `public/`
- lane parsing runs from `api/parse-lane.js`
- Microsoft redirect URI is derived from `window.location.origin`

### Deploy checklist

1. Import this GitHub repo into Vercel.
2. Set the project root to the `firefly-freight-monitor` folder if the repo contains other folders.
3. Deploy the project.
4. Test the generated Vercel URL.
5. Add the Vercel production URL to Microsoft Entra as a SPA redirect URI.
6. If using a custom domain, add that domain as another SPA redirect URI too.

Example production redirect URIs:

- `https://your-project.vercel.app/`
- `https://marginflow.com/`
- `https://www.marginflow.com/`

### Important note for Microsoft sign-in

Every live URL you use for Margin Flow must also be added in Microsoft Entra:

- `Authentication`
- `Single-page application`
- redirect URI matching the exact deployed URL

Keep localhost too for development:

- `http://localhost:3000/`

## Microsoft setup

You need a Microsoft Entra app registration before sign-in will work.

### App registration checklist

1. Open the Azure portal and go to `Microsoft Entra ID`.
2. Open `App registrations`.
3. Create a new registration named `Margin Flow`.
4. Supported account types:
   - `Accounts in any organizational directory and personal Microsoft accounts` is the most flexible option for testing.
5. Add a redirect URI:
   - Platform: `Single-page application`
   - URI: `http://localhost:3000`

### API permissions

Add these delegated Microsoft Graph permissions:

- `openid`
- `profile`
- `email`
- `User.Read`
- `Mail.Read`
- `Mail.Send`

Admin consent is usually not required for personal Microsoft accounts using delegated permissions, but tenant policy may differ for work or school accounts.

## Local config

Edit:

- `public/config.js`

Replace:

- `clientId` with your Microsoft app registration Application (client) ID
- `tenantId` with `common` for mixed personal/work testing, or your tenant ID if you want to lock it down

Example:

```js
window.FIREFLY_CONFIG = {
  clientId: "11111111-2222-3333-4444-555555555555",
  tenantId: "common",
  redirectUri: `${window.location.origin}/`,
  graphScopes: ["openid", "profile", "email", "User.Read", "Mail.Read", "Mail.Send"]
};
```

## Current workflow

1. Sign in with Microsoft
2. Enter pickup ZIP, dropoff ZIP, and shipper rate
3. Refresh inbox if needed
4. Scan the lane
5. Review the ranked carrier checklist
6. Type a counter price or reply text for any carrier
7. Send replies individually from the site

If the reply box contains only a number such as `1800`, Margin Flow automatically sends:

```text
Can you do $1800 all in?
```

If the reply box contains a full message, Margin Flow sends it as entered.

## Notes

- This is a local-first MVP. It uses browser sign-in and direct Microsoft Graph calls from the frontend.
- The server-side lane parsing endpoint exists at `/api/parse-lane`, but Microsoft authentication is currently browser-based rather than server-based.
- Old Outlook add-in experiments still exist elsewhere in the workspace, but Margin Flow is isolated in its own folder and does not depend on them.
