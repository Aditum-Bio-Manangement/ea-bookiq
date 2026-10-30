# AB Room IQ - Outlook Add-in for Conference Room Booking

**Intelligent conference room booking for Aditum Bio employees.**

AB Room IQ is a Microsoft Outlook add-in that streamlines the process of booking conference rooms. It automatically identifies the user's office location based on their security group membership (EA-Cambridge or EA-Oakland) and shows only relevant rooms with real-time availability.

## Features

- **Auto-detect office location** - Identifies whether the organizer should default to Cambridge or Oakland room inventory based on membership in `ea-cambridge@aditumbio.com` or `ea-oakland@aditumbio.com`
- **Site-specific rooms** - Shows only rooms for the correct site, not the entire tenant room inventory
- **Real-time availability** - Uses Microsoft Graph free/busy data for current meeting start and end times
- **One-click booking** - Adds the room to the meeting and updates the location with a single click
- **Persistent sign-in** - Users stay signed in across sessions

## Architecture

```
┌─────────────────────────────────────┐
│ Outlook Appointment Compose Window  │
│                                     │
│   event-based activation / command  │
└──────────────────┬──────────────────┘
                   │
                   ▼
┌─────────────────────────────────────┐
│     Outlook Add-in Task Pane        │
│                                     │
│       MSAL delegated token          │
└──────────────────┬──────────────────┘
                   │
                   ▼
┌─────────────────────────────────────┐
│         Microsoft Graph             │
│                                     │
│  /me/calendar/getSchedule           │
│  /me/memberOf (group membership)    │
│  /places/... (room lists and rooms) │
│  /groups (EA-Cambridge, EA-Oakland) │
└──────────────────┬──────────────────┘
                   │
                   ▼
┌─────────────────────────────────────┐
│     Office.js Appointment APIs      │
│                                     │
│  - Add room as recipient            │
│  - Set meeting location             │
│  - Refresh UX state                 │
└─────────────────────────────────────┘
```

## Project Structure

```
/app
  /taskpane           # Outlook task pane page
  layout.tsx          # Root layout with Aditum Bio branding
  globals.css         # Aditum Bio color theme

/components
  /outlook-addin
    TaskPane.tsx      # Main task pane component
    RoomCard.tsx      # Room display with booking action
    RoomList.tsx      # List of available/busy rooms

/lib/outlook-addin
  /auth
    msal.ts           # MSAL authentication (NAA when supported, popup fallback)
  /config
    offices.ts        # Cambridge/Oakland office configuration
  /domain
    booking.ts        # Room booking logic
    officeResolver.ts # Office detection from group membership
    roomRanker.ts     # Room sorting by availability/capacity
  /graph
    graphClient.ts    # Microsoft Graph client
    groups.ts         # Group membership queries
    places.ts         # Room/place queries
    schedule.ts       # Free/busy availability
  /office
    appointment.ts    # Office.js appointment APIs
    eventHandlers.ts  # Office.js event handling

/docs
  auth.md             # Entra redirect URIs and NAA setup

/public
  manifest.xml        # Outlook add-in manifest
```

## Prerequisites

- Microsoft 365 tenant with Exchange Online
- Azure AD app registration with the following permissions:
  - `User.Read`
  - `Calendars.Read.Shared`
  - `Place.Read.All`
  - `GroupMember.Read.All`
- Mail-enabled security groups: `ea-cambridge@aditumbio.com` and `ea-oakland@aditumbio.com`
- Conference rooms with naming convention: `Room Name - Location` (e.g., "Board Room - Cambridge")

## Hosting

- **Azure Static Web Apps (target host):** https://zealous-tree-04263bb0f.5.azurestaticapps.net
- **Planned custom domain:** https://roomiq.aditumbio.com (DNS is handled separately)
- **Render (fallback until cutover):** https://ea-bookiq.onrender.com remains the live host until Azure is cut over. Do not remove or change the Render service until that cutover is complete.

The app is a static Next.js export (`output: 'export'`) so Azure Static Web Apps can serve `/` and `/taskpane` as static files. `NEXT_PUBLIC_REDIRECT_URI` is required at build time. A missing value fails the build with a clear error.

## Setup

### 1. Clone and Install

```bash
git clone <repository-url>
cd ea-roomiq
npm install
```

### 2. Configure Environment Variables

Set these public build-time variables before `npm run build`:

```bash
# Azure AD App Registration
NEXT_PUBLIC_AZURE_CLIENT_ID=f6faff63-5957-497a-bb6f-7079d9a8e8f5
NEXT_PUBLIC_AZURE_TENANT_ID=672c7181-708e-4c64-9b93-4bc263e0a2e5
NEXT_PUBLIC_REDIRECT_URI=https://zealous-tree-04263bb0f.5.azurestaticapps.net/taskpane
```

For local development, point `NEXT_PUBLIC_REDIRECT_URI` at `http://localhost:3000/taskpane` and add that URI to the app registration SPA redirect list.

GitHub Actions reads the same values from repository variables or secrets:

- `AZURE_STATIC_WEB_APPS_API_TOKEN` (Actions secret)
- `NEXT_PUBLIC_AZURE_CLIENT_ID` (variable or secret)
- `NEXT_PUBLIC_AZURE_TENANT_ID` (variable or secret)
- `NEXT_PUBLIC_REDIRECT_URI` (variable or secret, Azure task pane URL)

### 3. Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com) > **Microsoft Entra ID** > **App registrations**
2. Create a new registration or use an existing one
3. Under **Authentication**:
   - Add a **Single-page application** platform
   - Add redirect URI: `https://roomiq.aditumbio.com/taskpane` (keep this)
   - Add SPA redirect URI: `brk-multihub://roomiq.aditumbio.com` (required for Nested App Auth)
   - Implicit grant Access tokens and ID tokens boxes can stay off (MSAL uses auth code + PKCE)
   - See `docs/auth.md` for the full Entra checklist
4. Under **API permissions**, add:
   - `User.Read`
   - `Calendars.Read.Shared`
   - `Place.Read.All`
   - `GroupMember.Read.All`
5. Grant admin consent for the permissions

### 4. Deploy

Azure Static Web Apps is the target host. The GitHub Actions workflow in `.github/workflows/azure-static-web-apps.yml` builds the static export and deploys it. Render remains the fallback host until cutover.

```bash
NEXT_PUBLIC_AZURE_CLIENT_ID=... \
NEXT_PUBLIC_AZURE_TENANT_ID=... \
NEXT_PUBLIC_REDIRECT_URI=https://zealous-tree-04263bb0f.5.azurestaticapps.net/taskpane \
npm run build
```

### 5. Install the Add-in

#### Via Microsoft 365 Admin Center (Organization-wide)

1. Go to [Microsoft 365 Admin Center](https://admin.microsoft.com)
2. Navigate to **Settings** > **Integrated apps** > **Upload custom apps**
3. Select **Office Add-in** and upload `manifest.xml` or provide the URL

#### Via Outlook on the Web (Personal)

1. Go to [Outlook on the Web](https://outlook.office.com)
2. Create a new meeting
3. Click **...** > **Get Add-ins** > **My add-ins** > **Add a custom add-in** > **Add from URL**
4. Enter: `https://zealous-tree-04263bb0f.5.azurestaticapps.net/manifest.xml`

## Development

```bash
# Start development server
npm run dev

# Unit tests for MSAL / NAA auth helpers
npm test

# Build the static export (requires NEXT_PUBLIC_* vars)
npm run build

# Serve the static export locally (also used if Render is redeployed from this branch)
npm start
```

## Configuration

### Adding New Office Locations

Edit `/lib/outlook-addin/config/offices.ts`:

```typescript
export const OFFICE_CONFIGS: Record<string, OfficeConfig> = {
  cambridge: {
    id: "cambridge",
    name: "Cambridge",
    displayName: "Cambridge Office",
    securityGroupEmail: "ea-cambridge@aditumbio.com",
    building: "Cambridge",
  },
  // Add new office here
  newoffice: {
    id: "newoffice",
    name: "New Office",
    displayName: "New Office Location",
    securityGroupEmail: "ea-newoffice@aditumbio.com",
    building: "New Office",
  },
};
```

### Room Naming Convention

Rooms are filtered by display name suffix matching the office location:
- `Board Room - Cambridge` → matches Cambridge office
- `Broadway Room - Oakland` → matches Oakland office

## Troubleshooting

### "Invalid Reply Address" Error

Ensure your Azure AD app registration has the correct redirect URI configured as a **Single-page application** platform, not Web.

### Rooms Not Loading

1. Check that the user is a member of `ea-cambridge@aditumbio.com` or `ea-oakland@aditumbio.com`
2. Verify the Microsoft Graph permissions are granted
3. Ensure rooms follow the naming convention `Room Name - Location`

### Sign-in Issues

1. Confirm Entra has the SPA redirects in `docs/auth.md` (`https://roomiq.aditumbio.com/taskpane` and `brk-multihub://roomiq.aditumbio.com`)
2. Verify the Azure AD app has the correct permissions with admin consent
3. If silent sign-in fails, use the Sign In button. Do not expect a popup to open on its own during room loads.

## License

Proprietary - Aditum Bio

## Support

For issues or questions, contact your IT administrator.
