# AB Book IQ - Outlook Add-in for Conference Room Booking

**Intelligent conference room booking for Aditum Bio employees.**

AB Book IQ is a Microsoft Outlook add-in that streamlines the process of booking conference rooms. It automatically identifies the user's office location based on their security group membership (EA-Cambridge or EA-Oakland) and shows only relevant rooms with real-time availability.

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
    msal.ts           # MSAL authentication (popup mode)
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

## Setup

### 1. Clone and Install

```bash
git clone <repository-url>
cd ea-bookiq
pnpm install
```

### 2. Configure Environment Variables

Copy `.env.example` to `.env.local` and fill in your values:

```bash
# Azure AD App Registration
NEXT_PUBLIC_AZURE_CLIENT_ID=your-application-client-id
NEXT_PUBLIC_AZURE_TENANT_ID=your-tenant-id
NEXT_PUBLIC_REDIRECT_URI=https://your-domain.com/taskpane
```

### 3. Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com) > **Microsoft Entra ID** > **App registrations**
2. Create a new registration or use an existing one
3. Under **Authentication**:
   - Add a **Single-page application** platform
   - Add redirect URI: `https://your-domain.com/taskpane`
   - Enable **Access tokens** and **ID tokens** under Implicit grant
4. Under **API permissions**, add:
   - `User.Read`
   - `Calendars.Read.Shared`
   - `Place.Read.All`
   - `GroupMember.Read.All`
5. Grant admin consent for the permissions

### 4. Deploy

Deploy to your hosting platform (Render, Vercel, etc.):

```bash
pnpm build
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
4. Enter: `https://your-domain.com/manifest.xml`

## Development

```bash
# Start development server
pnpm dev

# Build for production
pnpm build

# Start production server
pnpm start
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

1. Clear browser cache/localStorage
2. Verify the Azure AD app has the correct permissions with admin consent
3. Check that implicit grant is enabled for access and ID tokens

## License

Proprietary - Aditum Bio

## Support

For issues or questions, contact your IT administrator.
