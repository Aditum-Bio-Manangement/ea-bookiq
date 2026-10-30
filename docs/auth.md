# Microsoft Entra authentication for AB Room IQ

Room IQ uses MSAL Browser with Nested App Auth (NAA) when Outlook supports the `NestedAppAuth` 1.1 requirement set. Older hosts keep the existing popup flow.

## Entra app registration (Caleb)

In Microsoft Entra ID > App registrations > the Room IQ SPA > Authentication:

1. Add a **Single-page application** redirect URI: `brk-multihub://roomiq.aditumbio.com`
   - Host only. Do not add a path.
   - This lets Outlook broker tokens so users are not prompted about once a day.
2. Keep the existing SPA redirect URI: `https://roomiq.aditumbio.com/taskpane`
3. Implicit grant and hybrid flows: the Access tokens and ID tokens boxes can stay off. Room IQ uses the authorization code flow with PKCE. The add-in also requests `offline_access` so the popup fallback can renew tokens when the host does not support NAA.

Do not add `brk-multihub://...` as a Web platform redirect. It must be SPA.

The add-in derives the broker URI at runtime as `brk-multihub://${window.location.host}`. After the custom domain cutover that host is `roomiq.aditumbio.com`. If you test against another host (localhost or the Azure Static Web Apps default hostname), add a matching `brk-multihub://<that-host>` SPA redirect as well.

## Manifest

The Outlook XML manifest is unchanged. Microsoft does not support declaring `NestedAppAuth` in the Outlook add-in manifest. Support is detected at runtime with:

```js
Office.context.requirements.isSetSupported("NestedAppAuth", "1.1")
```

`WebApplicationInfo` is not required for this NAA path. That element is for the older IdentityAPI SSO model. Adding it would also need an Application ID URI (`api://roomiq.aditumbio.com/<client-id>`) in Entra and is not needed for `createNestablePublicClientApplication`.

## Runtime behavior

- Startup calls `acquireTokenSilent`, then `ssoSilent`, with `loginHint` from `Office.context.mailbox.userProfile.emailAddress`.
- The Sign In button is shown only when silent auth returns an interaction-required error.
- Room loads and meeting-time changes never open a login popup. They return to **Sign In Required** if a token cannot be acquired silently.
- The popup, when the user clicks Sign In, uses the same `loginHint`.
- MSAL cache stays in `localStorage`. It is cleared only on explicit sign-out or an unrecoverable `invalid_grant` that silent renewal cannot recover from. Closing or blocking a popup does not wipe the cache.
