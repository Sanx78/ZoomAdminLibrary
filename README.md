# ZoomAdminLibrary

A basic PowerShell library for managing Zoom users and groups. Functions implemented include:
* Querying and returning user information
* Adding, changing and removing users
* Adding and removing groups
* Querying and modifying membership of groups
* Querying and setting user delegates
* Resetting user's password
* Running reports on Zoom usage, meetings, webinars and participants

Requires an API Key and Secret for authentication. See https://marketplace.zoom.us/docs/guides/auth/jwt#key-secret for more information. Call ```Set-ZoomAuthToken -apiKey <apiKey> -apiSecret <apiSecret>``` before calling any other function.
