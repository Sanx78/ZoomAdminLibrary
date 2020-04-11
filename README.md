# ZoomAdminLibrary

A basic PowerShell library for managing Zoom users and groups. Functions implemented include:
* Querying and returning user information
* Adding, changing and removing users
* Adding and removing groups
* Querying and modifying membership of groups
* Querying and setting user delegates
* Resetting user's password
* Running reports on Zoom usage, meetings per user, and meeting participants

Requires a JWT token for authorisation. See https://marketplace.zoom.us/docs/guides/auth/jwt for more information. Call ```Set-ZoomAuthToken -Token <token>``` before calling any other function.



