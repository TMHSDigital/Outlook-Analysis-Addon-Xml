# Outlook Analysis Add-On

This repository hosts the manifest file for the Outlook Analysis Add-On, a custom Office add-in that analyzes recent emails to provide a summary of current work activities.

## About the Add-On

The Outlook Analysis Add-On integrates directly into Outlook as a task pane add-in. It analyzes your most recent emails to help you quickly understand what you're currently working on, providing context and summaries of ongoing projects and communications.

## Repository Purpose

This repository serves as the public hosting location for the add-in's manifest XML file, enabling automatic updates through GitHub Pages without requiring manual uploads for each iteration.

## Manifest URL

The manifest is publicly accessible at:
```
https://TMHSDigital.github.io/outlook-analysis-addon/manifest.xml
```

## Installation

To install this add-in in Outlook:

1. Open Outlook (desktop or web)
2. Navigate to Add-ins management
3. Select "My add-ins" or "Custom add-ins"
4. Choose "Add from URL"
5. Enter the manifest URL above
6. Follow the prompts to complete installation

## Features

- Task pane interface accessible from message read view
- Automatic email analysis
- Work summary generation
- Integrated into Outlook's ribbon interface

## Technical Details

- **Office.js API Version**: 1.3+
- **Supported Hosts**: Outlook (Desktop, Tablet, Phone)
- **Permissions**: ReadWriteMailbox
- **Authentication**: Microsoft Identity (Azure AD)

## Development

This add-in is currently in development. The main application is hosted separately and runs on localhost during development.

### Required Scopes
- user.read
- mail.read
- profile
- openid
- offline_access

## License
None

Internal use only.