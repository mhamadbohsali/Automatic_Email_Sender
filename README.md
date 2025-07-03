# Value Awards Email Sender

This script sends personalized HTML emails with an inline header image using Microsoft Outlook's COM interface. It's used to remind users to nominate colleagues for the Value Awards.

## Features

- Connects to local or network user sources (SQL Server).
- Embeds a header image in the email body.
- Sends emails via the Outlook application.
- Personalized greeting using the user's full name.
- Supports scheduling via Python `schedule`.

## Requirements

- Python 3.x
- Windows with Microsoft Outlook installed
- Install dependencies:

```bash
pip install pywin32 schedule pyodbc
