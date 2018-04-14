# Get Gmail Attachments

Utility app that downloads email attachments from a Gmail inbox filtering
for a given search string. You'll need to create a new
project in the [Google Developers Console](https://console.developers.google.com/)
and register an OAuth native client (called an *installed app* in the Google
developer console) and download the JSON representation of it and add it to the
project with the file name **client\_secret.json**. The JSON should look like this:

    {
        "installed": {
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "client_secret": "Get this from the Google developer console",
            "token_uri": "https://accounts.google.com/o/oauth2/token",
            "client_email": "",
            "redirect_uris": [ "urn:ietf:wg:oauth:2.0:oob", "oob" ],
            "client_x509_cert_url": "",
            "client_id": "Get this from the Google developer console",
            "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs"
        }
    }

The project in the Google developer console will need
the Gmail API enabled. All this is documented quite well in the
[Gmail REST API](https://developers.google.com/gmail/api/?hl=en_US) site.

## How to use

You can run the app from a console window like so:

    get-gmail-attachments "subject:Your bank statement" "c:\bank-statements"

This will get all attachments from all mails in your gmail inbox that match the string
`Your bank statement` in the subject line. You can use all the search operators
that Gmail supports. The attachments will get saved into the `c:\bank-statements`
folder.
