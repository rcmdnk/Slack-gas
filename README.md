# Slack-gas
Google Apps Script to retrieve Slack messages

## Preparation

* Prepare a slack app
    * Create a slack app from [here](https://api.slack.com/apps?new_app=1)
        * Select `From scratch`
        * Fill`App Name` as your like, pick a workspace from which you want to retrieve messages
        * Create App
    * Set scopes
        * Go `OAuth & Permissions`
        * In the `Scopes`, add scopes for `User Token Scopes`:
            * `channels:history`, `channels:read` for public channels
            * `groups:history`, `groups:read` for private channels
            * `im:history`, `im:read` for direct messages
            * `mpim:history`, `mpim:read` for group direct messages
            * `users:read` for the user list
        * If you do not need private or direct messages, you can omit these scopes.
    * Install and get token
        * Go `OAuth & Permissions`
        * Push `Install to Workspace` at `OAuth Tokens for Your Workspace`
        * Copy `User OAuth Token`
        * This `User OAuth Token` should be filled in [secret.gs](https://github.com/rcmdnk/Slack-gas/blob/main/secrets.gs)
* Prepare Google Sheets and an associated Apps Script project
    * Make new Sheets
    * Go `Extensions` -> `Apps Script`, then open Apps Script project
    * Make three `script` type files: [main.gs](https://github.com/rcmdnk/Slack-gas/blob/main/main.gs), [params.gs](https://github.com/rcmdnk/Slack-gas/blob/main/params.gs), [secret.gs](https://github.com/rcmdnk/Slack-gas/blob/main/secrets.gs)
    * Copy contents of these files in this repository to the above files.
    * Edit [secret.gs](https://github.com/rcmdnk/Slack-gas/blob/main/secrets.gs) in the Apps Script project to update `<YOUR_SLACK_TOKEN>` as your `User OAuth Token` that you got in the above procedure.


## Parameter settings

Set parameters as you like:

Name|Comment
:-|:-
TIME_ZONE|Set your time zone.
FULL_CHECK|false;
MAX_MESSAGES|500;
FILL_EACH|false;
CHANNEL_TYPES|'public_channel,private_channel,mpim,im';
CHANNEL_INCLUDE| [];
CHANNEL_EXCLUDE|[];
CHANNEL_EXCLUDE_ARCHIVED| false;
REMOVE_SAME_NAME_FILES| false;
DATETIME_COLUMN_WIDTH| 150;
USER_COLUMN_WIDTH| 100;
MESSAGE_COLUMN_WIDTH| 1000;
UNIXTIME_COLUMN_WIDTH| 150;
EDITED_TS_COLUMN_WIDTH| 150;
THREAD_TS_COLUMN_WIDTH| 150;
PIGINATE_LIMIT| 200;
