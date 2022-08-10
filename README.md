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
    * Edit **secret.gs** in the Apps Script project to update `<YOUR_SLACK_TOKEN>` as your `User OAuth Token` that you got in the above procedure.


## Parameter settings

Set parameters as you like:

Name|Comment
:-|:-
TIME_ZONE|Set your time zone. Set `null` to use the script's time zone.
FULL_CHECK|If true, no limit is set at API to get messages. If false, set the oldest as the latest message's time stamp in the channel sheet. If false, messages in the thread could be ignored even if they are newer but the parent messages are older.
MAX_MESSAGES|Max number of messages to retrieve in the one job. 100 or 200 are safe, but 500 may exceed the time limit of the Apps Script (6 min). But in most cases, you can just re-run the script and the job will retrieve the remaining messages.
FILL_EACH|If true, fill each message in the sheet one by one. Otherwise, all messages in the channel are filled at the same time. This option may be useful if you have many messages in the channel and want to set MAX_MESSAGES higher, but it makes the job slow.
CHANNEL_TYPES|Set channel types that you want to retrieve. Types are: `public_channel`, `private_channel`, `mpim`, and `im`. Multiple types can be set as a comma-separated string like `'public_channel,private_channel,mpim,im'`.
CHANNEL_INCLUDE|Channels to be included. If it is empty, no filter is applied by this value.
CHANNEL_EXCLUDE|Channels to be excluded.
CHANNEL_EXCLUDE_ARCHIVED|If true, archive channels are excluded.
SAVE_FILE|If true, files attached to messages are saved in the the Google Drive.
REMOVE_SAME_NAME_FILES|If true, older attached files with the same name are removed. This breaks links in the older messages.
COVERAGE|Time in Seconds of the retrieving period. `2592000` (`60*60*24*30`) to retrieve messages in about month. Set 0 or `null` to retrieve all messages.
Time threshold (in Unixtime) after which messages are retrieved. e.g. `1640995200` for 2022/01/01 00:00:00 (UTC). If it is `null`, no time threshold is set.
SAVE_MESSAGE_JASON|If true, messages are saved as a json format in the Google Drive. Note: This makes the job very slow and only ~100 messages can be retrieved.
REMOVE_OLD_MESSAGE|If true, saved message files are removed when newer same time stamp messages are saved (it happens when the message is edited).
UPDATE_COLUMN_NAME|If true, column names are updated even if the sheet already exists.
DATETIME_COLUMN_WIDTH|Column width of Datetime.
USER_COLUMN_WIDTH| Column width of User.
MESSAGE_COLUMN_WIDTH| Column width of Message.
UNIXTIME_COLUMN_WIDTH| Column width of Thread time stamp.
EDITED_TS_COLUMN_WIDTH| Column width of UnixTime.
THREAD_TS_COLUMN_WIDTH| Column width of Edited time stamp.
PIGINATE_LIMIT| A limit number to retrieve objects by the slack API at one time.

## Retrieve messages by hand

* Open **main.gs** in the Apps Script project
* Save files by pushing the save (disc icon) button.
* Select `run` function.
* Run the job by `Run`

If you see a time-out error, you can retry until all messages are retrieved.

## Schedule job

It is good to schedule the job to retrieve messages every day.

* Go `Trigger` (Clock icon) in the Apps Script project.
* `Add Trigger`
    * Choose which function to run: `run`
    * Which runs at deployment: `Head`
    * Select event source: `Time-driven`
    * Select type of time based trigger: `Day timer`
    * Select time of day: `0 am to 1 am`

This trigger runs the job between 0 am to 1 am every day.

## More comments on parameters

### SAVE_MESSAGE_JASON

Google Apps Script's job has a time limit of 6 min.
It is too short to retrieve many messages at once.

If `SAVE_MESSAGE_JASON` is true, about 100 messages are the upper limit to be done in 6 min.

If you don't have so much messages or want to retrieve only the recent and set short `COVERAGE` value,
it may not a problem.

Otherwise, set `MAX_MESSAGES` as 50~100 and run the job sometime until all messages can be retrieved.

These full message information may not be needed or you can obtain by the export function of the slack.

Therefore, `SAVE_MESSAGE_JSON = false` is recommended for most cases.

If your workspace have a few posts per day, you can set `SAVE_MESSAGE_JSON=false` for the daily job.


### FULL_CHECK/COVERAGE

If `FULL_CHECK = false`, only newer messages than the latest message in the channel retrieved in the Sheet.

It means if the thread starts before the latest message,
messages in the thread can not be retrieved even if they are newer than the latest message.




