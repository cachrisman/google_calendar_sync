Simple [Google Script](https://script.google.com) to sync events from your personal calendar to your work calendar so coworkers will see busy spots for personal events like doctor appointments, etc.

To use, create a new script project under your work Google account and copy the code from google_calendar_sync.gs into the project. Fill out the first 4 variables with your details including the ID of a new Google Sheet to hold the logs. You need to share your personal calendar with your work account, so your work account can read all the events from your personal calendar.

For logging, the script relies on BetterLog, so go to the Resources menu and click on Libraries, then in the Add a library field, put in `1DSyxam1ceq72bMHsE6aOVeOl94X78WCwiYPytKi7chlg4x5GqiNXSw0l` and click add. ![libraries](https://i.imgur.com/ZDetW52.png)

Once the sync script works properly, you need to set up a time driven trigger to run the script every 10 minutes (or whatever time period makes sense for you) ![triggers](https://i.imgur.com/FUGOiHp.png)
