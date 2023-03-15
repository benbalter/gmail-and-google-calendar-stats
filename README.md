# GMail and Google Calendar stats

*Scrapes your GMail and Google Calendar data and returns it as a CSV for further analysis.*

## Why? 

Because I wanted to understand meeting and email trends at my employer over time.

## How it works

The script uses the Google API to access your GMail and Google Calendar data. It then parses the data and returns it as a CSV file.

## What it looks at

* Internal-only meetings (all attendees in the same domain)
* Internal-only emails (all recipients in the same domain, excluding automated emails)

## How to use

1. Clone the repo
2. Follow [the quickstart instructions for the Google API](https://developers.google.com/google-apps/calendar/quickstart/nodejs) to create a project, and get a `client_secret.json` file.
3. `npm run run` to authorize the script to access your data, and then run it.
4. Grab the resulting `messages.csv` and `events.csv` files and open them in your favorite spreadsheet program.
5. :cry: at the results. :chart_with_upward_trend: