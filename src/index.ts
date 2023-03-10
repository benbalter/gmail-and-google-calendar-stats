import { promises as fs } from "fs";
import path from "path";
import process from "process";
import { authenticate } from "@google-cloud/local-auth";
import { google } from "googleapis";

const SCOPES = ["https://www.googleapis.com/auth/calendar.readonly"];
const TOKEN_PATH = path.join(process.cwd(), "token.json");
const CREDENTIALS_PATH = path.join(process.cwd(), "credentials.json");

/**
 * Reads previously authorized credentials from the save file.
 *
 * @return {Promise<OAuth2Client|null>}
 */
async function loadSavedCredentialsIfExist() {
  try {
    const content = await fs.readFile(TOKEN_PATH);
    const credentials = JSON.parse(content.toString());
    return google.auth.fromJSON(credentials);
  } catch (err) {
    return null;
  }
}

/**
 * Serializes credentials to a file compatible with GoogleAUth.fromJSON.
 *
 * @param {OAuth2Client} client
 * @return {Promise<void>}
 */
async function saveCredentials(client): Promise<void> {
  const content = await fs.readFile(CREDENTIALS_PATH);
  const keys = JSON.parse(content.toString());
  const key = keys.installed || keys.web;
  const payload = JSON.stringify({
    type: "authorized_user",
    client_id: key.client_id,
    client_secret: key.client_secret,
    refresh_token: client.credentials.refresh_token,
  });
  await fs.writeFile(TOKEN_PATH, payload);
}

/**
 * Load or request or authorization to call APIs.
 *
 */
async function authorize() {
  const savedClient = await loadSavedCredentialsIfExist();
  if (savedClient) {
    return savedClient;
  }
  const client = await authenticate({
    scopes: SCOPES,
    keyfilePath: CREDENTIALS_PATH,
  });
  if (client.credentials) {
    await saveCredentials(client);
  }
  return client;
}

/**
 * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
 */
async function getEvents(auth, year: number) {
  console.log("Getting events for", year);
  const calendar = google.calendar({ version: "v3", auth });
  const res = await calendar.events.list({
    calendarId: "primary",
    maxResults: 2500,
    singleEvents: true,
    orderBy: "startTime",
    timeMax: new Date(`${year}-12-31`).toISOString(),
    timeMin: new Date(`${year}-01-01`).toISOString(),
  });
  const events = res.data.items;
  if (!events || events.length === 0) {
    console.log("No events found.");
    return;
  }
  return events;
}

async function countEvents(auth, year: number) {
  const events = await getEvents(auth, year);
  if (!events) {
    return 0;
  }
  let count = 0;
  for (const event of events) {
    let attendeeEmails = event.attendees?.map((attendee) => attendee.email);
    attendeeEmails = attendeeEmails?.filter(
      (email) => email !== "benbalter@github.com"
    );
    let attendeeDomains = attendeeEmails?.map((email) => email?.split("@")[1]);
    attendeeDomains = [...new Set(attendeeDomains)];

    if (attendeeDomains?.length === 0) {
      continue;
    }

    if (JSON.stringify(attendeeDomains) !== JSON.stringify(["github.com"])) {
      continue;
    }

    console.log(event.summary);
    console.log(attendeeDomains);
    count++;
  }

  console.log(count);
  return count;
}

async function run() {
  const years = [
    2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023,
  ];
  const client = await authorize();
  const counts = {};
  for (const year of years) {
    counts[year] = await countEvents(client, year);
  }
  console.log(counts);
}

run();
