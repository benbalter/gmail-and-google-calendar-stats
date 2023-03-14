import { promises as fs } from "fs";
import path from "path";
import process from "process";
import { authenticate } from "@google-cloud/local-auth";
import { google, calendar_v3, gmail_v1 } from "googleapis";
import { stringify } from "csv-stringify";
import { urlToHttpOptions } from "url";
import addrs from "email-addresses";

const SCOPES = [
  "https://www.googleapis.com/auth/calendar.readonly",
  "https://www.googleapis.com/auth/gmail.readonly",
];
const TOKEN_PATH = path.join(process.cwd(), "token.json");
const CREDENTIALS_PATH = path.join(process.cwd(), "credentials.json");

// Your email address
const EMAIL = "benbalter@github.com";

// Years to pull event data for
const YEARS = [
  2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023,
];

// Internal domain (defaults to your email domain)
const DOMAIN = EMAIL.split("@")[1];

// "From" addresses to exclude from the email search
const EXCLUDE_FROM = [
  `notifications@${DOMAIN}`,
  `@users.noreply.${DOMAIN}`,
  `notifications@support.${DOMAIN}`,
  `shop@${DOMAIN}`,
  `noreply@${DOMAIN}`,
  `enterprise@${DOMAIN}`,
  `support@${DOMAIN}`,
];

const EXCLUDE_TO = [
  `@noreply${DOMAIN}`,
  `halp-noreply@${DOMAIN}`,
  `@lists.${DOMAIN}`,
  `all@${DOMAIN}`,
  `government@${DOMAIN}`,
  `@reply.${DOMAIN}`,
  `haps@${DOMAIN}`,
];

const EXCLUDE_SUBJECT = ["Hangout with ", "Thank you for your message"];

// Represents a single event on your calendar
class event {
  data: calendar_v3.Schema$Event;

  constructor(data: calendar_v3.Schema$Event) {
    this.data = data;
  }

  // Returns an array of attendees' email addresses
  // Note: Exclude the current user's email address
  attendeeEmails() {
    let attendeeEmails = this.data.attendees?.map((attendee) => attendee.email);
    attendeeEmails = attendeeEmails?.filter((email) => email !== EMAIL);
    return attendeeEmails;
  }

  // Returns an array of attendees' email domains
  attendeeDomains() {
    const attendeeDomains = this.attendeeEmails()?.map(
      (email) => email?.split("@")[1]
    );
    return [...new Set(attendeeDomains)];
  }

  // Returns true if the event is internal (i.e., all attendees are from the same domain)
  isInternal(): boolean {
    return JSON.stringify(this.attendeeDomains()) === JSON.stringify([DOMAIN]);
  }

  // Returns true if the current user is the event host
  isHost(): boolean {
    return this.data?.organizer?.email === EMAIL;
  }

  // Returns true if the event is a one-on-one
  isOneOnOne(): boolean {
    return this.numberOfAttendees() === 1;
  }

  // Returns the number of attendees
  numberOfAttendees() {
    return this.attendeeEmails()?.length || 0;
  }

  // Returns the event date
  date() {
    return this.data?.start?.dateTime;
  }

  // Returns the event summary (title)
  summary() {
    return this.data?.summary;
  }

  // Returns true if the event is confirmed
  isConfirmed() {
    return this.data?.status === "confirmed";
  }

  // Returns true if the current user is attending
  isAttending() {
    return this.data?.attendees?.find((attendee) => {
      return attendee.self && attendee.responseStatus === "accepted";
    });
  }

  // Returns true if the event should be included in the output
  shouldInclude() {
    return (
      this.isInternal() &&
      this.isConfirmed() &&
      (this.isHost() || this.isAttending()) &&
      this.numberOfAttendees() > 0 &&
      Date.parse(String(this.date())) <= Date.now()
    );
  }

  // Converts the event to a CSV-able row
  toRow() {
    return {
      date: this.date(),
      year: new Date(String(this.date())).getFullYear(),
      month: new Date(String(this.date())).getMonth() + 1,
      summary: this.summary(),
      isHost: this.isHost(),
      isOneOnOne: this.isOneOnOne(),
      numberOfAttendees: this.numberOfAttendees(),
    };
  }
}

class thread {
  data: gmail_v1.Schema$Thread;
  client: gmail_v1.Gmail;
  messages: message[] | undefined;

  constructor(data: gmail_v1.Schema$Thread, client: gmail_v1.Gmail) {
    this.data = data;
    this.client = client;
  }

  // Retrieves the threads messages
  async getMessages(): Promise<message[] | undefined> {
    if (this.messages) return this.messages;

    console.log(`Getting messages for thread ${this.data.id}`);
    const res = await this.client.users.threads.get({
      userId: "me",
      id: this.data.id || "",
      format: "metadata",
    });
    this.messages = res.data.messages?.map((m) => new message(m));
    return this.messages;
  }

  toRows() {
    return this.messages?.map((m) => m.toRow());
  }
}

class message {
  data: gmail_v1.Schema$Message;

  constructor(data: gmail_v1.Schema$Message) {
    this.data = data;
  }

  // Returns the message headers
  headers(): gmail_v1.Schema$MessagePartHeader[] | undefined {
    return this.data.payload?.headers;
  }

  getHeader(name: string) {
    const header = this.headers()?.find((h) => {
      return h.name === name;
    });
    return header?.value;
  }

  subject() {
    return this.getHeader("Subject");
  }

  to() {
    const to = this.getHeader("To");
    const addr = addrs.parseAddressList(to || "");
    if (!addr) {
      return;
    }

    const addresses = addr.map((a) => {
      if (a.type === "group") {
        return a.addresses;
      } else {
        return [a];
      }
    });

    return addresses.flat();
  }

  from() {
    const from = this.getHeader("From");
    const addr = addrs.parseOneAddress(from || "");

    if (addr?.type === "group") {
      return addr.addresses[0];
    } else {
      return addr;
    }
  }

  isSender() {
    const from = this.from();

    if (!from) {
      return false;
    }

    return from.address.replace(".", "") === EMAIL;
  }

  isInternal() {
    return (
      this.from()?.domain === DOMAIN &&
      this.to()?.every((to) => {
        return to.domain === DOMAIN;
      })
    );
  }

  excludedFrom() {
    return EXCLUDE_FROM.includes(this.from()?.address || "");
  }

  excludedTo() {
    return this.to()?.some((to) => {
      return EXCLUDE_TO.includes(to.address);
    });
  }

  shouldInclude() {
    return this.isInternal() && !this.excludedFrom() && !this.excludedTo();
  }

  toRow() {
    return {
      date: this.data.internalDate,
      year: new Date(Number(this.data.internalDate)).getFullYear(),
      month: new Date(Number(this.data.internalDate)).getMonth() + 1,
      subject: this.subject(),
      from: this.from()?.address,
      to: this.to()
        ?.map((to) => to.address)
        .join(","),
      isSender: this.isSender(),
    };
  }
}

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
  const items = res.data.items;
  if (!items || items.length === 0) {
    console.log("No events found.");
    return;
  }

  return items.map((item) => new event(item));
}

async function getThreads(auth, year: number) {
  const gmail = google.gmail({ version: "v1", auth });
  const fromFilter = EXCLUDE_FROM.map((email) => `-"from:${email}"`).join(" ");
  const toFilter = EXCLUDE_TO.map((email) => `-"to:${email}"`).join(" ");
  const subjectFilter = EXCLUDE_SUBJECT.map(
    (subject) => `-"subject:${subject}"`
  ).join(" ");
  const q = `to:(${DOMAIN}) from:(${DOMAIN}) ${fromFilter} ${toFilter} ${subjectFilter} -{"invite.ics"} after:${year}/01/01 before:${year}/12/31`;

  console.log(`Listing threads using the following query: ${q}`);
  const res = await gmail.users.threads.list({
    userId: "me",
    q,
    maxResults: 500,
    includeSpamTrash: false,
  });

  const threads = res.data.threads;
  if (!threads || threads.length === 0) {
    console.log("No threads found.");
    return;
  }

  console.log(`Found ${threads.length} threads`);
  return threads.map((t) => new thread(t, gmail));
}

// The main event
async function run() {
  const client = await authorize();

  const threads = await getThreads(client, 2013);

  if (threads) {
    let threadsFiltered: thread[] = [];
    for (const t of threads) {
      const messages = await t.getMessages();
      if (messages && messages.every((m) => m.shouldInclude())) {
        threadsFiltered = [...threadsFiltered, t];
      }
    }

    console.log(`Filtered to ${threadsFiltered.length} threads`);
    let rows = threadsFiltered.map((t) => t.toRows()).flat();
    stringify(
      rows,
      {
        header: true,
      },
      function (err, output) {
        console.log(output);
        fs.writeFile(`${__dirname}/../messages.csv`, output);
      }
    );
  }

  /* 
  let events: event[] = [];

  // Get events for each year
  for (const year of YEARS) {
    const yearEvents = await getEvents(client, year);
    if (!yearEvents) {
      continue;
    }
    events = [...events, ...yearEvents];
  }

  const rows = events.filter((e) => e.shouldInclude()).map((e) => e.toRow());
  console.log(rows);

  // Write to CSV
  stringify(
    rows,
    {
      header: true,
    },
    function (err, output) {
      console.log(output);
      fs.writeFile(`${__dirname}/../data.csv`, output);
    }
  );*/
}

run();
