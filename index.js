const fs = require('fs').promises;
const path = require('path');
const process = require('process');
const {authenticate} = require('@google-cloud/local-auth');
const {google} = require('googleapis');
const ical = require('ical');
const moment = require('moment-timezone');

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = path.join(process.cwd(), 'token.json');
const CREDENTIALS_PATH = path.join(process.cwd(), 'credentials.json');

const TIMEZONE = 'America/Chicago';

/**
 * Reads previously authorized credentials from the save file.
 *
 * @return {Promise<OAuth2Client|null>}
 */
async function loadSavedCredentialsIfExist() {
  try {
    const content = await fs.readFile(TOKEN_PATH);
    const credentials = JSON.parse(content);
    return google.auth.fromJSON(credentials);
  } catch (err) {
    return null;
  }
}

/**
 * Serializes credentials to a file compatible with GoogleAuth.fromJSON.
 *
 * @param {OAuth2Client} client
 * @return {Promise<void>}
 */
async function saveCredentials(client) {
  const content = await fs.readFile(CREDENTIALS_PATH);
  const keys = JSON.parse(content);
  const key = keys.installed || keys.web;
  const payload = JSON.stringify({
    type: 'authorized_user',
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
  let client = await loadSavedCredentialsIfExist();
  if (client) {
    return client;
  }
  client = await authenticate({
    scopes: SCOPES,
    keyfilePath: CREDENTIALS_PATH,
  });
  if (client.credentials) {
    await saveCredentials(client);
  }
  return client;
}

function isFileThisWeek(fileStats) {
    const now = new Date();
    const fileDate = new Date(fileStats.mtime);
    const weekStart = new Date(now.setDate(now.getDate() - now.getDay()));
    const weekEnd = new Date(now.setDate(now.getDate() + 6));
  
    return fileDate >= weekStart && fileDate <= weekEnd;
  }

async function listEvents(auth) {
    const sheets = google.sheets({ version: 'v4', auth });
    const spreadsheetId = '1dAW18Jk2J0BnMfrWwgnup2S3GftnXqK4Y_R21TPZIHs';
    const range = 'Sheet1!A1';

    sheets.spreadsheets.values.update({
        spreadsheetId,
        range,
        valueInputOption: 'RAW',
        resource: {
            values: [['Subject', 'Start Date', 'Start Time', 'End Date', 'End Time', 'All Day Event', 'Description', 'Location', 'Private']],
        },

    });

    const folderPath = '/Users/victoriaperkins/Downloads';

    try{
        const files = await fs.readdir(folderPath);
        console.log('Reading files in folder: ', folderPath);

        let values = [];
        for (const file of files) {
            const filePath = path.join(folderPath, file);
            const fileStats = await fs.stat(filePath);
            if(file.endsWith('.ics') && isFileThisWeek(fileStats)) {
                const data = await fs.readFile(path.join(folderPath, file), 'utf8');
                const events = ical.parseICS(data);
                for (let i in events) {
                    if (events.hasOwnProperty(i)) {
                        const event = events[i];
                        if (event.start && event.end && event.summary) {
                            const start = moment.tz(event.start, TIMEZONE);
                            const end = moment.tz(event.end, TIMEZONE);
                            values.push([
                                event.summary,
                                start.format('YYYY-MM-DD'),
                                start.format('HH:mm:ss'),
                                end.format('YYYY-MM-DD'),
                                end.format('HH:mm:ss'),
                                false,
                                event.description || '',
                                event.location || '',
                                false,
                            ]);
                        } 
                    }
                }
            }
        }

        await sheets.spreadsheets.values.append({
            spreadsheetId,
            range: 'Sheet1!A2',
            valueInputOption: 'RAW',
            resource: { values },
        });

        console.log('%d cells updated.', values.length * 5);

    } catch (err) {
        console.error('Unable to scan directory: ', err);
    }
}

authorize().then(listEvents).catch(console.error);
    