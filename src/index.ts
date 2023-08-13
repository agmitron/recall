import { google, sheets_v4 } from 'googleapis';
import * as fs from 'fs/promises';
import * as readline from 'readline';
import { OAuth2Client } from 'google-auth-library';

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly'];
const TOKEN_PATH = 'token.json'; // Change this if needed

// Load client secrets from a file you downloaded from the Google Cloud Console
async function loadClientSecrets(): Promise<any> {
  try {
    const content = await fs.readFile('credentials.json');
    return JSON.parse(content.toString());
  } catch (err: any) {
    throw new Error('Error loading client secret file: ' + err.message);
  }
}

async function authorize(credentials: any): Promise<sheets_v4.Sheets | void> {
  const { client_secret, client_id } = credentials.web;
  const oAuth2Client = new google.auth.OAuth2(
    client_id,
    client_secret,
    'https://google.com'
  );

  await getAccessToken(oAuth2Client);

  try {
    const tokenContent = await fs
      .readFile(TOKEN_PATH)
      .then((c) => c.toString())
      .then(JSON.parse);

    console.log({ tokenContent });
    oAuth2Client.setCredentials(tokenContent);
    return google.sheets({ version: 'v4', auth: oAuth2Client });
  } catch (err) {
    return getAccessToken(oAuth2Client);
  }
}

async function getAccessToken(oAuth2Client: OAuth2Client): Promise<void> {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });
  console.log('Authorize this app by visiting this URL:', authUrl);

  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  return new Promise((resolve, reject) => {
    rl.question('Enter the code from that page here: ', (code) => {
      rl.close();
      oAuth2Client.getToken({ code }, async (err, token: any) => {
        if (err)
          return reject(
            new Error('Error retrieving access token: ' + err.message)
          );
        oAuth2Client.setCredentials(token);
        try {
          await fs.writeFile(TOKEN_PATH, JSON.stringify(token));
          console.log('Token stored to', TOKEN_PATH);
          resolve();
        } catch (err) {
          console.error(err);
          reject(err);
        }
      });
    });
  });
}

async function listSpreadsheetValues(sheets: sheets_v4.Sheets): Promise<void> {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: '1GZxvoZ9ItrQ4BLlxY-e4JfImNIxz2yyiVzzKEQX-g7I',
      range: 'Sheet1!A1:B2', // Change this to the range you want to read
    });

    const rows = res.data.values;
    if (rows?.length) {
      console.log('Data:');
      rows.forEach((row: any) => {
        console.log(`${row[0]}, ${row[1]}`);
      });
    } else {
      console.log('No data found.');
    }
  } catch (err: any) {
    throw new Error('The API returned an error: ' + err.message);
  }
}

async function main() {
  try {
    const credentials = await loadClientSecrets();
    const sheets = await authorize(credentials);
    if (sheets) await listSpreadsheetValues(sheets);
  } catch (err) {
    console.error(err);
  }
}

main();
