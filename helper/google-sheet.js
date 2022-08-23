const fs = require('fs');
const { google } = require('googleapis');
const credentials = JSON.parse(fs.readFileSync('google-client-secret.json', 'utf-8')); 
const { client_secret: clientSecret, client_id: clientId, redirect_uris: redirectUris, } = credentials.installed;
const oAuth2Client = new google.auth.OAuth2( clientId, clientSecret, redirectUris[0], );
const token = fs.readFileSync('google-oauth-token.json', 'utf-8'); 
oAuth2Client.setCredentials(JSON.parse(token));

  /**
   * 새로운 시트
   * @param {string} spreadsheetId
   * @param {string} range
   * @returns {Promise}
   */
  exports.create = async (spreadsheetId,range,title) => {
     const sheets = await google.sheets({ version: 'v4', auth: oAuth2Client });
     return sheets.spreadsheets.create({
         resource: { 
            properties:{title}
        }
     });

     /** * Read a spreadsheet. 
      * @param {string} spreadsheetId 
      * @param {string} range 
      * @returns {Promise.} */
      exports.read = async (spreadsheetId, range) => { 
         const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
          return sheets.spreadsheets.values.get({ spreadsheetId, range, }) .then(_.property('data.values')); };
}