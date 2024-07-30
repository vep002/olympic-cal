const { google } = require('googleapis');
const { authorize } = require('./auth');
const { readIcsFiles } = require('./utils');

async function updateSheet(auth) {
    const sheets = google.sheets({ version: 'v4', auth });
    const spreadsheetId = '1dAW18Jk2J0BnMfrWwgnup2S3GftnXqK4Y_R21TPZIHs';
    const range = 'Sheet1!A1';

    try {
        await sheets.spreadsheets.values.update({
            spreadsheetId,
            range,
            valueInputOption: 'RAW',
            resource: {
                values: [['Subject', 'Start Date', 'Start Time', 'End Date', 'End Time', 'All Day Event', 'Description', 'Location', 'Private']],
            },
        });
        console.log('Sheet updated.');
        
        const folderPath = '/Users/victoriaperkins/Downloads';
        const values = await readIcsFiles(folderPath);

        if (values.length > 0) {
            await sheets.spreadsheets.values.append({
                spreadsheetId,
                range: 'Sheet1!A2',
                valueInputOption: 'RAW',
                resource: { values },
            });
            console.log('%d cells updated.', values.length * 7);
        } else {
            console.log('No events for this week.');
        } 
    } catch (e) {
            console.error('Error updating sheet: ', e);
        }
}

authorize().then(updateSheet).catch(console.error);
    