const fs = require('fs').promises;
const path = require('path');
const ical = require('ical');
const moment = require('moment-timezone');

const TIMEZONE = 'America/Chicago';

/** 
* Checks if file was modified in the last week
* @param {fs.Stats} fileStats
* @return {boolean}
*/

function isFileThisWeek(fileStats) {
    const now = new Date();
    const fileDate = new Date(fileStats.mtime);
    const weekStart = new Date(now.setDate(now.getDate() - now.getDay()));
    const weekEnd = new Date(now.setDate(now.getDate() + 6));
  
    return fileDate >= weekStart && fileDate <= weekEnd;
}

/**
 * Reads and parses .ics files from the specified folder, filtering files modified within the current week.
 * @param {string} folderPath
 * @return {Promise<Array>}
 */


async function readIcsFiles(folderPath) {
    try {
        const files = await fs.readdir(folderPath);
        let values = [];

        for (const file of files) {
            const filePath = path.join(folderPath, file);
            const fileStats = await fs.stat(filePath);

            if (file.endsWith('.ics') && isFileThisWeek(fileStats)) {
                const data = await fs.readFile(filePath, 'utf8');
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
        return values;
    } catch (err) {
        console.error('Unable to read .ics files: ', err);
        throw err;
    }
}

module.exports = {
    readIcsFiles,
};