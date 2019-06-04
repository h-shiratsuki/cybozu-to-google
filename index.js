#!/usr/bin/env node
const meow = require('meow');
const puppeteer = require('puppeteer');
const CalendarAPI = require('node-google-calendar');
const fs = require('fs');
const os = require('os');
const path = require('path');
const moment = require('moment');
const parseCsv = require('csv-parse/lib/sync');
const rc = require('rc');

const cli = meow(`
  c2g - Move schedules from Cybozu to Google Calendar.

  Usage
    $ c2g

  Options
    --config, -c    Pass config file. By default, c2g will read "~/.config/c2g".
    --json, -j      Pass config json when not use config file.
    --quiet,  -q    Hide debug messages
    --show, -s      Show browser window
    --version, -v   Show version
    --help, -h      Show this help

  Examples
    $ c2g -c ./config.json
      >>>> Fetching events from Cybozu Calendar...DONE
      >>>> Fetching events from Google Calendar...DONE
      >>>> Inserting new events...
        Inserted: [会議] B社MTG
        Inserted: [会議] 目標面談
      >>>> Inserted 2 events.
      >>>> Deleting removed events...
      	Deleted: [外出] 幕張メッセ
      >>>> Deleted 1 events.
`, {
  flags: {
    config: {
      type: 'string',
      alias: 'c',
    },
    json: {
      type: 'string',
      alias: 'j',
    },
    quiet: {
      type: 'boolean',
      alias: 'q',
    },
    show: {
      type: 'boolean',
      alias: 's',
    },
    help: {
      type: 'boolean',
      alias: 'h',
    },
    version: {
      type: 'boolean',
      alias: 'v',
    },
  }
});

// Detect CLI flags
if (cli.flags.help) {
  cli.showHelp();
}
if (cli.flags.version) {
  cli.showVersion();
}

const log = cli.flags.quiet ?
  () => {} :
  (str) => process.stdout.write(str);

let config = rc('c2g');
if (cli.flags.config) {
  try {
    config = JSON.parse(fs.readFileSync(cli.flags.config));
  } catch(e) {
    console.error(e);
    process.exit(-1);
  }
} else if (cli.flags.json) {
  try {
    config = JSON.parse(cli.flags.json);
  } catch(e) {
    console.error(e);
    process.exit(-1);
  }
}

// Utils
const calendar = new CalendarAPI(config.calendar);

const now = new Date();
const year = now.getFullYear();
const month = now.getMonth() + 1;
const day = now.getDate();

const csvDir = path.join(os.tmpdir(), Date.now() + '');
const csvPath = path.join(csvDir, 'schedule.csv');

const isEqualDate = (t1, t2) => {
  let isEqualStart = false;
  let isEqualEnd = false;
  const check = (d1, d2) => {
    return new Date(d1).getTime() === new Date(d2).getTime();
  };
  if (t1.start.date && t2.start.date && check(t1.start.date, t2.start.date)) {
    isEqualStart = true;
  } else if (t1.start.dateTime && t2.start.dateTime && check(t1.start.dateTime, t2.start.dateTime)) {
    isEqualStart = true;
  }
  if (t1.end.date && t2.end.date && check(t1.end.date, t2.end.date)) {
    isEqualEnd = true;
  } else if (t1.end.dateTime && t2.end.dateTime && check(t1.end.dateTime, t2.end.dateTime)) {
    isEqualEnd = true;
  }
  return isEqualStart && isEqualEnd;
};

// Main
(async () => {
  const browser = await puppeteer.launch({ headless: !cli.flags.show });
  const page = await browser.newPage();
  await page._client.send('Page.setDownloadBehavior', {
    behavior: 'allow',
    downloadPath: csvDir,
  });

  log('>>>> Fetching events from Cybozu Calendar...\n');

  // Login
  await page.goto(config.cybozuUrl);
  await page.type('input[name="username"]', config.username)
  await page.type('input[name="password"]', config.password);
  await page.waitFor(1000);
  await page.click('form input[type=submit]');
  await page.waitForNavigation({ timeout: 30000, waitUntil: 'domcontentloaded' });

  // Go to CSV exporter page
  await page.goto(`${config.cybozuUrl}o/ag.cgi?page=PersonalScheduleExport`, { waitUntil: 'domcontentloaded' });

  // Input date
  await page.select('select[name="SetDate.Year"]', year + '');
  await page.select('select[name="SetDate.Month"]', month + '');
  await page.select('select[name="SetDate.Day"]', day + '');
  await page.select('select[name="EndDate.Year"]', (year + 1) + '');
  await page.select('select[name="EndDate.Month"]', month + '');
  await page.select('select[name="EndDate.Day"]', day + '');
  await page.select('select[name="oencoding"]', 'UTF-8');

  // Download CSV
  await page.click('.vr_hotButton');
  await page.waitFor(3000);
  await browser.close();

  // Parse CSV
  const newEvents = [];
  const csv = fs.readFileSync(csvPath, 'utf8');
  parseCsv(csv).forEach((line, i) => {
    if (i === 0) { return; }
    try {
      let startDate = line[0];
      let startTime = line[1];
      let endDate = line[2];
      let endTime = line[3];

      // Same day.
      if (startDate === endDate) {
        // Only end time is empty.
        if (startTime !== '' && endTime === '') {
          endTime = '23:59:59';
        }
      } else if (startDate > endDate) {
        // Swap if the date is reversed.
        const tmpDate = startDate;
        const tmpTime = startTime;
        startDate = endDate;
        startTime = endTime;
        endDate = tmpDate;
        endTime = tmpTime;
      }

      const startMoment = moment(new Date(startDate + ' ' + startTime));
      let endMoment = moment(new Date(endDate + ' ' + endTime));
      
      let summary = `${line[5]}`;
      if (line[4] !== '') {
        summary = `[${line[4]}] ` + summary;
      }
      const description = line[6];
      const location = line[8];

      let start, end;
      if (startTime === '' && endTime === '') {
        if (startMoment.isSameOrAfter(endMoment)) {
          endMoment = startMoment.clone();
          endMoment.add(1, 'd');
        }
        start = { date: startMoment.format("YYYY-MM-DD") };
        end = { date: endMoment.format("YYYY-MM-DD") };
      } else {
        start = { dateTime: startMoment.toISOString() };
        end = { dateTime: endMoment.toISOString() };
      }

      newEvents.push({ start, end, location, summary, description });
    } catch(e) {
      log(e + '\n');
    }
  });

  log('DONE\n');
  log('>>>> Fetching events from Google Calendar...\n');

  const oldEvents = await calendar.Events.list(config.calendar.calendarId.primary, {
    timeMin: moment().startOf('day').toISOString(),
    timeMax: moment().startOf('day').add(1, 'year').toISOString(),
    q: '',
    singleEvents: true,
    orderBy: 'startTime'
  });

  log('DONE\n');
  log(`>>>> Inserting new events...\n`);

  let insertedCount = 0;
  for (const event of newEvents) {
    let isEqual = false;

    for (const old of oldEvents) {
      if (old.summary === event.summary) {
        // Check date and time.
        if (isEqual = isEqualDate(old, event)) {
          break;
        }
      }
    }
    
    if (!isEqual) {
      await calendar.Events.insert(config.calendar.calendarId.primary, event);
  
      log(`\tInserted: ${event.summary}\n`);
      insertedCount++;
    }
  }

  log(`>>>> Inserted ${insertedCount} events.\n`);
  log(`>>>> Deleting removed events...\n`);

  let deletedCount = 0;
  for (const old of oldEvents) {
    let isEqual = false;

    for (const event of newEvents) {
      if (old.summary === event.summary) {
        // Check date and time.
        if (isEqual = isEqualDate(old, event)) {
          break;
        }
      }
    }

    if (!isEqual) {
      await calendar.Events.delete(config.calendar.calendarId.primary, old.id, { sendNotifications: true });

      log(`\tDeleted: ${old.summary}\n`);
      deletedCount++;
    }
  }
  log(`>>>> Deleted ${deletedCount} events.\n`);
})();
