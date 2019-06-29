"use strict";
require('dotenv').config();
const {
  WebClient
} = require('@slack/web-api');
const GSS = require('google-spreadsheet');
const async = require('async');
const {
  promisify
} = require("util");
const keys = require('./keys.json');
const doc = new GSS(process.env.SHEET_ID);
const webClient = new WebClient(process.env.BOT_TOKEN);
const port = process.env.PORT || 3000;

const http = require("http");
var CronJob = require('cron').CronJob;

const sheetRowTitle = ['icons', 'name', 'displayname', 'email', 'slack'];



async function SaveMember() {
  try {
    //get info from slack
    let getSlackMember = await webClient.users.list({});
    getSlackMember = getSlackMember.members.filter(x => !x.is_bot && !x.id.match(/bot/gi));
    let memberIcons = getSlackMember.map(ele => `=IMAGE("${ele.profile.image_72}")`);
    let slackMember = getSlackMember.map(ele => Object({
      name: ele.profile.real_name,
      displayname: ele.profile.display_name,
      email: ele.profile.email,
      slack: `${process.env.WORKSPACE_TEAM_URL}${ele.id}`
    }));


    // sheet
    await promisify(doc.useServiceAccountAuth)(keys);
    const sheetInfo = await promisify(doc.getInfo)();
    const sheet = sheetInfo.worksheets[0];
    const cells = await promisify(sheet.getCells)({
      'min-row': 1,
      'max-row': 100,
      'min-col': 1,
      'max-col': 5,
      'return-empty': true
    });


    for (const cell of cells) {
      if (cell.row === 1 && cell.value.length === 0) {
        cells[cell.col - 1].value = sheetRowTitle[cell.col - 1];
        cells[cell.col - 1].save();
      } else {
        break;
      }
    }

    const rows = await promisify(sheet.getRows)({
      offset: 1
    });
    if (rows.length === 0) {




      slackMember.forEach(async function (member, index) {
        member["icons"] = memberIcons[index];
        await promisify(sheet.addRow)(member);
      });
    } else {
      let sheetMail = [];

      for (const row of rows) {
        let memberInSheet = JSON.stringify(Object({
          name: row.name,
          displayname: row.displayname,
          email: row.email,
          slack: row.slack
        }));


        let slackMemberString = slackMember.map(ele => JSON.stringify(ele));

        if (!slackMemberString.includes(memberInSheet)) {
          row.del();
          console.log(`deleted the ${row.email} from sheet`);
        }
      }


      let upDateRows = await promisify(sheet.getRows)({
        offset: 1
      });
      for (const ele of upDateRows) {
        console.log(ele.email);
        sheetMail.push(ele.email);
      }
      if (sheetMail.length !== 0) {
        // slackMember.forEach(async (member, index) => sheetMail.includes(member.email) ? true : member[icons] = memberIcons[index] await promisify(sheet.addRow)(member));
        for (let i = 0; i < slackMember.length; i++) {
          if (!sheetMail.includes(slackMember[i].email)) {
            slackMember[i]['icons'] = memberIcons[i];
            await promisify(sheet.addRow)(slackMember[i]);
          }
        }
      }

    }

  } catch (e) {
    console.log(e);
  }
}
const task = new CronJob('0 */2 * * * *', function () {
  SaveMember()
});
task.start();


http.createServer((req, res) => {
  res.writeHead(200, {
    'Content-Type': 'text/html'
  });
  res.write('Translate on Slack');
  res.end();
}).listen(port, '0.0.0.0', () => console.log(`Server running at ${port}`));