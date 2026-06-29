// messenger.js
const { chromium } = require("playwright");
const ExcelJS = require("exceljs");
const { sleep, randomDelay, getFirstName } = require("./utils");

const EXCEL_FILE = "./contacts.xlsx";
const SHEET_NAME = "Sheet1";
const MAX_MESSAGES_PER_RUN = 15; // safety limit

// Customize your message here
const buildMessage = (firstName) =>
  `Hi ${firstName}, thanks for connecting! I wanted to reach out regarding a potential collaboration opportunity. Would love to have a quick chat when you're free!`;

const run = async () => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(EXCEL_FILE);
  const sheet = workbook.getWorksheet(SHEET_NAME);

  const rows = [];
  sheet.eachRow((row, i) => {
    if (i === 1) return;
    rows.push({ row, i });
  });

  const browser = await chromium.launchPersistentContext("./profile", {
    headless: false,
    viewport: { width: 1280, height: 800 }
  });
  const page = await browser.newPage();

  // Check login
  await page.goto("https://www.linkedin.com/feed");
  await sleep(4000);
  const loggedIn = await page.$(".global-nav").catch(() => null);
  if (!loggedIn) {
    console.log("⚠️  Please log in manually, then press ENTER...");
    await new Promise(r => process.stdin.once("data", r));
  } else {
    console.log("✅ Logged in!\n");
  }

  let sentCount = 0;

  for (const { row, i } of rows) {
    if (sentCount >= MAX_MESSAGES_PER_RUN) {
      console.log(`\n⚠️  Reached message limit of ${MAX_MESSAGES_PER_RUN}. Stop for today.`);
      break;
    }

    const name = row.getCell(1).value;
    const linkedinUrl = row.getCell(3).value;
    const connectionStatus = row.getCell(6).value;
    const messageSentDate = row.getCell(7).value;

    // Only message people who accepted AND haven't been messaged yet
    if (connectionStatus !== "Accepted") {
      console.log(`Row ${i}: Skipping ${name} — not accepted yet (${connectionStatus})`);
      continue;
    }
    if (messageSentDate) {
      console.log(`Row ${i}: Skipping ${name} — message already sent`);
      continue;
    }

    console.log(`\nRow ${i}: Messaging ${name}...`);

    try {
      await page.goto(linkedinUrl);
      await randomDelay(3000, 5000);

      // Message button only shows for 1st degree connections
      const msgBtn = await page.$('button:has-text("Message")');
      if (!msgBtn) {
        console.log(`  ⚠️  Message button not found — may not be 1st degree`);
        row.getCell(9).value = "Message button not found";
        row.commit();
        await workbook.xlsx.writeFile(EXCEL_FILE);
        continue;
      }

      await msgBtn.click();
      await sleep(2000);

      // Type message in the chat box
      const msgBox = await page.$('.msg-form__contenteditable');
      if (!msgBox) {
        console.log(`  ❌ Message box not found`);
        continue;
      }

      const message = buildMessage(getFirstName(name));
      await msgBox.click();
      await msgBox.type(message, { delay: 40 }); // human typing speed
      await sleep(1500);

      // Send the message
      const sendBtn = await page.$('button.msg-form__send-button');
      if (sendBtn) {
        await sendBtn.click();
        await sleep(2000);

        const today = new Date().toLocaleDateString("en-IN");
        row.getCell(7).value = today;     // Message Sent Date
        row.getCell(8).value = message;   // Message Text (for your records)
        row.commit();
        await workbook.xlsx.writeFile(EXCEL_FILE);

        console.log(`  ✅ Message sent to ${name}`);
        sentCount++;
      } else {
        console.log(`  ❌ Send button not found`);
        row.getCell(9).value = "Error - send button not found";
        row.commit();
        await workbook.xlsx.writeFile(EXCEL_FILE);
      }

    } catch (err) {
      console.log(`  ❌ Error: ${err.message}`);
      row.getCell(9).value = `Error: ${err.message}`;
      row.commit();
      await workbook.xlsx.writeFile(EXCEL_FILE);
    }

    await randomDelay(5000, 10000);
  }

  console.log(`\n✅ messenger.js done! Sent ${sentCount} messages.`);
  await browser.close();
};

run().catch(console.error);