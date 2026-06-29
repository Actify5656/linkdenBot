// connector.js
const { chromium } = require("playwright");
const ExcelJS = require("exceljs");
const { sleep, randomDelay, getFirstName } = require("./utils");

const EXCEL_FILE = "./contacts.xlsx";
const SHEET_NAME = "Sheet1";
const MAX_CONNECTIONS_PER_RUN = 20; // safety limit

// Customize your connection note here
const buildNote = (firstName) =>
  `Hi ${firstName}, I came across your profile and would love to connect!`;

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

  // Check LinkedIn login
  console.log("Checking LinkedIn login...");
  await page.goto("https://www.linkedin.com/feed");
  await sleep(4000);

  const loggedIn = await page.$(".global-nav").catch(() => null);
  if (!loggedIn) {
    console.log("⚠️  Not logged in! Please log in manually in the browser.");
    console.log("Press ENTER here after you are logged in...");
    await new Promise(r => process.stdin.once("data", r));
  } else {
    console.log("✅ Logged in!\n");
  }

  let sentCount = 0;

  for (const { row, i } of rows) {
    if (sentCount >= MAX_CONNECTIONS_PER_RUN) {
      console.log(`\n⚠️  Reached daily limit of ${MAX_CONNECTIONS_PER_RUN}. Stop for today.`);
      break;
    }

    const name = row.getCell(1).value;
    const linkedinUrl = row.getCell(3).value;
    const urlStatus = row.getCell(4).value;
    const connectionStatus = row.getCell(6).value;

    // Skip if no URL, or connection already sent/attempted
    if (!linkedinUrl || urlStatus !== "URL Found") {
      console.log(`Row ${i}: Skipping — no valid URL`);
      continue;
    }
    if (connectionStatus && connectionStatus !== "") {
      console.log(`Row ${i}: Skipping ${name} — already processed (${connectionStatus})`);
      continue;
    }

    console.log(`\nRow ${i}: Sending connection to ${name}...`);

    try {
      await page.goto(linkedinUrl);
      await randomDelay(3000, 5000);

      // Try to find Connect button directly on profile
      let connectBtn = await page.$('button:has-text("Connect")');

      // If not visible, it might be inside "More" dropdown
      if (!connectBtn) {
        const moreBtn = await page.$('button:has-text("More")');
        if (moreBtn) {
          await moreBtn.click();
          await sleep(1500);
          connectBtn = await page.$('div[role="option"]:has-text("Connect")');
        }
      }

      if (!connectBtn) {
        console.log(`  ⚠️  Connect button not found — may already be connected`);
        row.getCell(6).value = "Already Connected";
        row.commit();
        await workbook.xlsx.writeFile(EXCEL_FILE);
        continue;
      }

      await connectBtn.click();
      await sleep(2000);

      // Add personalized note
      const addNoteBtn = await page.$('button:has-text("Add a note")');
      if (addNoteBtn) {
        await addNoteBtn.click();
        await sleep(1500);
        const textarea = await page.$('textarea[name="message"]');
        if (textarea) {
          await textarea.fill(buildNote(getFirstName(name)));
          await sleep(1000);
        }
      }

      // Click Send
      const sendBtn = await page.$('button:has-text("Send")');
      if (sendBtn) {
        await sendBtn.click();
        await sleep(2000);

        const today = new Date().toLocaleDateString("en-IN");
        row.getCell(5).value = today;       // Connection Sent Date
        row.getCell(6).value = "Sent";      // Connection Status
        row.commit();
        await workbook.xlsx.writeFile(EXCEL_FILE);

        console.log(`  ✅ Connection sent to ${name}`);
        sentCount++;
      } else {
        console.log(`  ❌ Send button not found`);
        row.getCell(6).value = "Error - Send btn not found";
        row.commit();
        await workbook.xlsx.writeFile(EXCEL_FILE);
      }

    } catch (err) {
      console.log(`  ❌ Error: ${err.message}`);
      row.getCell(6).value = "Error";
      row.commit();
      await workbook.xlsx.writeFile(EXCEL_FILE);
    }

    // Wait between each person — very important
    await randomDelay(5000, 10000);
  }

  console.log(`\n✅ connector.js done! Sent ${sentCount} connection requests.`);
  await browser.close();
};

run().catch(console.error);



