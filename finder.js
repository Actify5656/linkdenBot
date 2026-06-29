// finder.js
const { chromium } = require("playwright");
const ExcelJS = require("exceljs");
const { randomDelay } = require("./utils");

const EXCEL_FILE = "./contacts.xlsx";
const SHEET_NAME = "Sheet1";

const run = async () => {
  // Load Excel
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(EXCEL_FILE);
  const sheet = workbook.getWorksheet(SHEET_NAME);

  // Collect rows first (can't do async inside eachRow properly)
  const rows = [];
  sheet.eachRow((row, i) => {
    if (i === 1) return; // skip header
    rows.push({ row, i });
  });

  // Launch browser with persistent profile (so login saves)
  const browser = await chromium.launchPersistentContext("./profile", {
    headless: false,
    viewport: { width: 1280, height: 800 }
  });
  const page = await browser.newPage();

  for (const { row, i } of rows) {
    const name = row.getCell(1).value;
    const company = row.getCell(2).value;
    const existingUrl = row.getCell(3).value;

    // Skip if no name, or URL already found
    if (!name || existingUrl) {
      console.log(`Row ${i}: Skipping — already has URL or no name`);
      continue;
    }

    console.log(`\nRow ${i}: Searching for ${name} @ ${company}...`);

    try {
      // Search Google
      const query = `"${name}" "${company}" site:linkedin.com/in`;
      await page.goto(`https://www.google.com/search?q=${encodeURIComponent(query)}`);
      await randomDelay(2000, 4000);

      // Get all links from Google results
      const links = await page.$$eval("a", els =>
        els.map(a => a.href).filter(h => h.includes("linkedin.com/in/"))
      );

      if (links.length > 0) {
        // Clean the URL (remove Google redirect wrapper)
        const raw = links[0];
        const match = raw.match(/https:\/\/[a-z.]*linkedin\.com\/in\/[^&?/]+/);
        const cleanUrl = match ? match[0] : raw;

        row.getCell(3).value = cleanUrl;
        row.getCell(4).value = "URL Found";
        console.log(`  ✅ Found: ${cleanUrl}`);
      } else {
        row.getCell(3).value = "";
        row.getCell(4).value = "URL Not Found";
        console.log(`  ❌ Not found`);
      }

      row.commit();
      // Save after every row so data is not lost if script crashes
      await workbook.xlsx.writeFile(EXCEL_FILE);

    } catch (err) {
      console.log(`  ❌ Error on row ${i}: ${err.message}`);
      row.getCell(4).value = "Error";
      row.commit();
      await workbook.xlsx.writeFile(EXCEL_FILE);
    }

    // Wait between searches — avoid Google blocking
    await randomDelay(4000, 8000);
  }

  console.log("\n✅ finder.js done! Check your Excel.");
  await browser.close();
};

run().catch(console.error);


