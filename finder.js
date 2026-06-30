
// const { chromium } = require("playwright");
// const ExcelJS = require("exceljs");
// const { randomDelay } = require("./utils");

// const EXCEL_FILE = "./contacts.xlsx";
// const SHEET_NAME = "Sheet1";

// const findLinkedInUrl = async (page, name, company) => {
//   // Try 3 different search strategies
//   const queries = [
//     `${name} ${company} linkedin`,
//     `${name} linkedin ${company}`,
//     `${name} linkedin India`,
//   ];

//   for (const query of queries) {
//     try {
//       await page.goto(`https://www.bing.com/search?q=${encodeURIComponent(query)}`);
//       await randomDelay(2000, 4000);

//       const links = await page.$$eval("a", els =>
//         els.map(a => a.href).filter(h =>
//           h.includes("linkedin.com/in/") && !h.includes("bing.com")
//         )
//       );

//       if (links.length > 0) {
//         const raw = links[0];
//         const match = raw.match(/https:\/\/[a-z.]*linkedin\.com\/in\/[^&?/]+/);
//         return match ? match[0] : raw;
//       }

//       console.log(`  Tried: "${query}" — no result, trying next...`);
//       await randomDelay(2000, 3000);

//     } catch (err) {
//       console.log(`  Search error: ${err.message}`);
//     }
//   }

//   return null;
// };

// const run = async () => {
//   const workbook = new ExcelJS.Workbook();
//   await workbook.xlsx.readFile(EXCEL_FILE);
//   const sheet = workbook.getWorksheet(SHEET_NAME);

//   const rows = [];
//   sheet.eachRow((row, i) => {
//     if (i === 1) return;
//     rows.push({ row, i });
//   });

//   const browser = await chromium.launchPersistentContext("./profile", {
//     headless: false,
//     viewport: { width: 1280, height: 800 }
//   });
//   const page = await browser.newPage();

//   for (const { row, i } of rows) {
//     const name = (row.getCell(1).value || "").toString().trim();
//     const company = (row.getCell(2).value || "").toString().trim();
//     const existingUrl = row.getCell(3).value;
//     const existingStatus = row.getCell(4).value;

//     // Skip if already found or already marked not found
//     if (!name) {
//       console.log(`Row ${i}: Skipping — no name`);
//       continue;
//     }
//     if (existingUrl) {
//       console.log(`Row ${i}: Skipping ${name} — URL already exists`);
//       continue;
//     }

//     console.log(`\nRow ${i}: Searching for "${name}" @ "${company}"...`);

//     const url = await findLinkedInUrl(page, name, company);

//     if (url) {
//       row.getCell(3).value = url;
//       row.getCell(4).value = "URL Found";
//       console.log(`  ✅ Found: ${url}`);
//     } else {
//       row.getCell(3).value = "";
//       row.getCell(4).value = "URL Not Found";
//       console.log(`  ❌ Not found after all attempts`);
//     }

//     row.commit();
//     await workbook.xlsx.writeFile(EXCEL_FILE);
//     await randomDelay(4000, 7000);
//   }

//   console.log("\n✅ finder.js done! Check your Excel.");
//   await browser.close();
// };

// run().catch(console.error);










// finder.js
// Finds LinkedIn profile URLs for each contact in contacts.xlsx
// Strategy: 1) Search LinkedIn's own people search (most accurate, needs login)
//           2) Fall back to Google search with site:linkedin.com/in

const { chromium } = require("playwright");
const ExcelJS = require("exceljs");
const path = require("path");
const { sleep, randomDelay } = require("./utils");

const EXCEL_FILE = "./contacts.xlsx";
const SHEET_NAME = "Sheet1";
const PROFILE_DIR = path.join(__dirname, "profile"); // absolute path, same as connector.js / messenger.js

// ---- Check if logged into LinkedIn ----
async function isLoggedIn(page) {
  try {
    await page.goto("https://www.linkedin.com/feed", { waitUntil: "domcontentloaded" });
    await page.waitForSelector(".global-nav", { timeout: 8000 });
    return true;
  } catch {
    return false;
  }
}

// ---- Strategy 1: LinkedIn's own people search ----
async function searchLinkedIn(page, name, company) {
  try {
    const query = `${name} ${company}`;
    await page.goto(
      `https://www.linkedin.com/search/results/people/?keywords=${encodeURIComponent(query)}`,
      { waitUntil: "domcontentloaded" }
    );
    await randomDelay(2500, 4500);

    // Grab the first profile link in results
    const links = await page.$$eval('a[href*="/in/"]', els =>
      els.map(a => a.href)
    );

    if (links.length > 0) {
      // Clean up tracking params after the profile slug
      const raw = links[0];
      const match = raw.match(/https:\/\/[a-z.]*linkedin\.com\/in\/[^?/]+/);
      return match ? match[0] : raw.split("?")[0];
    }
    return null;
  } catch (err) {
    console.log(`  LinkedIn search error: ${err.message}`);
    return null;
  }
}

// ---- Strategy 2: Google fallback ----
async function searchGoogle(page, name, company) {
  const queries = [
    `"${name}" "${company}" site:linkedin.com/in`,
    `${name} ${company} linkedin`,
  ];

  for (const query of queries) {
    try {
      await page.goto(`https://www.google.com/search?q=${encodeURIComponent(query)}`, {
        waitUntil: "domcontentloaded",
      });
      await randomDelay(2500, 4500);

      const links = await page.$$eval("a", els =>
        els.map(a => a.href).filter(h => h.includes("linkedin.com/in/"))
      );

      if (links.length > 0) {
        const raw = links[0];
        const match = raw.match(/https:\/\/[a-z.]*linkedin\.com\/in\/[^&?/]+/);
        return match ? match[0] : raw;
      }

      console.log(`  Google: "${query}" — no result, trying next...`);
      await randomDelay(2000, 3000);
    } catch (err) {
      console.log(`  Google search error: ${err.message}`);
    }
  }
  return null;
}

const run = async () => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(EXCEL_FILE);
  const sheet = workbook.getWorksheet(SHEET_NAME);

  const rows = [];
  sheet.eachRow((row, i) => {
    if (i === 1) return;
    rows.push({ row, i });
  });

  const browser = await chromium.launchPersistentContext(PROFILE_DIR, {
    headless: false,
    viewport: { width: 1280, height: 800 },
  });

  try {
    const page = await browser.newPage();

    // ---- Login check ----
    console.log("Checking LinkedIn login...");
    let loggedIn = await isLoggedIn(page);

    if (!loggedIn) {
      console.log("⚠️  Not logged into LinkedIn in this browser profile.");
      console.log("Please log in manually in the window that opened, then press ENTER here...");
      await new Promise(r => process.stdin.once("data", r));
      loggedIn = await isLoggedIn(page);
    }

    if (loggedIn) {
      console.log("✅ Logged in! LinkedIn search will be used as primary method.\n");
    } else {
      console.log("⚠️  Still not logged in — will use Google search only (less accurate).\n");
    }

    // ---- Main loop ----
    for (const { row, i } of rows) {
      const name = (row.getCell(1).value || "").toString().trim();
      const company = (row.getCell(2).value || "").toString().trim();
      const existingUrl = row.getCell(3).value;

      if (!name) {
        console.log(`Row ${i}: Skipping — no name`);
        continue;
      }
      if (existingUrl) {
        console.log(`Row ${i}: Skipping ${name} — URL already exists`);
        continue;
      }

      console.log(`\nRow ${i}: Searching for "${name}" @ "${company}"...`);

      let url = null;
      let source = "";

      // Try LinkedIn first (only if logged in)
      if (loggedIn) {
        url = await searchLinkedIn(page, name, company);
        if (url) source = "LinkedIn";
      }

      // Fall back to Google
      if (!url) {
        url = await searchGoogle(page, name, company);
        if (url) source = "Google";
      }

      if (url) {
        row.getCell(3).value = url;
        row.getCell(4).value = "URL Found";
        row.getCell(9).value = `Found via ${source}`;
        console.log(`  ✅ Found via ${source}: ${url}`);
      } else {
        row.getCell(3).value = "";
        row.getCell(4).value = "URL Not Found";
        console.log(`  ❌ Not found via LinkedIn or Google`);
      }

      row.commit();
      try {
        await workbook.xlsx.writeFile(EXCEL_FILE);
      } catch (err) {
        console.log(`  ⚠️  Could not save Excel (is it open in another program?): ${err.message}`);
      }
 
      await randomDelay(4000, 7000);
    }

    console.log("\n✅ finder.js done! Check your Excel.");
  } finally {
    await browser.close();
  }
};

run().catch(console.error);