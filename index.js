const puppeteer = require('puppeteer');
const xlsx = require('xlsx');

// Function to read URLs from an Excel file
function getUrlsFromExcel(filename, sheetName) {
  const workbook = xlsx.readFile(filename);
  const sheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
  
  return data.slice(1).map(row => row[0]).filter(Boolean); // Return an array of URLs
}

// Function to save results into a new Excel file
function saveEmailsToExcel(data, filename) {
  const workbook = xlsx.utils.book_new();
  const worksheetData = [['URL', 'Emails']];
  
  data.forEach(({ url, emails }) => {
    worksheetData.push([url, emails.length > 0 ? emails.join(', ') : 'N/A']);
  });

  const worksheet = xlsx.utils.aoa_to_sheet(worksheetData);
  xlsx.utils.book_append_sheet(workbook, worksheet, 'Emails');
  xlsx.writeFile(workbook, filename);
  console.log(`Emails saved to ${filename}`);
}

async function extractEmailsFromMultiplePages(urls) {
  const browser = await puppeteer.launch();
  const emailResults = [];

  for (const url of urls) {
    const page = await browser.newPage();
    console.log(`Visiting ${url}...`);

    await page.setRequestInterception(true);
    page.on('request', (request) => {
      const resourceType = request.resourceType();
      if (['stylesheet', 'image', 'font', 'script'].includes(resourceType)) {
        request.abort();
      } else {
        request.continue();
      }
    });

    try {
      await page.goto(url, { waitUntil: 'domcontentloaded' });

      const content = await page.content();
      const emailPattern = /mailto:([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})|([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/g;
      const matches = content.matchAll(emailPattern);

      const emails = [];
      for (const match of matches) {
        emails.push(match[1] || match[2]);
      }

      // If no emails found, mark as N/A
      emailResults.push({ url, emails: emails.length > 0 ? emails : ['N/A'] });

    } catch (error) {
      console.error(`Error visiting ${url}: ${error.message}`);
      emailResults.push({ url, emails: ['N/A'] }); // If error occurs, mark as N/A
    } finally {
      await page.close();
    }
  }

  await browser.close();
  return emailResults;
}

// Main execution
const urls = getUrlsFromExcel('websites.xlsx', 'Sheet1');
console.log('Loaded URLs:', urls);  

extractEmailsFromMultiplePages(urls).then((emailResults) => {
  console.log('Emails found on each site:', emailResults);
  saveEmailsToExcel(emailResults, 'scraped_emails.xlsx');  // Save results to new Excel file
});
