const Excel = require("exceljs");
const puppeteer = require("puppeteer");
const chilkat = require("@chilkat/ck-node14-macosx");

let workbook = new Excel.Workbook();

const formatUrl = (url) => {
  if (typeof url !== "string") url = url.text;
  url = url.match(
    /^(?:https?:)?(?:\/\/)?(?:[^@\n]+@)?(?:www\.)?([^:\/\n]+)/im
  )[1];
  return url;
};

const getRevocationStatus = async (url) => {
  const http = new chilkat.Http();
  const status = await http.OcspCheck(url, 443);
  if (status === 0) return "Good";
  if (status === 1) return "Revoked";
  if (status === 2) return "Unknown";
  if (status < 0) return "Expired";
  return "N/A";
};

(async () => {
  workbook = await workbook.xlsx.readFile("Phishing_websites.xlsx");
  let worksheet = await workbook.getWorksheet(1);
  let urls = worksheet.getColumn(1).values.slice(2);
  let ctlList = worksheet.getColumn(7);
  let revokedList = worksheet.getColumn(8);

  const browser = await puppeteer.launch({ headless: false });
  const page = await browser.newPage();
  await page.goto(
    "https://transparencyreport.google.com/https/certificates?hl=en"
  );
  let ctlStatus = [];
  let revokedStatus = [];
  let count = 0;

  for (let url of urls) {
    url = formatUrl(url);
    const revocStatus = await getRevocationStatus(url);
    revokedStatus.push(revocStatus);
    await page.type("input.ng-valid", url, { delay: 100 });
    await page.click("search-box i");
    console.log(count);
    await page.waitForTimeout(2000);
    let data = null;
    try {
      data = await page.$eval(
        "tbody tr.google-visualization-table-tr-even",
        (el) => el.innerHTML
      );
    } catch (err) {
      console.log(err);
    }

    console.log("data ", data);
    if (data) {
      ctlStatus.push("yes");
    } else {
      ctlStatus.push("no");
    }
    await page.$eval(`input.ng-valid`, (el) => (el.value = ""));
    count++;
  }

  ctlList.values = [, , ...ctlStatus];
  revokedList.values = [, , ...revokedStatus];

  workbook.xlsx
    .writeFile("new_websites.xlsx")
    .then(() => {
      console.log("Done");
    })
    .catch((err) => {
      console.log(err);
    });
})();
