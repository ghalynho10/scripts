const puppeteer = require("puppeteer");
const fs = require("fs");
const sslCertificate = require("get-ssl-certificate");
const Exceljs = require("exceljs");

const todaysDate = new Date();

let workbook = new Exceljs.Workbook();
let worksheet = workbook.addWorksheet("Sheet 1");
worksheet.columns = [
  { header: "URL", key: "url", width: 30 },
  { header: "Certificate", key: "certificate", width: 30 },
  { header: "Validation Level", key: "validation", width: 10 },
];

const getCertificateInfo = async (url) => {
  // url = url.match(
  //   /^(?:https?:)?(?:\/\/)?(?:[^@\n]+@)?(?:www\.)?([^:\/\n]+)/im
  // )[1];
  let certificate = [];
  let type = "";
  try {
    const infos = await sslCertificate.get(url);
    if (infos.subject.O && infos.subject.businessCategory) {
      type = "EV";
    } else if (infos.subject.O) {
      type = "OV";
    } else {
      type = "DV";
    }
    certificate = await [
      `${infos.issuer.O}, ${infos.issuer.CN}`,
      new Date(infos.valid_to) > todaysDate ? "Valid" : "Expired",
      type,
    ];
  } catch (error) {
    certificate = ["No certificate", "No Certificate", "N/A"];
    console.log(error);
  }
  return certificate;
};

(async () => {
  const browser = await puppeteer.launch({ headless: false });
  const page = await browser.newPage();
  await page.goto("https://www.alexa.com/topsites");

  const sites = await page.$$eval(".DescriptionCell p a", (options) =>
    options.map((option) => option.innerHTML)
  );

  console.log(sites);
  for (let urls of sites) {
    const [certAut, isValid, type] = await getCertificateInfo(urls);
    console.log(urls);
    console.log(certAut, isValid);
    worksheet.addRow({
      url: urls,
      certificate: certAut,
      validation: type,
    });
    workbook.xlsx
      .writeFile("Valid_Sites.xlsx")
      .then(() => {
        console.log("Done");
      })
      .catch((err) => {
        console.log(err);
      });
  }
})();
