const sslCertificate = require("get-ssl-certificate");
const websites = require("./websites.json");
const Excel = require("exceljs");

const todaysDate = new Date();
let workbook = new Excel.Workbook();
const worksheet = workbook.addWorksheet("SSL Certificate");
worksheet.columns = [
  { header: "Website name", key: "website", width: 30 },
  { header: "Cert authority", key: "certificate", width: 30 },
  { header: "Site Target", key: "target", width: 10 },
  { header: "is it Valid?", key: "valid", width: 10 },
  { header: "Kind of website ", key: "type", width: 10 },
  { header: "Method of getting cert info", key: "method", width: 10 },
  { header: "Level of Security", key: "level", width: 10 },
  { header: "Free or Paid", key: "cost", width: 10 },
  { header: "is it in CTL?", key: "status", width: 10 },
];

const getCertificateInfo = async (url) => {
  url = url.match(
    /^(?:https?:)?(?:\/\/)?(?:[^@\n]+@)?(?:www\.)?([^:\/\n]+)/im
  )[1];
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
  for (let i = 0; i < 700; i++) {
    const site = websites.table[i];
    const [certAut, isValid, type] = await getCertificateInfo(site.url);
    console.log(site.url);
    console.log(certAut, isValid);
    worksheet.addRow({
      website: site.url,
      target: site.target,
      certificate: certAut,
      valid: isValid,
      method: "Web Scrapping/TLS",
      type: type,
      cost: "",
      status: "",
      level: "",
    });
  }
  workbook.xlsx
    .writeFile("Phishing_Sites.xlsx")
    .then(() => {
      console.log("Done");
    })
    .catch((err) => {
      console.log(err);
    });
})();
