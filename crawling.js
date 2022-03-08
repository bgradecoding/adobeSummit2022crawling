import axios from "axios";
import cheerio from "cheerio";
import puppeteer from "puppeteer";
import Excel from "exceljs";

let workbook = new Excel.Workbook();
let worksheet = workbook.addWorksheet("AdobeSummit");
worksheet.columns = [
  { header: "Session Title", key: "sessionTitle" },
  { header: "Session Track", key: "sessionTrack" },
  { header: "Session Product", key: "sessionProduct" },
  { header: "Session Schedule 1", key: "sessionSchedule1" },
  { header: "Session Schedule 2", key: "sessionSchedule2" },
  { header: "Session Schedule 3", key: "sessionSchedule3" },
];

worksheet.columns.forEach((column) => {
  column.width = column.header.length < 12 ? 12 : column.header.length;
});
worksheet.getRow(1).font = { bold: true };

const getHtml = async () => {
  const browser = await puppeteer.launch({
    headless: false,
  });

  // 새로운 페이지를 연다.
  const page = await browser.newPage();
  // 페이지의 크기를 설정한다.
  await page.setViewport({
    width: 1366,
    height: 768,
  });
  await page.goto("https://portal.adobe.com/widget/adobe/as22/sessions");
  // 페이지의 HTML을 가져온다.
  setTimeout(async () => {
    await page.focus("button.mdBtnR.mdBtnR-primary.show-more-btn");
    await page.keyboard.type("\n");

    await page.waitForTimeout(5000);

    await page.focus("button.mdBtnR.mdBtnR-primary.show-more-btn");
    await page.keyboard.type("\n");

    await page.waitForTimeout(5000);
    await page.focus("button.mdBtnR.mdBtnR-primary.show-more-btn");
    await page.keyboard.type("\n");

    await page.waitForTimeout(5000);
    await page.focus("button.mdBtnR.mdBtnR-primary.show-more-btn");
    await page.keyboard.type("\n");
    await page.waitForTimeout(5000);

    const html = await page.content();

    const $ = cheerio.load(html);
    const lists = $("li.catalog-result.session-result.show-session-title-icon");

    lists.each((index, list) => {
      const sessionTitle = $(list)
        .find("div.catalog-result-title-text > button > div")
        .text();
      const sessionTrack = $(list)
        .find("div.attribute-Track > span.attribute-values")
        .text();
      const sessionProduct = $(list)
        .find("div.attribute-Product > span.attribute-values")
        .text();
      const sessionSchedule1 =
        $(list)
          .find("ul.session-actions > li")
          .eq(0)
          .find("span.session-date")
          .text() +
        " " +
        $(list)
          .find("ul.session-actions > li")
          .eq(0)
          .find("span.session-time")
          .text();
      const sessionSchedule2 =
        $(list)
          .find("ul.session-actions > li")
          .eq(1)
          .find("span.session-date")
          .text() +
        " " +
        $(list)
          .find("ul.session-actions > li")
          .eq(1)
          .find("span.session-time")
          .text();
      const sessionSchedule3 =
        $(list)
          .find("ul.session-actions > li")
          .eq(2)
          .find("span.session-date")
          .text() +
        " " +
        $(list)
          .find("ul.session-actions > li")
          .eq(2)
          .find("span.session-time")
          .text();

      worksheet.addRow({
        sessionTitle: sessionTitle,
        sessionTrack: sessionTrack,
        sessionProduct: sessionProduct,
        sessionSchedule1: sessionSchedule1,
        sessionSchedule2: sessionSchedule2,
        sessionSchedule3: sessionSchedule3,
      });
    });
    workbook.xlsx.writeFile("AdobeSummit.xlsx");
  }, 10000);
};

getHtml();
