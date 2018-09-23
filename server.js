const express = require("express");
const app = express();
const port = 3000;
const fs = require("fs");
const path = require("path");
const conversionFactory = require("html-to-xlsx");
const puppeteer = require("puppeteer");
const chromeEval = require("chrome-page-eval")({ puppeteer });
const conversion = conversionFactory({
  extract: chromeEval
});

app.get("/getExcel", (req, res) => {
  let html = fs.readFileSync(path.resolve(__dirname, "test.html"), "utf8");
  (async () => {
    const stream = await conversion(html);
    res.setHeader("Content-disposition", "attachment; filename=excel.xlsx");
    res.setHeader(
      "Content-type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    stream.pipe(res);
  })();
});

app.listen(port, () => console.log(`Example app listening on port ${port}!`));
