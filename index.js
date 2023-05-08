const express = require("express");
const fs = require("fs");
const app = express();
const multer = require("multer");
const path = require("path");
const bodyParser = require("body-parser");
const XLSX = require("xlsx");
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    return cb(null, "./uploads");
  },
  filename: function (req, file, cb) {
    return cb(null, `${Date.now()}-${file.originalname}`);
  },
});
const upload = multer({ storage });
app.use(express.static("public"));
app.use(express.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.set("view engine", "ejs");
app.set("views", "views");
app.get("/", (req, res) => {
  res.render("home");
});
const up = upload.fields([
  {
    name: "excelsheet1",
  },
  {
    name: "excelsheet2",
  },
  {
    name: "excelsheet3",
  },
]);
app.post("/profile", up, (req, res, next) => {
  let pathOfFile1 = path.join(
    __dirname,
    "./uploads",
    req.files.excelsheet1[0].filename
  );
  let pathOfFile2 = path.join(
    __dirname,
    "./uploads",
    req.files.excelsheet2[0].filename
  );
  let pathOfFile3 = path.join(
    __dirname,
    "./uploads",
    req.files.excelsheet3[0].filename
  );
  function adjustAccToNumber(num) {
    if (num / 10000000 > 0) {
      num = Math.round(num / 10000000);
      let a = " Cr";
      a = num + a;
      return a;
    }
  }

  console.log(pathOfFile1);
  const workbook = XLSX.readFile(pathOfFile1);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json(worksheet);
  let arrofWorksheet = jsonData;
  let credit = 0;
  let debit = 0;
  for (let i = 0; i < jsonData.length; i++) {
    if (arrofWorksheet[i]["Credit"] != undefined) {
      credit += arrofWorksheet[i]["Credit"];
    }
    if (arrofWorksheet[i]["Debit"] != undefined) {
      debit += arrofWorksheet[i]["Debit"];
    }
  }
  const workbook1 = XLSX.readFile(pathOfFile2);
  const sheetName1 = workbook1.SheetNames[0];
  const worksheet1 = workbook1.Sheets[sheetName1];
  const jsonData1 = XLSX.utils.sheet_to_json(worksheet1);
  let arrofWorksheet1 = jsonData1;
  let ecash = arrofWorksheet1[0]["Total Ecash"];
  const workbook3 = XLSX.readFile(pathOfFile3);
  const sheetName3 = workbook3.SheetNames[0];
  const worksheet3 = workbook3.Sheets[sheetName3];
  const jsonData3 = XLSX.utils.sheet_to_json(worksheet3);
  let arrofWorksheet3 = jsonData3;
  let credit_of_sheet3 = 0;
  let debit_of_sheet3 = 0;
  for (let i = 0; i < arrofWorksheet3.length; i++) {
    if (arrofWorksheet3[i]["Credit"] != undefined) {
      credit_of_sheet3 += arrofWorksheet3[i]["Credit"];
    }
    if (arrofWorksheet3[i]["Debit"] != undefined) {
      debit_of_sheet3 += arrofWorksheet3[i]["Debit"];
    }
  }
  console.log(credit_of_sheet3);
  debit = debit * -1;

  let totalCredit = credit + credit_of_sheet3;
  let totalDebit = debit_of_sheet3 + debit;

  fs.unlinkSync(pathOfFile1);
  fs.unlinkSync(pathOfFile2);
  fs.unlinkSync(pathOfFile3);

  res.render("display", {
    totalcredit: adjustAccToNumber(totalCredit),
    totaldebit: adjustAccToNumber(totalDebit),
    openingbalance: adjustAccToNumber(
      Math.floor(arrofWorksheet3[0]["Opening Balance"])
    ),
    totalvalueinbothbanks: adjustAccToNumber(
      Math.floor(
        arrofWorksheet3[0]["Opening Balance"] + totalCredit - totalDebit
      )
    ),
    walletbalance: adjustAccToNumber(Math.floor(ecash)),
    discrepancyfound: adjustAccToNumber(
      Math.floor(
        ecash -
          (arrofWorksheet3[0]["Opening Balance"] + totalCredit - totalDebit)
      )
    ),
  });

  next();
});
app.post("/newform", (req, res) => {
  let discrepancyfound = req.body.discrepancyfound;
  console.log(discrepancyfound);
  res.render("kpi");
});
const uploadFiles = upload.fields([
  {
    name: "excelsheet4",
  },
  {
    name: "excelsheet5",
  },
  {
    name: "excelsheet6",
  },
  {
    name: "excelsheet7",
  },
  {
    name: "excelsheet8",
  },
  {
    name: "excelsheet9",
  },
  {
    name: "excelsheet10",
  },
  {
    name: "excelsheet11",
  },
]);
app.post("/kpi", uploadFiles, (req, res, next) => {
  let pathOfFile1 = path.join(
    __dirname,
    "./uploads",
    req.files.excelsheet4[0].filename
  );
  let pathOfFile2 = path.join(
    __dirname,
    "./uploads",
    req.files.excelsheet5[0].filename
  );
  let pathOfFile3 = path.join(
    __dirname,
    "./uploads",
    req.files.excelsheet6[0].filename
  );
  let pathOfFile4 = path.join(
    __dirname,
    "./uploads",
    req.files.excelsheet7[0].filename
  );
  let pathOfFile5 = path.join(
    __dirname,
    "./uploads",
    req.files.excelsheet8[0].filename
  );
  let pathOfFile6 = path.join(
    __dirname,
    "./uploads",
    req.files.excelsheet9[0].filename
  );
  let pathOfFile7 = path.join(
    __dirname,
    "./uploads",
    req.files.excelsheet10[0].filename
  );
  let pathOfFile8 = path.join(
    __dirname,
    "./uploads",
    req.files.excelsheet11[0].filename
  );

  function adjustAccToNumber(num) {
    if (num / 10000000 > 0) {
      num = Math.round(num / 10000000);
      let a = " Cr";
      a = num + a;
      return a;
    }
  }
  console.log(pathOfFile1);
  const workbook = XLSX.readFile(pathOfFile1);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json(worksheet);
  let arrofWorksheet = jsonData;

  const workbook1 = XLSX.readFile(pathOfFile2);
  const sheetName1 = workbook1.SheetNames[0];
  const worksheet1 = workbook1.Sheets[sheetName1];
  const jsonData1 = XLSX.utils.sheet_to_json(worksheet1);
  let arrofWorksheet1 = jsonData1;

  const workbook3 = XLSX.readFile(pathOfFile3);
  const sheetName3 = workbook3.SheetNames[0];
  const worksheet3 = workbook3.Sheets[sheetName3];
  const jsonData3 = XLSX.utils.sheet_to_json(worksheet3);
  let arrofWorksheet3 = jsonData3;

  const workbook4 = XLSX.readFile(pathOfFile4);
  const sheetName4 = workbook4.SheetNames[0];
  const worksheet4 = workbook4.Sheets[sheetName4];
  const jsonData4 = XLSX.utils.sheet_to_json(worksheet4);
  let arrofWorksheet4 = jsonData4;

  const workbook5 = XLSX.readFile(pathOfFile5);
  const sheetName5 = workbook5.SheetNames[0];
  const worksheet5 = workbook5.Sheets[sheetName5];
  const jsonData5 = XLSX.utils.sheet_to_json(worksheet5);
  let arrofWorksheet5 = jsonData5;

  const workbook6 = XLSX.readFile(pathOfFile6);
  const sheetName6 = workbook6.SheetNames[0];
  const worksheet6 = workbook6.Sheets[sheetName6];
  const jsonData6 = XLSX.utils.sheet_to_json(worksheet6);
  let arrofWorksheet6 = jsonData6;

  const workbook7 = XLSX.readFile(pathOfFile7);
  const sheetName7 = workbook7.SheetNames[0];
  const worksheet7 = workbook7.Sheets[sheetName7];
  const jsonData7 = XLSX.utils.sheet_to_json(worksheet7);
  let arrofWorksheet7 = jsonData7;

  const workbook8 = XLSX.readFile(pathOfFile8);
  const sheetName8 = workbook8.SheetNames[0];
  const worksheet8 = workbook8.Sheets[sheetName8];
  const jsonData8 = XLSX.utils.sheet_to_json(worksheet8);
  let arrofWorksheet8 = jsonData8;

  fs.unlinkSync(pathOfFile1);
  fs.unlinkSync(pathOfFile2);
  fs.unlinkSync(pathOfFile3);
  fs.unlinkSync(pathOfFile4);
  fs.unlinkSync(pathOfFile5);
  fs.unlinkSync(pathOfFile6);
  fs.unlinkSync(pathOfFile7);
  fs.unlinkSync(pathOfFile8);
  let arr = [];
  arr.push(arrofWorksheet);
  arr.push(arrofWorksheet1);
  arr.push(arrofWorksheet3);
  arr.push(arrofWorksheet4);
  arr.push(arrofWorksheet5);
  arr.push(arrofWorksheet6);
  arr.push(arrofWorksheet7);
  arr.push(arrofWorksheet8);
  console.log(adjustAccToNumber(13604991340 - 1046374296 + 651732878));
  res.render("abc", {
    prevDescripency: adjustAccToNumber(13604991340),
    totalcredit: adjustAccToNumber(1046374296),
    totalDebit: adjustAccToNumber(651732878),
    totalDescripency: adjustAccToNumber(13604991340 - 1046374296 + 651732878),
  });
  next();
});
app.listen(3000, (err) => {
  console.log(`Your app is running on port 3000`);
});
