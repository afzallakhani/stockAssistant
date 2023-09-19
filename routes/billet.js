const express = require("express");
const router = express.Router();
const path = require("path");
const docx = require("docx");
const { promisify } = require("util");
const puppeteer = require("puppeteer");
const { urlencoded, query, json } = require("express");
// const unlinkAsync = promisify(fs.unlink);
const catchAsync = require("../utils/catchAsync");
const methodOverride = require("method-override");
const multer = require("multer");
const fs = require("fs");
require("dotenv/config");
const ExpressError = require("../utils/ExpressError");
const Party = require("../models/partyMaster");
const Items = require("../models/elafStock");
const Tc = require("../models/billetTc");

const Billets = require("../models/billetList");
const ItemCategories = require("../models/itemCategories");
const Images = require("../models/images");
const multerStorage = require("../utils/multerStorage");
const bodyParser = require("body-parser");

const validateItem = require("../utils/validateItem");
const events = require("events");
const billetTc = require("../models/billetTc");
const eventEmitter = new events.EventEmitter();

let upload = multer({ storage: multerStorage });
// router.get("/list", (req, res) => {
//   res.render("billets/list");
// });

router.get(
  "/list",
  catchAsync(async (req, res) => {
    const list = await Billets.find({});
    console.log(list);
    res.render("billets/list", { list });
  })
);
router.get("/new", (req, res) => {
  res.render("billets/new");
});

// searched Heat display
router.get("/search", async (req, res) => {
  const queryString = req.query.heat;
  let query = queryString.heatNo.toString();
  let queryList = query.split(" ");
  // console.log(queryList);
  // console.log(req.query.heat);

  // for (let query of queryString) {
  //     //     console.log(typeof queryString);
  //     console.log(query.itemName);
  // }

  const heatss = await Billets.find({ heatNo: { $in: queryList } });
  // console.log(heatss);

  res.render("billets/search", { heatss });
});

router.get(
  "/newTc",
  catchAsync(async (req, res, next) => {
    const party = await Party.find({ partyType: { $in: "Buyer" } });
    const list = await Billets.find({});
    const tc = await Tc.find({});
    // console.log(tc);
    res.render("billets/newTc", { list, party });
  })
);
router.post(
  "/newTc",
  catchAsync(async (req, res, next) => {
    // console.log(req.body.tc);
    let tc = new Tc(req.body.tc);

    let queryList = req.body.billet;
    let heatArray = Object.values(queryList);
    let heatQuery = heatArray.toString().split(" ");
    console.log(heatQuery);
    let heats = await Billets.find({ heatNo: { $in: heatQuery } });

    console.log(heats);

    const heatId = heats.map((item) => item._id);
    tc.heatNo = heatId;
    // console.log(item););
    console.log(tc);
    await tc.save();
    let = currentTc = await Tc.findById(tc.id);
    // currentTc.heatNo.push(heats);
    // await tc.save();
    // let = finalTc = await Tc.findById(tc.id);

    console.log(currentTc);
    // console.log(finalTc);
    // let newTc = Tc.find({});
    // console.log(newTc);
    res.redirect("/billets/newTc");
  })
);

// EDIT TC VIEW ROUTE
router.get(
  "/:id/editTc",
  catchAsync(async (req, res, next) => {
    const tc = await Tc.findById(req.params.id).populate("heatNo");
    const party = await Party.find({ partyType: { $in: "Buyer" } });
    console.log(party);
    res.render("billets/editTc", { tc, party });
  })
);

// UPDATE TC
// router.put(
//   "/:id",
//   catchAsync(async (req, res, next) => {
//     const { id } = req.params;
//     await Tc.findByIdAndUpdate(id, { ...req.body.billet });
//     res.redirect("/billets/tcList");
//   })
// );
// DELETE TC
router.delete(
  "/:id",
  catchAsync(async (req, res) => {
    const { id } = req.params;
    console.log("deleted Heat");
    await Tc.findByIdAndDelete(id);
    res.redirect("/billets/tcList");
  })
);

// const generatePdf = async (req, res) => {
//   try {
//     const browser = await puppeteer.launch();
//     const page = await browser.newPage();
//     const url =
//       `${req.protocol}://${req.get("host")}` +
//       "/billets/" +
//       `${req.params.id}` +
//       "/tcPdf";

//     console.log(url);
//     await page.goto(
//       `${req.protocol}://${req.get("host")}` +
//         "/billets/" +
//         `${req.params.id}` +
//         "/tcPdf",
//       {
//         waitUntil: "networkidle2",
//       }
//     );

//     await page.setViewport({ width: 1980, height: 1050 });
//     const todayDate = new Date();
//     const pdfn = await page.pdf({
//       path: `${path.join(__dirname, "../files", todayDate.getTime() + ".pdf")}`,
//       format: "a4",
//     });
//     await browser.close();
//     const pdfURL = path.join(
//       __dirname,
//       "../files",
//       todayDate.getTime() + ".pdf"
//     );
//     res.set({
//       "Content-Type": "application/pdf",
//       "Content-Length": pdfn.length,
//     });
//     res.sendFile(pdfURL);
//   } catch (error) {
//     console.log(error.message);
//   }
// };
// TC PREVIEW GET
router.get(
  "/:id/tcPreview",
  catchAsync(async (req, res, next) => {
    const tc = await Tc.findById(req.params.id).populate("heatNo");
    const tcList = await Tc.find({});
    const billet = await Billets.find({});
    console.log(tc);
    res.render("billets/tcPreview", { tc, tcList });
  })
);
// TC READY FOR MAKING PDF OPEN CASTING
router.get(
  "/:id/tcPdf",
  catchAsync(async (req, res, next) => {
    const tcId = req.params.id;

    const tc = await Tc.findById(req.params.id).populate("heatNo");
    const tcList = await Tc.find({});
    const billet = await Billets.find({});
    // console.log(tc);

    res.render("billets/tcPdf", { tc, tcList });
  })
);
// TC READY FOR MAKING PDF CLOSE CASTING
router.get(
  "/:id/closeCastTcPdf",
  catchAsync(async (req, res, next) => {
    const tcId = req.params.id;

    const tc = await Tc.findById(req.params.id).populate("heatNo");
    const tcList = await Tc.find({});
    const billet = await Billets.find({});
    // console.log(tc);

    res.render("billets/closeCastTcPdf", { tc, tcList });
  })
);
router.get(
  "/:id/genratePdf",
  catchAsync(async (req, res) => {
    console.log("hi");
    try {
      console.log("hi");
      const browser = await puppeteer.launch({ headless: true });
      const page = await browser.newPage();
      const url =
        `${req.protocol}://${req.get("host")}` +
        "/billets/" +
        `${req.params.id}` +
        "/tcPdf";

      console.log(url);
      await page.goto(
        `${req.protocol}://${req.get("host")}` +
          "/billets/" +
          `${req.params.id}` +
          "/tcPdf",
        {
          waitUntil: "networkidle2",
        }
      );

      await page.setViewport({ width: 1600, height: 1050 });
      const todayDate = new Date();
      const pdfn = await page.pdf({
        path: `${path.join(
          __dirname,
          "../files",
          todayDate.getTime() + ".pdf"
        )}`,
        format: "a4",
      });
      await browser.close();
      const pdfURL = path.join(
        __dirname,
        "../files",
        todayDate.getTime() + ".pdf"
      );
      res.set({
        "Content-Type": "application/pdf",
        "Content-Length": pdfn.length,
      });
      res.sendFile(pdfURL);
    } catch (error) {
      console.log(error.message);
    }
  })
);
router.get(
  "/:id/genratePdfCc",
  catchAsync(async (req, res) => {
    console.log("hi");
    try {
      console.log("hi");
      const browser = await puppeteer.launch({ headless: true });
      const page = await browser.newPage();
      const url =
        `${req.protocol}://${req.get("host")}` +
        "/billets/" +
        `${req.params.id}` +
        "/closeCastTcPdf";
      await page.goto(
        `${req.protocol}://${req.get("host")}` +
          "/billets/" +
          `${req.params.id}` +
          "/closeCastTcPdf",
        {
          waitUntil: "networkidle2",
        }
      );

      await page.setViewport({ width: 1800, height: 1050 });
      const todayDate = new Date();
      const pdfn = await page.pdf({
        path: `${path.join(
          __dirname,
          "../files",
          todayDate.getTime() + ".pdf"
        )}`,
        format: "a3",
        margin: { right: "0px", left: "0px" },
      });
      await browser.close();
      const pdfURL = path.join(
        __dirname,
        "../files",
        todayDate.getTime() + ".pdf"
      );
      res.set({
        "Content-Type": "application/pdf",
        "Content-Length": pdfn.length,
      });
      res.sendFile(pdfURL);
    } catch (error) {
      console.log(error.message);
    }
  })
);
// GENERATE PDF OPEN CASTING
router.get(
  "/:id/genratePdf",
  catchAsync(async (req, res) => {
    console.log("hi");
    try {
      console.log("hi");
      const browser = await puppeteer.launch({ headless: true });
      const page = await browser.newPage();
      const url =
        `${req.protocol}://${req.get("host")}` +
        "/billets/" +
        `${req.params.id}` +
        "/tcPdf";

      console.log(url);
      await page.goto(
        `${req.protocol}://${req.get("host")}` +
          "/billets/" +
          `${req.params.id}` +
          "/tcPdf",
        {
          waitUntil: "networkidle2",
        }
      );

      await page.setViewport({ width: 1600, height: 1050 });
      const todayDate = new Date();
      const pdfn = await page.pdf({
        path: `${path.join(
          __dirname,
          "../files",
          todayDate.getTime() + ".pdf"
        )}`,
        format: "a4",
      });
      await browser.close();
      const pdfURL = path.join(
        __dirname,
        "../files",
        todayDate.getTime() + ".pdf"
      );
      res.set({
        "Content-Type": "application/pdf",
        "Content-Length": pdfn.length,
      });
      res.sendFile(pdfURL);
    } catch (error) {
      console.log(error.message);
    }
  })
);
router.get(
  "/:id/tcPreview/closeCast",
  catchAsync(async (req, res, next) => {
    // try {
    //   const browser = await puppeteer.launch();
    //   const page = await browser.newPage();
    //   await page.goto(
    //     `${req.protocol}://${req.get("host")}` + "/views/billets/:id/tcPdf",
    //     {
    //       waitUntil: "networkidle2",
    //     }
    //   );
    //   await page.setViewport({ width: 1680, height: 1050 });
    //   const todayDate = new Date();
    //   const pdfn = await page.pdf({
    //     path: `${path.join(
    //       __dirname,
    //       "../files",
    //       todayDate.getTime() + ".pdf"
    //     )}`,
    //     format: "a4",
    //   });
    //   await browser.close();
    //   const pdfURL = path.join(
    //     __dirname,
    //     "../files",
    //     todayDate.getTime() + ".pdf"
    //   );
    //   res.set({
    //     "Content-Type": "application/pdf",
    //     "Content-Length": pdfn.length,
    //   });
    //   res.sendFile(pdfURL);
    // } catch (error) {
    //   console.log(error.message);
    // }

    const tc = await Tc.findById(req.params.id).populate("heatNo");
    const tcList = await Tc.find({});
    const billet = await Billets.find({});
    console.log(tc);
    res.render("billets/closeCastTcPreview", { tc, tcList });
  })
);

// Function for generating pdf
const g = async (req, res) => {
  try {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    await page.goto(
      `${req.protocol}://${req.get("host")}` + "/views/billets/tcPdf.ejs",
      {
        waitUntil: "networkidle2",
      }
    );
    await page.setViewport({ width: 1680, height: 1050 });
    const todayDate = new Date();
    const pdfn = await page.pdf({
      path: `${path.join(__dirname, "../files", todayDate.getTime() + ".pdf")}`,
      format: "a4",
    });
    await browser.close();
    const pdfURL = path.join(
      __dirname,
      "../files",
      todayDate.getTime() + ".pdf"
    );
    res.set({
      "Content-Type": "application/pdf",
      "Content-Length": pdfn.length,
    });
    res.sendFile(pdfURL);
  } catch (error) {
    console.log(error.message);
  }
};
// TC PDF
// router.get(
//   "/tcPdf",
//   catchAsync(async (req, res, next) => {
//     const tc = await Tc.findById(req.params.id).populate("heatNo");
//     const billet = await Billets.find({});
//     console.log(tc);
//     res.redirect("/billets/tcPdf");
//   })
// );
router.get(
  "/tcList",
  catchAsync(async (req, res, next) => {
    const tcList = await Tc.find({}).populate("heatNo");
    const billet = await Billets.find({});
    console.log(tcList);
    res.render("billets/tcList", { tcList });
  })
);
// SEARCH TC
router.get(
  "/tcList/search",
  catchAsync(async (req, res, next) => {
    const queryString = req.query.tc;
    let query = queryString.tcNo.toString();
    let queryList = query.split(" ");
    // console.log(queryList);
    // console.log(req.query.heat);

    // for (let query of queryString) {
    //     //     console.log(typeof queryString);
    //     console.log(query.itemName);
    // }

    const tcResults = await Tc.find({ tcNo: { $in: queryList } });
    // console.log(heatss);
    console.log(tcResults);

    res.render("billets/tcList", { tcResults });
  })
);
router.get(
  "/:id/edit",
  catchAsync(async (req, res, next) => {
    const billet = await Billets.findById(req.params.id);
    res.render("billets/edit", { billet });
  })
);
router.post(
  "/",
  upload.fields([]),
  catchAsync(async (req, res, next) => {
    // if (!req.body.item) throw new ExpressError("Invalid Item Data", 400);

    let billet = new Billets(req.body.billet);
    console.log(req.body.billet.productionDate);
    await billet.save();
    const newBillet = Billets.find({});
    console.log(newBillet);
    res.redirect("/billets/list");
  })
);

router.put(
  "/:id",
  catchAsync(async (req, res, next) => {
    const { id } = req.params;
    await Billets.findByIdAndUpdate(id, { ...req.body.billet });
    res.redirect("/billets/list");
  })
);
// DELETE HEAT
router.delete(
  "/:id",
  catchAsync(async (req, res) => {
    const { id } = req.params;
    console.log("deleted Heat");
    await Billets.findByIdAndDelete(id);
    res.redirect("/billets/list");
  })
);
module.exports = router;
