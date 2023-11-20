const express = require("express");
const router = express.Router();
const path = require("path");
const officegen = require("officegen");
// const docx = officegen("docx");

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
const { date } = require("joi");
const eventEmitter = new events.EventEmitter();

let upload = multer({ storage: multerStorage });
// router.get("/list", (req, res) => {
//   res.render("billets/list");
// });

// Documents contain sections, you can have multiple sections per document, go here to learn more about sections
// This simple example will only contain one section

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
  "/:id/tc",
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

// TC PREVIEW GET
router.get(
  "/:id/tcPreview",
  catchAsync(async (req, res, next) => {
    const tc = await Tc.findById(req.params.id).populate("heatNo");
    let tcHeats = tc.heatNo;
    let heats = tcHeats.map(({ _id, heatNo }) => ({
      _id,
      heatNo,
    }));
    // let heats = tcHeats.map();
    const tcList = await Tc.find({});
    const billet = await Billets.find({});
    // console.log(tc);
    let heatId = heats.map((heats) => heats._id);

    let heatData = await Billets.find({ _id: { $in: heatId } });
    console.log(heatData[0].heatNo);

    if (heatData.length == 1) {
      const doc1 = new docx.Document({
        sections: [
          {
            children: [
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "TEST CERTIFICATE FOR CARBON/ALLOY STEEL BILLETS FOR RE-ROLLING INTO GENERAL/STRUCTURAL PURPOSES.",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "It is certified that the material described below fully confirms to IS: 2830: 2012 Chemical Composition tested in accordance with scheme of testing and inspection contained in BIS certification marks License.",
                    font: "Calibri",
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "License Serial No. - ES/BIS/2830/7600107011.",
                    font: "Calibri",
                    bold: true,
                    size: 22,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Inv. No. ${tc.billNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Date - ${tc.tcDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. No. ${tc.poNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. Date - ${tc.poDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Total Qty. - ${tc.totalQtyMts}0`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Vehicle No. - ${tc.vehicleNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        columnSpan: 2,
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                text: `Total Pcs. - ${tc.totalPcs}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
                width: {
                  size: 5000,
                  type: docx.WidthType.DXA,
                },
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `BUYER :- ${tc.buyerName}`,
                    font: "Calibri",
                    bold: true,
                    size: 24,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `CHEMICAL COMPOSITION`,
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Size MM",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Heat No.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Qty. Pcs.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "C %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Mn %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "P %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "S %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Si %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Grade",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Color Code",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    height: { value: 50 },
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    text: "PHYSICAL PARAMETERS",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "All Physical Properties are as per & within IS 2830:2012 Requirements & Tolerances.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "The material is free from Radioactive Contamination, Lead, Mercury and Surface Defects.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 1,
                  }),
                  new docx.TextRun({
                    text: "Process Route - EIF-CCM",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "                                                                                                                                                      Q.C.",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.RIGHT,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "ELAF STEEL",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
            ],
          },
        ],
      });
      docx.Packer.toBuffer(doc1).then((buffer) => {
        fs.writeFileSync(
          `${tc.buyerName} ${tc.billNo} ${heatData[0].sectionSize}MM.docx`,
          buffer
        );
      });
    } else if (heatData.length == 2) {
      const doc2 = new docx.Document({
        sections: [
          {
            children: [
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "TEST CERTIFICATE FOR CARBON/ALLOY STEEL BILLETS FOR RE-ROLLING INTO GENERAL/STRUCTURAL PURPOSES.",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "It is certified that the material described below fully confirms to IS: 2830: 2012 Chemical Composition tested in accordance with scheme of testing and inspection contained in BIS certification marks License.",
                    font: "Calibri",
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "License Serial No. - ES/BIS/2830/7600107011.",
                    font: "Calibri",
                    bold: true,
                    size: 22,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Inv. No. ${tc.billNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Date - ${tc.tcDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. No. ${tc.poNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. Date - ${tc.poDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Total Qty. - ${tc.totalQtyMts}0`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Vehicle No. - ${tc.vehicleNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        columnSpan: 2,
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                text: `Total Pcs. - ${tc.totalPcs}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
                width: {
                  size: 5000,
                  type: docx.WidthType.DXA,
                },
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `BUYER :- ${tc.buyerName}`,
                    font: "Calibri",
                    bold: true,
                    size: 24,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `CHEMICAL COMPOSITION`,
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Size MM",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Heat No.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Qty. Pcs.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "C %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Mn %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "P %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "S %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Si %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Grade",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Color Code",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    text: "PHYSICAL PARAMETERS",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "All Physical Properties are as per & within IS 2830:2012 Requirements & Tolerances.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "The material is free from Radioactive Contamination, Lead, Mercury and Surface Defects.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 1,
                  }),
                  new docx.TextRun({
                    text: "Process Route - EIF-CCM",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "                                                                                                                                                      Q.C.",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.RIGHT,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "ELAF STEEL",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
            ],
          },
        ],
      });
      docx.Packer.toBuffer(doc2).then((buffer) => {
        fs.writeFileSync(
          `${tc.buyerName} ${tc.billNo} ${heatData[0].sectionSize}MM.docx`,
          buffer
        );
      });
    } else if (heatData.length == 3) {
      const doc3 = new docx.Document({
        sections: [
          {
            children: [
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "TEST CERTIFICATE FOR CARBON/ALLOY STEEL BILLETS FOR RE-ROLLING INTO GENERAL/STRUCTURAL PURPOSES.",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "It is certified that the material described below fully confirms to IS: 2830: 2012 Chemical Composition tested in accordance with scheme of testing and inspection contained in BIS certification marks License.",
                    font: "Calibri",
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "License Serial No. - ES/BIS/2830/7600107011.",
                    font: "Calibri",
                    bold: true,
                    size: 22,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Inv. No. ${tc.billNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Date - ${tc.tcDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. No. ${tc.poNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. Date - ${tc.poDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Total Qty. - ${tc.totalQtyMts}0`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Vehicle No. - ${tc.vehicleNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        columnSpan: 2,
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                text: `Total Pcs. - ${tc.totalPcs}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
                width: {
                  size: 5000,
                  type: docx.WidthType.DXA,
                },
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `BUYER :- ${tc.buyerName}`,
                    font: "Calibri",
                    bold: true,
                    size: 24,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `CHEMICAL COMPOSITION`,
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Size MM",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Heat No.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Qty. Pcs.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "C %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Mn %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "P %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "S %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Si %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Grade",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Color Code",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    text: "PHYSICAL PARAMETERS",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "All Physical Properties are as per & within IS 2830:2012 Requirements & Tolerances.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "The material is free from Radioactive Contamination, Lead, Mercury and Surface Defects.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 1,
                  }),
                  new docx.TextRun({
                    text: "Process Route - EIF-CCM",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "                                                                                                                                                      Q.C.",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.RIGHT,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "ELAF STEEL",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
            ],
          },
        ],
      });
      docx.Packer.toBuffer(doc3).then((buffer) => {
        fs.writeFileSync(
          `${tc.buyerName} ${tc.billNo} ${heatData[0].sectionSize}MM.docx`,
          buffer
        );
      });
    } else if (heatData.length == 4) {
      const doc4 = new docx.Document({
        sections: [
          {
            children: [
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "TEST CERTIFICATE FOR CARBON/ALLOY STEEL BILLETS FOR RE-ROLLING INTO GENERAL/STRUCTURAL PURPOSES.",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "It is certified that the material described below fully confirms to IS: 2830: 2012 Chemical Composition tested in accordance with scheme of testing and inspection contained in BIS certification marks License.",
                    font: "Calibri",
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "License Serial No. - ES/BIS/2830/7600107011.",
                    font: "Calibri",
                    bold: true,
                    size: 22,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Inv. No. ${tc.billNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Date - ${tc.tcDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. No. ${tc.poNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. Date - ${tc.poDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Total Qty. - ${tc.totalQtyMts}0`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Vehicle No. - ${tc.vehicleNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        columnSpan: 2,
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                text: `Total Pcs. - ${tc.totalPcs}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
                width: {
                  size: 5000,
                  type: docx.WidthType.DXA,
                },
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `BUYER :- ${tc.buyerName}`,
                    font: "Calibri",
                    bold: true,
                    size: 24,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `CHEMICAL COMPOSITION`,
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Size MM",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Heat No.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Qty. Pcs.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "C %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Mn %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "P %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "S %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Si %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Grade",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Color Code",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    height: { value: 50 },
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    text: "PHYSICAL PARAMETERS",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "All Physical Properties are as per & within IS 2830:2012 Requirements & Tolerances.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "The material is free from Radioactive Contamination, Lead, Mercury and Surface Defects.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 1,
                  }),
                  new docx.TextRun({
                    text: "Process Route - EIF-CCM",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "                                                                                                                                                      Q.C.",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.RIGHT,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "ELAF STEEL",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
            ],
          },
        ],
      });
      docx.Packer.toBuffer(doc4).then((buffer) => {
        fs.writeFileSync(
          `${tc.buyerName} ${tc.billNo} ${heatData[0].sectionSize}MM.docx`,
          buffer
        );
      });
    } else if (heatData.length == 5) {
      const doc5 = new docx.Document({
        sections: [
          {
            children: [
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "TEST CERTIFICATE FOR CARBON/ALLOY STEEL BILLETS FOR RE-ROLLING INTO GENERAL/STRUCTURAL PURPOSES.",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "It is certified that the material described below fully confirms to IS: 2830: 2012 Chemical Composition tested in accordance with scheme of testing and inspection contained in BIS certification marks License.",
                    font: "Calibri",
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "License Serial No. - ES/BIS/2830/7600107011.",
                    font: "Calibri",
                    bold: true,
                    size: 22,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Inv. No. ${tc.billNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Date - ${tc.tcDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. No. ${tc.poNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. Date - ${tc.poDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Total Qty. - ${tc.totalQtyMts}0`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Vehicle No. - ${tc.vehicleNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        columnSpan: 2,
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                text: `Total Pcs. - ${tc.totalPcs}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
                width: {
                  size: 5000,
                  type: docx.WidthType.DXA,
                },
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `BUYER :- ${tc.buyerName}`,
                    font: "Calibri",
                    bold: true,
                    size: 24,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `CHEMICAL COMPOSITION`,
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Size MM",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Heat No.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Qty. Pcs.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "C %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Mn %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "P %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "S %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Si %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Grade",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Color Code",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    height: { value: 50 },
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    text: "PHYSICAL PARAMETERS",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "All Physical Properties are as per & within IS 2830:2012 Requirements & Tolerances.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "The material is free from Radioactive Contamination, Lead, Mercury and Surface Defects.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 1,
                  }),
                  new docx.TextRun({
                    text: "Process Route - EIF-CCM",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "                                                                                                                                                      Q.C.",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.RIGHT,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "ELAF STEEL",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
            ],
          },
        ],
      });
      docx.Packer.toBuffer(doc5).then((buffer) => {
        fs.writeFileSync(
          `${tc.buyerName} ${tc.billNo} ${heatData[0].sectionSize}MM.docx`,
          buffer
        );
      });
    }

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
// router.get(
//   "/:id/genratePdf",
//   catchAsync(async (req, res) => {
//     console.log("hi");
//     try {
//       console.log("hi");
//       const browser = await puppeteer.launch({ headless: true });
//       const page = await browser.newPage();
//       const url =
//         `${req.protocol}://${req.get("host")}` +
//         "/billets/" +
//         `${req.params.id}` +
//         "/tcPreview";

//       console.log(url);
//       await page.goto(
//         `${req.protocol}://${req.get("host")}` +
//           "/billets/" +
//           `${req.params.id}` +
//           "/tcPreview",
//         {
//           waitUntil: "networkidle2",
//         }
//       );

//       await page.setViewport({ width: 1600, height: 1050 });
//       await page.evaluate(() => {
//         (document.querySelectorAll("a") || []).forEach((el) => el.remove());
//       });

//       // let div_selector_to_remove = "removeBtnPuppteer";
//       // await page.evaluate((sel) => {
//       //   var elements = document.querySelectorAll(sel);
//       //   for (var i = 0; i < elements.length; i++) {
//       //     elements[i].parentNode.removeChild(elements[i]);
//       //   }
//       // }, div_selector_to_remove);

//       const todayDate = new Date();
//       const pdfn = await page.pdf({
//         path: `${path.join(
//           __dirname,
//           "../files",
//           todayDate.getTime() + ".pdf"
//         )}`,
//         format: "a4",
//       });
//       await browser.close();
//       const pdfURL = path.join(
//         __dirname,
//         "../files",
//         todayDate.getTime() + ".pdf"
//       );
//       res.set({
//         "Content-Type": "application/pdf",
//         "Content-Length": pdfn.length,
//       });
//       res.sendFile(pdfURL);
//     } catch (error) {
//       console.log(error.message);
//     }
//   })
// );
// router.get(
//   "/:id/genratePdfCc",
//   catchAsync(async (req, res) => {
//     console.log("hi");
//     try {
//       console.log("hi");
//       const browser = await puppeteer.launch({ headless: true });
//       const page = await browser.newPage();
//       const url =
//         `${req.protocol}://${req.get("host")}` +
//         "/billets/" +
//         `${req.params.id}` +
//         "/closeCastTcPdf";
//       await page.goto(
//         `${req.protocol}://${req.get("host")}` +
//           "/billets/" +
//           `${req.params.id}` +
//           "/closeCastTcPdf",
//         {
//           waitUntil: "networkidle2",
//         }
//       );

//       await page.setViewport({ width: 1800, height: 1050 });
//       const todayDate = new Date();
//       const pdfn = await page.pdf({
//         path: `${path.join(
//           __dirname,
//           "../files",
//           todayDate.getTime() + ".pdf"
//         )}`,
//         format: "a3",
//         margin: { right: "0px", left: "0px" },
//       });
//       await browser.close();
//       const pdfURL = path.join(
//         __dirname,
//         "../files",
//         todayDate.getTime() + ".pdf"
//       );
//       res.set({
//         "Content-Type": "application/pdf",
//         "Content-Length": pdfn.length,
//       });
//       res.sendFile(pdfURL);
//     } catch (error) {
//       console.log(error.message);
//     }
//   })
// );
// GENERATE Word File (.docx) OPEN CASTING
router.get(
  "/:id/generateDocx",
  catchAsync(async (req, res, next) => {
    console.log("hi");
    const tc = await Tc.findById(req.params.id).populate("heatNo");
    console.log(tc._id);
    let tcHeats = tc.heatNo;
    let heats = tcHeats.map(({ _id, heatNo }) => ({
      _id,
      heatNo,
    }));
    // let heats = tcHeats.map();
    const tcList = await Tc.find({});
    const billet = await Billets.find({});
    // console.log(tc);
    let heatId = heats.map((heats) => heats._id);

    let heatData = await Billets.find({ _id: { $in: heatId } });
    console.log(heatData[1].heatNo);

    if (heatData.length == 1) {
      const doc1 = new docx.Document({
        sections: [
          {
            children: [
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "TEST CERTIFICATE FOR CARBON/ALLOY STEEL BILLETS FOR RE-ROLLING INTO GENERAL/STRUCTURAL PURPOSES.",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "It is certified that the material described below fully confirms to IS: 2830: 2012 Chemical Composition tested in accordance with scheme of testing and inspection contained in BIS certification marks License.",
                    font: "Calibri",
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "License Serial No. - ES/BIS/2830/7600107011.",
                    font: "Calibri",
                    bold: true,
                    size: 22,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Inv. No. ${tc.billNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Date - ${tc.tcDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. No. ${tc.poNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. Date - ${tc.poDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Total Qty. - ${tc.totalQtyMts}0`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Vehicle No. - ${tc.vehicleNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        columnSpan: 2,
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                text: `Total Pcs. - ${tc.totalPcs}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
                width: {
                  size: 5000,
                  type: docx.WidthType.DXA,
                },
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `BUYER :- ${tc.buyerName}`,
                    font: "Calibri",
                    bold: true,
                    size: 24,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `CHEMICAL COMPOSITION`,
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Size MM",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Heat No.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Qty. Pcs.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "C %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Mn %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "P %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "S %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Si %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Grade",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Color Code",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    height: { value: 50 },
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    text: "PHYSICAL PARAMETERS",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "All Physical Properties are as per & within IS 2830:2012 Requirements & Tolerances.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "The material is free from Radioactive Contamination, Lead, Mercury and Surface Defects.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 1,
                  }),
                  new docx.TextRun({
                    text: "Process Route - EIF-CCM",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "                                                                                                                                                      Q.C.",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.RIGHT,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "ELAF STEEL",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
            ],
          },
        ],
      });
      docx.Packer.toBuffer(doc1).then((buffer) => {
        fs.writeFileSync(
          `${tc.buyerName} ${tc.billNo} ${heatData[0].sectionSize}MM.docx`,
          buffer
        );
      });
    } else if (heatData.length == 2) {
      const doc2 = new docx.Document({
        sections: [
          {
            children: [
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "TEST CERTIFICATE FOR CARBON/ALLOY STEEL BILLETS FOR RE-ROLLING INTO GENERAL/STRUCTURAL PURPOSES.",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "It is certified that the material described below fully confirms to IS: 2830: 2012 Chemical Composition tested in accordance with scheme of testing and inspection contained in BIS certification marks License.",
                    font: "Calibri",
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "License Serial No. - ES/BIS/2830/7600107011.",
                    font: "Calibri",
                    bold: true,
                    size: 22,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Inv. No. ${tc.billNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Date - ${tc.tcDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. No. ${tc.poNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. Date - ${tc.poDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Total Qty. - ${tc.totalQtyMts}0`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Vehicle No. - ${tc.vehicleNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        columnSpan: 2,
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                text: `Total Pcs. - ${tc.totalPcs}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
                width: {
                  size: 5000,
                  type: docx.WidthType.DXA,
                },
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `BUYER :- ${tc.buyerName}`,
                    font: "Calibri",
                    bold: true,
                    size: 24,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `CHEMICAL COMPOSITION`,
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Size MM",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Heat No.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Qty. Pcs.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "C %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Mn %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "P %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "S %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Si %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Grade",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Color Code",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    text: "PHYSICAL PARAMETERS",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "All Physical Properties are as per & within IS 2830:2012 Requirements & Tolerances.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "The material is free from Radioactive Contamination, Lead, Mercury and Surface Defects.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 1,
                  }),
                  new docx.TextRun({
                    text: "Process Route - EIF-CCM",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "                                                                                                                                                      Q.C.",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.RIGHT,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "ELAF STEEL",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
            ],
          },
        ],
      });
      docx.Packer.toBuffer(doc2).then((buffer) => {
        fs.writeFileSync(
          `${tc.buyerName} ${tc.billNo} ${heatData[0].sectionSize}MM.docx`,
          buffer
        );
      });
    } else if (heatData.length == 3) {
      const doc3 = new docx.Document({
        sections: [
          {
            children: [
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "TEST CERTIFICATE FOR CARBON/ALLOY STEEL BILLETS FOR RE-ROLLING INTO GENERAL/STRUCTURAL PURPOSES.",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "It is certified that the material described below fully confirms to IS: 2830: 2012 Chemical Composition tested in accordance with scheme of testing and inspection contained in BIS certification marks License.",
                    font: "Calibri",
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "License Serial No. - ES/BIS/2830/7600107011.",
                    font: "Calibri",
                    bold: true,
                    size: 22,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Inv. No. ${tc.billNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Date - ${tc.tcDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. No. ${tc.poNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. Date - ${tc.poDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Total Qty. - ${tc.totalQtyMts}0`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Vehicle No. - ${tc.vehicleNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        columnSpan: 2,
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                text: `Total Pcs. - ${tc.totalPcs}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
                width: {
                  size: 5000,
                  type: docx.WidthType.DXA,
                },
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `BUYER :- ${tc.buyerName}`,
                    font: "Calibri",
                    bold: true,
                    size: 24,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `CHEMICAL COMPOSITION`,
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Size MM",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Heat No.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Qty. Pcs.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "C %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Mn %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "P %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "S %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Si %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Grade",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Color Code",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    text: "PHYSICAL PARAMETERS",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "All Physical Properties are as per & within IS 2830:2012 Requirements & Tolerances.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "The material is free from Radioactive Contamination, Lead, Mercury and Surface Defects.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 1,
                  }),
                  new docx.TextRun({
                    text: "Process Route - EIF-CCM",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "                                                                                                                                                      Q.C.",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.RIGHT,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "ELAF STEEL",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
            ],
          },
        ],
      });
      docx.Packer.toBuffer(doc3).then((buffer) => {
        fs.writeFileSync(
          `${tc.buyerName} ${tc.billNo} ${heatData[0].sectionSize}MM.docx`,
          buffer
        );
      });
    } else if (heatData.length == 4) {
      const doc4 = new docx.Document({
        sections: [
          {
            children: [
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "TEST CERTIFICATE FOR CARBON/ALLOY STEEL BILLETS FOR RE-ROLLING INTO GENERAL/STRUCTURAL PURPOSES.",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "It is certified that the material described below fully confirms to IS: 2830: 2012 Chemical Composition tested in accordance with scheme of testing and inspection contained in BIS certification marks License.",
                    font: "Calibri",
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "License Serial No. - ES/BIS/2830/7600107011.",
                    font: "Calibri",
                    bold: true,
                    size: 22,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Inv. No. ${tc.billNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Date - ${tc.tcDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. No. ${tc.poNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. Date - ${tc.poDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Total Qty. - ${tc.totalQtyMts}0`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Vehicle No. - ${tc.vehicleNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        columnSpan: 2,
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                text: `Total Pcs. - ${tc.totalPcs}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
                width: {
                  size: 5000,
                  type: docx.WidthType.DXA,
                },
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `BUYER :- ${tc.buyerName}`,
                    font: "Calibri",
                    bold: true,
                    size: 24,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `CHEMICAL COMPOSITION`,
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Size MM",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Heat No.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Qty. Pcs.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "C %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Mn %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "P %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "S %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Si %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Grade",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Color Code",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    height: { value: 50 },
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    text: "PHYSICAL PARAMETERS",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "All Physical Properties are as per & within IS 2830:2012 Requirements & Tolerances.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "The material is free from Radioactive Contamination, Lead, Mercury and Surface Defects.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 1,
                  }),
                  new docx.TextRun({
                    text: "Process Route - EIF-CCM",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "                                                                                                                                                      Q.C.",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.RIGHT,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "ELAF STEEL",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
            ],
          },
        ],
      });
      docx.Packer.toBuffer(doc4).then((buffer) => {
        fs.writeFileSync(
          `${tc.buyerName} ${tc.billNo} ${heatData[0].sectionSize}MM.docx`,
          buffer
        );
      });
    } else if (heatData.length == 5) {
      const doc5 = new docx.Document({
        sections: [
          {
            children: [
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "TEST CERTIFICATE FOR CARBON/ALLOY STEEL BILLETS FOR RE-ROLLING INTO GENERAL/STRUCTURAL PURPOSES.",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "It is certified that the material described below fully confirms to IS: 2830: 2012 Chemical Composition tested in accordance with scheme of testing and inspection contained in BIS certification marks License.",
                    font: "Calibri",
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "License Serial No. - ES/BIS/2830/7600107011.",
                    font: "Calibri",
                    bold: true,
                    size: 22,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Inv. No. ${tc.billNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Date - ${tc.tcDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. No. ${tc.poNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. Date - ${tc.poDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Total Qty. - ${tc.totalQtyMts}0`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Vehicle No. - ${tc.vehicleNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        columnSpan: 2,
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                text: `Total Pcs. - ${tc.totalPcs}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
                width: {
                  size: 5000,
                  type: docx.WidthType.DXA,
                },
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `BUYER :- ${tc.buyerName}`,
                    font: "Calibri",
                    bold: true,
                    size: 24,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `CHEMICAL COMPOSITION`,
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Size MM",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Heat No.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Qty. Pcs.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "C %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Mn %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "P %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "S %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Si %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Grade",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Color Code",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    height: { value: 50 },
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    text: "PHYSICAL PARAMETERS",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "All Physical Properties are as per & within IS 2830:2012 Requirements & Tolerances.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "The material is free from Radioactive Contamination, Lead, Mercury and Surface Defects.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 1,
                  }),
                  new docx.TextRun({
                    text: "Process Route - EIF-CCM",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "                                                                                                                                                      Q.C.",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.RIGHT,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "ELAF STEEL",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
            ],
          },
        ],
      });
      docx.Packer.toBuffer(doc5).then((buffer) => {
        fs.writeFileSync(
          `${tc.buyerName} ${tc.billNo} ${heatData[0].sectionSize}MM.docx`,
          buffer
        );
      });
    }

    res.render("billets/tcPreview", { tc, tcList });
  })
);
router.get(
  "/:id/tcPreview/closeCast",
  catchAsync(async (req, res, next) => {
    const tc = await Tc.findById(req.params.id).populate("heatNo");
    const tcList = await Tc.find({});
    const billet = await Billets.find({});
    console.log(tc);
    res.render("billets/closeCastTcPreview", { tc, tcList });
  })
);
router.get(
  "/:id/generateDocxCc",
  catchAsync(async (req, res, next) => {
    console.log("hi");
    const tc = await Tc.findById(req.params.id).populate("heatNo");
    console.log(tc._id);
    let tcHeats = tc.heatNo;
    let heats = tcHeats.map(({ _id, heatNo }) => ({
      _id,
      heatNo,
    }));
    // let heats = tcHeats.map();
    const tcList = await Tc.find({});
    const billet = await Billets.find({});
    // console.log(tc);
    let heatId = heats.map((heats) => heats._id);

    let heatData = await Billets.find({ _id: { $in: heatId } });
    console.log(heatData[1].heatNo);

    if (heatData.length == 1) {
      const doc1 = new docx.Document({
        sections: [
          {
            children: [
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "TEST CERTIFICATE FOR CARBON/ALLOY STEEL BILLETS FOR RE-ROLLING INTO GENERAL/STRUCTURAL PURPOSES.",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "It is certified that the material described below fully confirms to IS: 2830: 2012 Chemical Composition tested in accordance with scheme of testing and inspection contained in BIS certification marks License.",
                    font: "Calibri",
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "License Serial No. - ES/BIS/2830/7600107011.",
                    font: "Calibri",
                    bold: true,
                    size: 22,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Inv. No. ${tc.billNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Date - ${tc.tcDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. No. ${tc.poNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. Date - ${tc.poDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Total Qty. - ${tc.totalQtyMts}0`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Vehicle No. - ${tc.vehicleNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        columnSpan: 2,
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                text: `Total Pcs. - ${tc.totalPcs}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
                width: {
                  size: 5000,
                  type: docx.WidthType.DXA,
                },
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `BUYER :- ${tc.buyerName}`,
                    font: "Calibri",
                    bold: true,
                    size: 24,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `CHEMICAL COMPOSITION`,
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Size MM",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Heat No.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Qty. Pcs.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "C %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Mn %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "P %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "S %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Si %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Grade",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Color Code",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    height: { value: 50 },
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    text: "PHYSICAL PARAMETERS",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "All Physical Properties are as per & within IS 2830:2012 Requirements & Tolerances.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "The material is free from Radioactive Contamination, Lead, Mercury and Surface Defects.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 1,
                  }),
                  new docx.TextRun({
                    text: "Process Route - EIF-CCM",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "                                                                                                                                                      Q.C.",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.RIGHT,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "ELAF STEEL",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
            ],
          },
        ],
      });
      docx.Packer.toBuffer(doc1).then((buffer) => {
        fs.writeFileSync(
          `${tc.buyerName} ${tc.billNo} ${heatData[0].sectionSize}MM.docx`,
          buffer
        );
      });
    } else if (heatData.length == 2) {
      const doc2 = new docx.Document({
        sections: [
          {
            children: [
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "TEST CERTIFICATE FOR CARBON/ALLOY STEEL BILLETS FOR RE-ROLLING INTO GENERAL/STRUCTURAL PURPOSES.",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "It is certified that the material described below fully confirms to IS: 2830: 2012 Chemical Composition tested in accordance with scheme of testing and inspection contained in BIS certification marks License.",
                    font: "Calibri",
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "License Serial No. - ES/BIS/2830/7600107011.",
                    font: "Calibri",
                    bold: true,
                    size: 22,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Inv. No. ${tc.billNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Date - ${tc.tcDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. No. ${tc.poNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. Date - ${tc.poDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Total Qty. - ${tc.totalQtyMts}0`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Vehicle No. - ${tc.vehicleNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        columnSpan: 2,
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                text: `Total Pcs. - ${tc.totalPcs}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
                width: {
                  size: 5000,
                  type: docx.WidthType.DXA,
                },
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `BUYER :- ${tc.buyerName}`,
                    font: "Calibri",
                    bold: true,
                    size: 24,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `CHEMICAL COMPOSITION`,
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Size MM",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Heat No.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Qty. Pcs.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "C %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Mn %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "P %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "S %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Si %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Grade",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Color Code",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    text: "PHYSICAL PARAMETERS",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "All Physical Properties are as per & within IS 2830:2012 Requirements & Tolerances.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "The material is free from Radioactive Contamination, Lead, Mercury and Surface Defects.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 1,
                  }),
                  new docx.TextRun({
                    text: "Process Route - EIF-CCM",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "                                                                                                                                                      Q.C.",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.RIGHT,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "ELAF STEEL",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
            ],
          },
        ],
      });
      docx.Packer.toBuffer(doc2).then((buffer) => {
        fs.writeFileSync(
          `${tc.buyerName} ${tc.billNo} ${heatData[0].sectionSize}MM.docx`,
          buffer
        );
      });
    } else if (heatData.length == 3) {
      const doc3 = new docx.Document({
        sections: [
          {
            children: [
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "TEST CERTIFICATE FOR CARBON/ALLOY STEEL BILLETS FOR RE-ROLLING INTO GENERAL/STRUCTURAL PURPOSES.",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "It is certified that the material described below fully confirms to IS: 2830: 2012 Chemical Composition tested in accordance with scheme of testing and inspection contained in BIS certification marks License.",
                    font: "Calibri",
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "License Serial No. - ES/BIS/2830/7600107011.",
                    font: "Calibri",
                    bold: true,
                    size: 22,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Inv. No. ${tc.billNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Date - ${tc.tcDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. No. ${tc.poNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. Date - ${tc.poDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Total Qty. - ${tc.totalQtyMts}0`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Vehicle No. - ${tc.vehicleNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        columnSpan: 2,
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                text: `Total Pcs. - ${tc.totalPcs}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
                width: {
                  size: 5000,
                  type: docx.WidthType.DXA,
                },
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `BUYER :- ${tc.buyerName}`,
                    font: "Calibri",
                    bold: true,
                    size: 24,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `CHEMICAL COMPOSITION`,
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Size MM",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Heat No.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Qty. Pcs.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "C %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Mn %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "P %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "S %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Si %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Grade",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Color Code",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    text: "PHYSICAL PARAMETERS",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "All Physical Properties are as per & within IS 2830:2012 Requirements & Tolerances.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "The material is free from Radioactive Contamination, Lead, Mercury and Surface Defects.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 1,
                  }),
                  new docx.TextRun({
                    text: "Process Route - EIF-CCM",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "                                                                                                                                                      Q.C.",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.RIGHT,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "ELAF STEEL",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
            ],
          },
        ],
      });
      docx.Packer.toBuffer(doc3).then((buffer) => {
        fs.writeFileSync(
          `${tc.buyerName} ${tc.billNo} ${heatData[0].sectionSize}MM.docx`,
          buffer
        );
      });
    } else if (heatData.length == 4) {
      const doc4 = new docx.Document({
        sections: [
          {
            children: [
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "TEST CERTIFICATE FOR CARBON/ALLOY STEEL BILLETS FOR RE-ROLLING INTO GENERAL/STRUCTURAL PURPOSES.",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "It is certified that the material described below fully confirms to IS: 2830: 2012 Chemical Composition tested in accordance with scheme of testing and inspection contained in BIS certification marks License.",
                    font: "Calibri",
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "License Serial No. - ES/BIS/2830/7600107011.",
                    font: "Calibri",
                    bold: true,
                    size: 22,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Inv. No. ${tc.billNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Date - ${tc.tcDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. No. ${tc.poNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. Date - ${tc.poDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Total Qty. - ${tc.totalQtyMts}0`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Vehicle No. - ${tc.vehicleNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        columnSpan: 2,
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                text: `Total Pcs. - ${tc.totalPcs}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
                width: {
                  size: 5000,
                  type: docx.WidthType.DXA,
                },
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `BUYER :- ${tc.buyerName}`,
                    font: "Calibri",
                    bold: true,
                    size: 24,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `CHEMICAL COMPOSITION`,
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Size MM",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Heat No.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Qty. Pcs.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "C %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Mn %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "P %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "S %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Si %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Grade",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Color Code",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    height: { value: 50 },
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    text: "PHYSICAL PARAMETERS",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "All Physical Properties are as per & within IS 2830:2012 Requirements & Tolerances.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "The material is free from Radioactive Contamination, Lead, Mercury and Surface Defects.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 1,
                  }),
                  new docx.TextRun({
                    text: "Process Route - EIF-CCM",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "                                                                                                                                                      Q.C.",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.RIGHT,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "ELAF STEEL",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
            ],
          },
        ],
      });
      docx.Packer.toBuffer(doc4).then((buffer) => {
        fs.writeFileSync(
          `${tc.buyerName} ${tc.billNo} ${heatData[0].sectionSize}MM.docx`,
          buffer
        );
      });
    } else if (heatData.length == 5) {
      const doc5 = new docx.Document({
        sections: [
          {
            children: [
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "TEST CERTIFICATE FOR CARBON/ALLOY STEEL BILLETS FOR RE-ROLLING INTO GENERAL/STRUCTURAL PURPOSES.",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 2,
                    text: "It is certified that the material described below fully confirms to IS: 2830: 2012 Chemical Composition tested in accordance with scheme of testing and inspection contained in BIS certification marks License.",
                    font: "Calibri",
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "License Serial No. - ES/BIS/2830/7600107011.",
                    font: "Calibri",
                    bold: true,
                    size: 22,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Inv. No. ${tc.billNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Date - ${tc.tcDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. No. ${tc.poNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Po. Date - ${tc.poDate}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Total Qty. - ${tc.totalQtyMts}0`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.TextRun({
                                text: `Vehicle No. - ${tc.vehicleNo}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        columnSpan: 2,
                        width: {
                          size: 2500,
                          type: docx.WidthType.DXA,
                        },
                        borders: {
                          top: {
                            style: docx.BorderStyle.SINGLE,
                            size: 1,
                          },
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                text: `Total Pcs. - ${tc.totalPcs}`,
                                font: "Calibri",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
                width: {
                  size: 5000,
                  type: docx.WidthType.DXA,
                },
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `BUYER :- ${tc.buyerName}`,
                    font: "Calibri",
                    bold: true,
                    size: 24,

                    underline: {
                      type: docx.UnderlineType.SINGLE,
                    },
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: `CHEMICAL COMPOSITION`,
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
              new docx.Table({
                alignment: docx.AlignmentType.CENTER,
                rows: [
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Size MM",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Heat No.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Qty. Pcs.",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "C %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Mn %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "P %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "S %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Si %",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Grade",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,

                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: "Color Code",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    height: { value: 50 },
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[0].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[1].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[2].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[3].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new docx.TableRow({
                    children: [
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].sectionSize,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].heatNo,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].totalPcs,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].c,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].mn,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].p,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].s,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].si,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].gradeName,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new docx.TableCell({
                        verticalAlign: docx.VerticalAlign.CENTER,

                        width: {
                          size: 1900,
                          type: docx.WidthType.DXA,
                        },
                        children: [
                          new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                              new docx.TextRun({
                                font: "Calibri",
                                text: heatData[4].colorCode,
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 3,
                  }),
                  new docx.TextRun({
                    text: "PHYSICAL PARAMETERS",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "All Physical Properties are as per & within IS 2830:2012 Requirements & Tolerances.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                bullet: { level: 0 },
                children: [
                  new docx.TextRun({
                    text: "The material is free from Radioactive Contamination, Lead, Mercury and Surface Defects.",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    break: 1,
                  }),
                  new docx.TextRun({
                    text: "Process Route - EIF-CCM",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "                                                                                                                                                     Q.C.",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                  new docx.TextRun({
                    break: 1,
                  }),
                ],
              }),
              new docx.Paragraph({
                alignment: docx.AlignmentType.RIGHT,
                children: [
                  new docx.TextRun({
                    break: 1,
                    text: "ELAF STEEL",
                    font: "Calibri",
                    bold: true,
                    size: 22,
                  }),
                ],
              }),
            ],
          },
        ],
      });
      docx.Packer.toBuffer(doc5).then((buffer) => {
        fs.writeFileSync(
          `${tc.buyerName} ${tc.billNo} ${heatData[0].sectionSize}MM.docx`,
          buffer
        );
      });
    }

    res.render("billets/closeCastTcPreview", { tc, tcList });
  })
);
// // Function for generating pdf
// const g = async (req, res) => {
//   try {
//     const browser = await puppeteer.launch();
//     const page = await browser.newPage();
//     await page.goto(
//       `${req.protocol}://${req.get("host")}` + "/views/billets/tcPdf.ejs",
//       {
//         waitUntil: "networkidle2",
//       }
//     );
//     await page.setViewport({ width: 1680, height: 1050 });
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
