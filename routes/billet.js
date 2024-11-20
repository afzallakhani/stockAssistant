const express = require("express");
const router = express.Router();
const path = require("path");
const officegen = require("officegen");
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
const sanitize = require("sanitize-filename");
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
const { Console } = require("console");
const eventEmitter = new events.EventEmitter();
let upload = multer({ storage: multerStorage });

const outputDirectory = path.join(__dirname, "..", "output-files docx");

if (!fs.existsSync(outputDirectory)) {
    fs.mkdirSync(outputDirectory, { recursive: true });
}

router.get(
    "/list",
    catchAsync(async(req, res) => {
        const list = await Billets.find({}).sort({ createdAt: -1 }); // Newest to oldest
        res.render("billets/list", { list });
    })
);

router.get("/new", (req, res) => {
    res.render("billets/new");
});

// searched Heat display
router.get("/search", async(req, res) => {
    const queryString = req.query.heat;
    let query = queryString.heatNo.toString();
    let queryList = query.split(" ");

    const heats = await Billets.find({ heatNo: { $in: queryList } });

    res.render("billets/search", { heats });
});

router.get(
    "/tcList",
    catchAsync(async(req, res, next) => {
        const tcList = await Tc.find({}).sort({ createdAt: -1 }).populate("heatNo");
        const billet = await Billets.find({});
        // console.log(tcList);
        // billetTc.updateMany(
        //   {},
        //   { $set: { createdAt: new Date() } },
        //   (err, result) => {
        //     if (err) {
        //       console.error(err);
        //     } else {
        //       console.log(`${result.nModified} documents updated`);
        //     }
        //   }
        // );
        res.render("billets/tcList", { tcList });
    })
);
router.get(
    "/newTc",
    catchAsync(async(req, res, next) => {
        const party = await Party.find();
        const list = await Billets.find({});
        const tc = await Tc.find({});
        // console.log(tc);
        res.render("billets/newTc", { list, party });
    })
);
router.post(
    "/newTc",
    catchAsync(async(req, res, next) => {
        console.log(req.body.billet);
        let tc = new Tc(req.body.tc);

        let queryList = req.body.billet;
        let heatArray = Object.values(queryList);
        let heatQuery = heatArray.toString().split(" ");
        // console.log(heatQuery);
        let heats = await Billets.find({ heatNo: { $in: heatQuery } });
        if (heats.length === 0) {
            res.status(404).json({ message: "Please Enter Correct Heat No." });
        } else {
            console.log(heats);

            const heatId = heats.map((item) => item._id);
            tc.heatNo = heatId;
            // console.log(tc);
            await tc.save();
            let = currentTc = await Tc.findById(tc._id);
            // console.log(currentTc);
            res.redirect(`/billets/${tc._id}/tcPreview`);
        }
    })
);

// EDIT TC VIEW ROUTE
router.get(
    "/:id/editTc",
    catchAsync(async(req, res, next) => {
        const tc = await Tc.findById(req.params.id).populate("heatNo");
        const party = await Party.find();
        // console.log(party);
        res.render("billets/editTc", { tc, party });
    })
);
// EDIT TC in data base
router.put(
    "/:id/editTc",
    catchAsync(async(req, res, next) => {
        const { id } = req.params;
        const tc = await Tc.findById(id).populate("heatNo");

        let updatedTc = req.body.tc;
        let updatedHeatNoList = req.body.billet;
        let updatedHeatArray = Object.values(updatedHeatNoList);
        let heatDataQuery = updatedHeatArray.toString().split(" ");
        let newHeatsData = await Billets.find({ heatNo: { $in: heatDataQuery } });
        let newHeatsIds = newHeatsData.map((heat) => heat._id);
        updatedTc.heatNo = newHeatsIds;
        let editedTc = await Tc.findByIdAndUpdate(id, updatedTc);
        res.redirect(`/billets/${id}/tcPreview`);
    })
);

// DELETE TC
router.delete(
    "/:id/tc",
    catchAsync(async(req, res) => {
        const { id } = req.params;
        console.log("deleted Tc");
        await Tc.findByIdAndDelete(id);
        res.redirect("/billets/tcList");
    })
);

// TC PREVIEW GET
router.get(
    "/:id/tcPreview",
    catchAsync(async(req, res, next) => {
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
        // console.log(heatData[0].heatNo);

        res.render("billets/tcPreview", { tc, tcList });
    })
);

// TC READY FOR MAKING PDF OPEN CASTING
router.get(
    "/:id/tcPdf",
    catchAsync(async(req, res, next) => {
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
    catchAsync(async(req, res, next) => {
        const tcId = req.params.id;

        const tc = await Tc.findById(req.params.id).populate("heatNo");
        const tcList = await Tc.find({});
        const billet = await Billets.find({});
        // console.log(tc);

        res.render("billets/closeCastTcPdf", { tc, tcList });
    })
);
// GENERATE Word File (.docx) OPEN CASTING
router.get(
    "/:id/generateDocx",
    catchAsync(async(req, res, next) => {
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
        // console.log(heatData[1].heatNo);

        if (heatData.length == 1) {
            const doc1 = new docx.Document({
                compatibility: {
                    version: 16,
                },
                styles: {
                    default: {
                        document: {
                            run: {
                                size: "10pt",
                                font: "Calibri",
                            },
                        },
                    },
                },
                sections: [{
                    headers: {
                        default: new docx.Header({
                            children: [
                                new docx.Paragraph({
                                    alignment: docx.AlignmentType.RIGHT,
children:[
    new docx.TextRun({
        text: "IS2830:2012",
        font: "Calibri",
        bold: true,
        size:18,
    }),
    new docx.TextRun({
        text: "CM/L-2809767",
        font: "Calibri",
        bold: true,
        size:18,
    }),
],
                                }),
                                new docx.Paragraph({
                                alignment: docx.AlignmentType.RIGHT,
                                children: [
                                    new docx.TextRun({
                                        text: "Is2830:2012",
                                        font: "Calibri",
                                    }),
                                    new docx.ImageRun({
                                        data: fs.readFileSync('D:\\stockAssistant\\ISI_Mark_Logo_Cdr-Vector-file-.jpg'),
                                        transformation: {
                                            width: 50,
                                            height: 40,
                                        }
                                    }),
                                ],
                            }),],
                        }),
                    },
                    children: [
                        new docx.Paragraph({
                            children: [
                                new docx.ImageRun({
                                    data: fs.readFileSync('D:\\stockAssistant\\ISI_Mark_Logo_Cdr-Vector-file-.jpg'),
                                    transformation: {
                                        width: 50,
                                        height: 50,
                                    }
                                }),
                            ],
                        }),
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
                                    text: "License Serial No. - SA/BIS/2830/2809767.",
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                            text: `Date - ${tc.formattedTcDate}`,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                            text: `Po. Date - ${tc.formattedPoDate}`,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                            width: {
                                size: 10000,
                                type: docx.WidthType.DXA,
                            },
                            rows: [
                                new docx.TableRow({
                                    children: [
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "Pcs",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                }, ],
            });
            // docx.Packer.toBuffer(doc1).then((buffer) => {
            //   fs.writeFileSync(
            //     `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} ${heatData[0].sectionSize}MM.docx`,
            //     buffer
            //   );
            // });
            docx.Packer.toBuffer(doc1).then((buffer) => {
                const filename = `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} ${heatData[0].sectionSize}MM.docx`;
                const sanitizedFilename = sanitize(filename);

                // Combine directory and sanitized filename
                const filePath = path.join(outputDirectory, sanitizedFilename);

                fs.writeFileSync(filePath, buffer);
                console.log(`File saved as: ${filePath}`);
            });
        } else if (heatData.length == 2) {
            const doc2 = new docx.Document({
                compatibility: {
                    version: 16,
                },
                styles: {
                    default: {
                        document: {
                            run: {
                                size: "10pt",
                                font: "Calibri",
                            },
                        },
                    },
                },
                sections: [{
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                            text: `Date - ${tc.formattedTcDate}`,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                            text: `Po. Date - ${tc.formattedPoDate}`,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                            width: {
                                size: 10000,
                                type: docx.WidthType.DXA,
                            },
                            rows: [
                                new docx.TableRow({
                                    children: [
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "Pcs",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                }, ],
            });
            // docx.Packer.toBuffer(doc2).then((buffer) => {
            //   fs.writeFileSync(
            //     `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} ${heatData[0].sectionSize}MM.docx`,
            //     buffer
            //   );
            // });
            docx.Packer.toBuffer(doc2).then((buffer) => {
                const filename = `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} ${heatData[0].sectionSize}MM.docx`;
                const sanitizedFilename = sanitize(filename);

                // Combine directory and sanitized filename
                const filePath = path.join(outputDirectory, sanitizedFilename);

                fs.writeFileSync(filePath, buffer);
                console.log(`File saved as: ${filePath}`);
            });
        } else if (heatData.length == 3) {
            const doc3 = new docx.Document({
                compatibility: {
                    version: 16,
                },
                styles: {
                    default: {
                        document: {
                            run: {
                                size: "10pt",
                                font: "Calibri",
                            },
                        },
                    },
                },
                sections: [{
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
                                    text: "License Serial No. - SA/BIS/2830/2809767.",
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                            text: `Date - ${tc.formattedTcDate}`,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                            text: `Po. Date - ${tc.formattedPoDate}`,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                            width: {
                                size: 10000,
                                type: docx.WidthType.DXA,
                            },
                            alignment: docx.AlignmentType.CENTER,
                            rows: [
                                new docx.TableRow({
                                    children: [
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "Pcs",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                }, ],
            });
            // docx.Packer.toBuffer(doc3).then((buffer) => {
            //   fs.writeFileSync(
            //     `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} ${heatData[0].sectionSize}MM.docx`,
            //     buffer
            //   );
            // });
            docx.Packer.toBuffer(doc3).then((buffer) => {
                const filename = `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} ${heatData[0].sectionSize}MM.docx`;
                const sanitizedFilename = sanitize(filename);

                // Combine directory and sanitized filename
                const filePath = path.join(outputDirectory, sanitizedFilename);

                fs.writeFileSync(filePath, buffer);
                console.log(`File saved as: ${filePath}`);
            });
        } else if (heatData.length == 4) {
            const doc4 = new docx.Document({
                compatibility: {
                    version: 16,
                },
                styles: {
                    default: {
                        document: {
                            run: {
                                size: "10pt",
                                font: "Calibri",
                            },
                        },
                    },
                },
                sections: [{
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                            text: `Date - ${tc.formattedTcDate}`,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                            text: `Po. Date - ${tc.formattedPoDate}`,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                            width: {
                                size: 10000,
                                type: docx.WidthType.DXA,
                            },
                            alignment: docx.AlignmentType.CENTER,
                            rows: [
                                new docx.TableRow({
                                    children: [
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "Pcs",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                }, ],
            });
            // docx.Packer.toBuffer(doc4).then((buffer) => {
            //   fs.writeFileSync(
            //     `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} ${heatData[0].sectionSize}MM.docx`,
            //     buffer
            //   );
            // });
            docx.Packer.toBuffer(doc4).then((buffer) => {
                const filename = `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} ${heatData[0].sectionSize}MM.docx`;
                const sanitizedFilename = sanitize(filename);

                // Combine directory and sanitized filename
                const filePath = path.join(outputDirectory, sanitizedFilename);

                fs.writeFileSync(filePath, buffer);
                console.log(`File saved as: ${filePath}`);
            });
        } else if (heatData.length == 5) {
            const doc5 = new docx.Document({
                compatibility: {
                    version: 16,
                },
                styles: {
                    default: {
                        document: {
                            run: {
                                size: "10pt",
                                font: "Calibri",
                            },
                        },
                    },
                },
                sections: [{
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                            text: `Date - ${tc.formattedTcDate}`,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                            text: `Po. Date - ${tc.formattedPoDate}`,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                            width: {
                                size: 11700,
                                type: docx.WidthType.DXA,
                            },
                            alignment: docx.AlignmentType.CENTER,
                            rows: [
                                new docx.TableRow({
                                    children: [
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "Pcs",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                }, ],
            });
            // docx.Packer.toBuffer(doc5).then((buffer) => {
            //   fs.writeFileSync(
            //     `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} ${heatData[0].sectionSize}MM.docx`,
            //     buffer
            //   );
            // });
            docx.Packer.toBuffer(doc5).then((buffer) => {
                const filename = `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} ${heatData[0].sectionSize}MM.docx`;
                const sanitizedFilename = sanitize(filename);

                // Combine directory and sanitized filename
                const filePath = path.join(outputDirectory, sanitizedFilename);

                fs.writeFileSync(filePath, buffer);
                console.log(`File saved as: ${filePath}`);
            });
        }
        // let docxUrl = path.join(
        //     __dirname,
        //     "../",
        //     `${tc.buyerName} ${tc.billNo} ${heatData[0].sectionSize}MM.docx`
        // );
        // console.log(docxUrl);

        // res.render("billets/closeCastTcPreview", { tc, tcList });
        // res.download(docxUrl, function(e) {
        //     if (e) {
        //         console.log(e);
        //     } else {
        //         console.log("Download Complete!!!");
        //     }
        // });
        res.render("billets/tcPreview", { tc, tcList });
    })
);
router.get(
    "/:id/downloadDocx",
    catchAsync(async(req, res, next) => {
        console.log("Downloading....");
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
        // let docxUrl = path.join(
        //   __dirname,
        //   "../",
        //   `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} ${heatData[0].sectionSize}MM.docx`
        // );
        // console.log(docxUrl);
        // console.log("agh");
        // res.download(docxUrl, function (e) {
        //   if (e) {
        //     console.log(e);
        //   } else {
        //     console.log("Download Complete!!!");
        //   }
        // });
        // Define the directory where the files are saved
        const outputDirectory = path.join(__dirname, "..", "output-files docx");

        // Construct the filename and sanitize it
        const filename = `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} ${heatData[0].sectionSize}MM.docx`;
        const sanitizedFilename = sanitize(filename);

        // Combine directory and sanitized filename
        let docxUrl = path.join(outputDirectory, sanitizedFilename);

        // Download the file
        res.download(docxUrl, function(e) {
            if (e) {
                console.log(e);
            } else {
                console.log("Download Complete!!!");
            }
        });
    })
);
router.get(
    "/:id/tcPreview/closeCast",
    catchAsync(async(req, res, next) => {
        const tc = await Tc.findById(req.params.id).populate("heatNo");
        const tcList = await Tc.find({});
        const billet = await Billets.find({});
        console.log(tc);
        res.render("billets/closeCastTcPreview", { tc, tcList });
    })
);
router.get(
    "/:id/generateDocxCc",
    catchAsync(async(req, res, next) => {
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

        if (heatData.length == 1) {
            const doc1Cc = new docx.Document({
                compatibility: {
                    version: 16,
                },
                styles: {
                    default: {
                        document: {
                            run: {
                                size: "10pt",
                                font: "Calibri",
                            },
                        },
                    },
                },
                sections: [{
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                            text: `Date - ${tc.formattedTcDate}`,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                            text: `Po. Date - ${tc.formattedPoDate}`,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                            width: {
                                size: 10000,
                                type: docx.WidthType.DXA,
                            },
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
                            width: {
                                size: 11700,
                                type: docx.WidthType.DXA,
                            },
                            alignment: docx.AlignmentType.CENTER,
                            rows: [
                                new docx.TableRow({
                                    children: [
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "Pcs",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "AL %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "CR %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "NI %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "MO %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "CU %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "V %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "CE %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].al,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].cr,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].ni,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].mo,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].cu,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].v,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].ce,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                            bullet: { level: 0 },
                            children: [
                                new docx.TextRun({
                                    text: "The material is Argon Purged through the ladle.",
                                    font: "Calibri",
                                    size: 22,
                                }),
                            ],
                        }),
                        new docx.Paragraph({
                            bullet: { level: 0 },
                            children: [
                                new docx.TextRun({
                                    text: `ALL CLOSE CAST BILLETS OF ${heatData[0].gradeName} ARE COLORED ${tc.colorCode} AT ONE END.`,
                                    font: "Calibri",
                                    size: 22,
                                    bold: true,
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
                }, ],
            });
            // docx.Packer.toBuffer(doc1Cc).then((buffer) => {
            //   fs.writeFileSync(
            //     `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} CC ${heatData[0].sectionSize}MM.docx`,
            //     buffer
            //   );
            // });
            docx.Packer.toBuffer(doc1Cc).then((buffer) => {
                const filename = `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} CC ${heatData[0].sectionSize}MM.docx`;
                const sanitizedFilename = sanitize(filename);

                // Combine directory and sanitized filename
                const filePath = path.join(outputDirectory, sanitizedFilename);

                fs.writeFileSync(filePath, buffer);
                console.log(`File saved as: ${filePath}`);
            });
        } else if (heatData.length == 2) {
            const doc2Cc = new docx.Document({
                compatibility: {
                    version: 16,
                },
                styles: {
                    default: {
                        document: {
                            run: {
                                size: "10pt",
                                font: "Calibri",
                            },
                        },
                    },
                },
                sections: [{
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                            text: `Date - ${tc.formattedTcDate}`,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                            text: `Po. Date - ${tc.formattedPoDate}`,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                            width: {
                                size: 11700,
                                type: docx.WidthType.DXA,
                            },
                            alignment: docx.AlignmentType.CENTER,
                            rows: [
                                new docx.TableRow({
                                    children: [
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "Pcs",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "AL %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "CR %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "NI %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "MO %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "CU %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "V %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "CE %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].al,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].cr,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].ni,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].mo,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].cu,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].v,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].ce,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].al,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].cr,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].ni,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].mo,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].cu,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].v,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].ce,
                                                            // Doc5Cc update chemistry column completed 0 and 1
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                            bullet: { level: 0 },
                            children: [
                                new docx.TextRun({
                                    text: "The material is Argon Purged through the ladle.",
                                    font: "Calibri",
                                    size: 22,
                                }),
                            ],
                        }),
                        new docx.Paragraph({
                            bullet: { level: 0 },
                            children: [
                                new docx.TextRun({
                                    text: `ALL CLOSE CAST BILLETS OF ${heatData[1].gradeName} ARE COLORED ${tc.colorCode} AT ONE END.`,
                                    font: "Calibri",
                                    size: 22,
                                    bold: true,
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
                }, ],
            });
            // docx.Packer.toBuffer(doc2Cc).then((buffer) => {
            //   fs.writeFileSync(
            //     `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} CC ${heatData[0].sectionSize}MM.docx`,
            //     buffer
            //   );
            // });
            docx.Packer.toBuffer(doc2Cc).then((buffer) => {
                const filename = `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} CC ${heatData[0].sectionSize}MM.docx`;
                const sanitizedFilename = sanitize(filename);

                // Combine directory and sanitized filename
                const filePath = path.join(outputDirectory, sanitizedFilename);

                fs.writeFileSync(filePath, buffer);
                console.log(`File saved as: ${filePath}`);
            });
        } else if (heatData.length == 3) {
            const doc3Cc = new docx.Document({
                compatibility: {
                    version: 16,
                },
                styles: {
                    default: {
                        document: {
                            run: {
                                size: "10pt",
                                font: "Calibri",
                            },
                        },
                    },
                },
                sections: [{
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                            text: `Date - ${tc.formattedTcDate}`,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                            text: `Po. Date - ${tc.formattedPoDate}`,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                            width: {
                                size: 11700,
                                type: docx.WidthType.DXA,
                            },
                            alignment: docx.AlignmentType.CENTER,
                            rows: [
                                new docx.TableRow({
                                    children: [
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "Pcs",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "AL %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "CR %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "NI %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "MO %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "CU %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "V %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "CE %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].al,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].cr,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].ni,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].mo,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].cu,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].v,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].ce,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].al,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].cr,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].ni,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].mo,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].cu,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].v,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].ce,
                                                            // Doc5Cc update chemistry column completed 0 and 1
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[2].al,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[2].cr,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[2].ni,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[2].mo,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[2].cu,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[2].v,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[2].ce,
                                                            // Doc5Cc update chemistry column completed 0 and 1
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                            bullet: { level: 0 },
                            children: [
                                new docx.TextRun({
                                    text: "The material is Argon Purged through the ladle.",
                                    font: "Calibri",
                                    size: 22,
                                }),
                            ],
                        }),
                        new docx.Paragraph({
                            bullet: { level: 0 },
                            children: [
                                new docx.TextRun({
                                    text: `ALL CLOSE CAST BILLETS OF ${heatData[2].gradeName} ARE COLORED ${tc.colorCode} AT ONE END.`,
                                    font: "Calibri",
                                    size: 22,
                                    bold: true,
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
                }, ],
            });
            // docx.Packer.toBuffer(doc3Cc).then((buffer) => {
            //   fs.writeFileSync(
            //     `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} CC ${heatData[0].sectionSize}MM.docx`,
            //     buffer
            //   );
            // });
            docx.Packer.toBuffer(doc3Cc).then((buffer) => {
                const filename = `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} CC ${heatData[0].sectionSize}MM.docx`;
                const sanitizedFilename = sanitize(filename);

                // Combine directory and sanitized filename
                const filePath = path.join(outputDirectory, sanitizedFilename);

                fs.writeFileSync(filePath, buffer);
                console.log(`File saved as: ${filePath}`);
            });
        } else if (heatData.length == 4) {
            const doc4Cc = new docx.Document({
                compatibility: {
                    version: 16,
                },
                styles: {
                    default: {
                        document: {
                            run: {
                                size: "10pt",
                                font: "Calibri",
                            },
                        },
                    },
                },
                sections: [{
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                            text: `Date - ${tc.formattedTcDate}`,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                            text: `Po. Date - ${tc.formattedPoDate}`,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                            width: {
                                size: 11700,
                                type: docx.WidthType.DXA,
                            },
                            alignment: docx.AlignmentType.CENTER,
                            rows: [
                                new docx.TableRow({
                                    children: [
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "Pcs",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "AL %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "CR %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "NI %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "MO %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "CU %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "V %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "CE %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].al,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].cr,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].ni,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].mo,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].cu,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].v,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].ce,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].al,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].cr,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].ni,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].mo,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].cu,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].v,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].ce,
                                                            // Doc5Cc update chemistry column completed 0 and 1
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[2].al,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[2].cr,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[2].ni,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[2].mo,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[2].cu,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[2].v,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[2].ce,
                                                            // Doc5Cc update chemistry column completed 0 and 1
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[3].al,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[3].cr,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[3].ni,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[3].mo,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[3].cu,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[3].v,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[3].ce,
                                                            // Doc5Cc update chemistry column completed 0 and 1
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                            bullet: { level: 0 },
                            children: [
                                new docx.TextRun({
                                    text: "The material is Argon Purged through the ladle.",
                                    font: "Calibri",
                                    size: 22,
                                }),
                            ],
                        }),
                        new docx.Paragraph({
                            bullet: { level: 0 },
                            children: [
                                new docx.TextRun({
                                    text: `ALL CLOSE CAST BILLETS OF ${heatData[3].gradeName} ARE COLORED ${tc.colorCode} AT ONE END.`,
                                    font: "Calibri",
                                    size: 22,
                                    bold: true,
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
                }, ],
            });
            // docx.Packer.toBuffer(doc4Cc).then((buffer) => {
            //   fs.writeFileSync(
            //     `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} CC ${heatData[0].sectionSize}MM.docx`,
            //     buffer
            //   );
            // });
            docx.Packer.toBuffer(doc4Cc).then((buffer) => {
                const filename = `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} CC ${heatData[0].sectionSize}MM.docx`;
                const sanitizedFilename = sanitize(filename);

                // Combine directory and sanitized filename
                const filePath = path.join(outputDirectory, sanitizedFilename);

                fs.writeFileSync(filePath, buffer);
                console.log(`File saved as: ${filePath}`);
            });
        } else if (heatData.length == 5) {
            const doc5Cc = new docx.Document({
                compatibility: {
                    version: 16,
                },
                styles: {
                    default: {
                        document: {
                            run: {
                                size: "10pt",
                                font: "Calibri",
                            },
                        },
                    },
                },
                sections: [{
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
                            width: {
                                size: 10000,
                                type: docx.WidthType.DXA,
                            },
                            rows: [
                                new docx.TableRow({
                                    children: [
                                        new docx.TableCell({
                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                            text: `Date - ${tc.formattedTcDate}`,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                            text: `Po. Date - ${tc.formattedPoDate}`,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                            width: {
                                size: 11700,
                                type: docx.WidthType.DXA,
                                //   type:docx.WidthType.AUTO
                            },
                            alignment: docx.AlignmentType.CENTER,
                            rows: [
                                new docx.TableRow({
                                    children: [
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "Pcs",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "AL %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "CR %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "NI %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "MO %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "CU %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "V %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,

                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: "CE %",
                                                            bold: true,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].al,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].cr,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].ni,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].mo,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].cu,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].v,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[0].ce,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].al,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].cr,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].ni,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].mo,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].cu,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].v,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[1].ce,
                                                            // Doc5Cc update chemistry column completed 0 and 1
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[2].al,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[2].cr,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[2].ni,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[2].mo,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[2].cu,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[2].v,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[2].ce,
                                                            // Doc5Cc update chemistry column completed 0 and 1
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[3].al,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[3].cr,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[3].ni,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[3].mo,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[3].cu,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[3].v,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[3].ce,
                                                            // Doc5Cc update chemistry column completed 0 and 1
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[4].al,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[4].cr,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[4].ni,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[4].mo,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[4].cu,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[4].v,
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),

                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: heatData[4].ce,
                                                            // Doc5Cc update chemistry column completed 0 and 1
                                                        }),
                                                    ],
                                                }),
                                            ],
                                        }),
                                        new docx.TableCell({
                                            verticalAlign: docx.VerticalAlign.CENTER,

                                            width: {
                                                size: 10000,
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
                                                size: 10000,
                                                type: docx.WidthType.DXA,
                                            },
                                            children: [
                                                new docx.Paragraph({
                                                    alignment: docx.AlignmentType.CENTER,
                                                    children: [
                                                        new docx.TextRun({
                                                            font: "Calibri",
                                                            text: tc.colorCode,
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
                            bullet: { level: 0 },
                            children: [
                                new docx.TextRun({
                                    text: "The material is Argon Purged through the ladle.",
                                    font: "Calibri",
                                    size: 22,
                                }),
                            ],
                        }),
                        new docx.Paragraph({
                            bullet: { level: 0 },
                            children: [
                                new docx.TextRun({
                                    text: `ALL CLOSE CAST BILLETS OF ${heatData[4].gradeName} ARE COLORED ${tc.colorCode} AT ONE END.`,
                                    font: "Calibri",
                                    size: 22,
                                    bold: true,
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
                }, ],
            });
            // docx.Packer.toBuffer(doc5Cc).then((buffer) => {
            //   fs.writeFileSync(
            //     `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} CC ${heatData[0].sectionSize}MM.docx`,
            //     buffer
            //   );
            // });
            docx.Packer.toBuffer(doc5Cc).then((buffer) => {
                const filename = `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} CC ${heatData[0].sectionSize}MM.docx`;
                const sanitizedFilename = sanitize(filename);

                // Combine directory and sanitized filename
                const filePath = path.join(outputDirectory, sanitizedFilename);

                fs.writeFileSync(filePath, buffer);
                console.log(`File saved as: ${filePath}`);
            });
        }
        // let docxUrl = path.join(
        //     __dirname,
        //     "../",
        //     `${tc.buyerName} ${tc.billNo} ${heatData[0].sectionSize}MM.docx`
        // );
        // console.log(docxUrl);

        res.render("billets/closeCastTcPreview", { tc, tcList });
        // res.download(docxUrl, function(e) {
        //     if (e) {
        //         console.log(e);
        //     } else {
        //         console.log("Download Complete!!!");
        //     }
        // });
    })
);
router.get(
    "/:id/downloadDocxCc",
    catchAsync(async(req, res, next) => {
        console.log("Downloading....");
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
        // let docxUrl = path.join(
        //   __dirname,
        //   "../",
        //   `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} CC ${heatData[0].sectionSize}MM.docx`
        // );
        // res.download(docxUrl, function (e) {
        //   if (e) {
        //     console.log(e);
        //   } else {
        //     console.log("Download Complete!!!");
        //   }
        // });
        // Define the directory where the files are saved
        const outputDirectory = path.join(__dirname, "..", "output-files docx");

        // Construct the filename and sanitize it
        const filename = `${tc.buyerName} ${tc.billNo} ${heatData[0].gradeName} CC ${heatData[0].sectionSize}MM.docx`;
        const sanitizedFilename = sanitize(filename);

        // Combine directory and sanitized filename
        let docxUrl = path.join(outputDirectory, sanitizedFilename);

        // Download the file
        res.download(docxUrl, function(e) {
            if (e) {
                console.log(e);
            } else {
                console.log("Download Complete!!!");
            }
        });
    })
);

// SEARCH TC
router.get(
    "/tcList/search",
    catchAsync(async(req, res, next) => {
        const queryString = req.query.tc;
        let query = queryString.tcNo.toString();
        console.log(query);

        let queryList = query.split(" ");
        console.log(queryList);

        const tcResults = await Tc.find({ tcNo: { $in: queryList } });
        console.log(tcResults);

        res.render("billets/searchTc", { tcResults });
    })
);
router.get(
    "/:id/edit",
    catchAsync(async(req, res, next) => {
        const billet = await Billets.findById(req.params.id);
        res.render("billets/edit", { billet });
    })
);
router.post(
    "/",
    upload.fields([]),
    catchAsync(async(req, res, next) => {
        try {
            const existingBillet = await Billets.findOne(req.body.billet);

            if (existingBillet) {
                // If a matching document is found, you can choose to replace it or prompt the user
                // For simplicity, let's replace it directly
                await Billets.findOneAndReplace({ _id: existingBillet._id },
                    req.body.billet
                );
                console.log("Existing billet replaced:", existingBillet._id);
            } else {
                // If no matching document is found, create a new one
                let newBillet = new Billets(req.body.billet);
                await newBillet.save();
                console.log("New billet created:", newBillet._id);
            }

            if (req.body.oneMore == "one") {
                res.redirect("/billets/new");
            } else {
                res.redirect("/billets/list");
            }
        } catch (error) {
            console.error("Error processing billet:HI00000", error);
            // Handle the error and send an appropriate response
            res.status(500).send("Internal Server Error");
        }
    })
);
// router.post(
//   "/",
//   upload.fields([]),
//   catchAsync(async (req, res, next) => {
//     // Console.log("HIHIHIHI");
//     // let billet = new Billets(req.body.billet);
//     // console.log(billet.sectionSize);
//     // await billet.save();
//     // const newBillet = Billets.find({});
//     if (req.body.oneMore == "one") {
//       let billet = new Billets(req.body.billet);
//       console.log(billet.sectionSize);
//       await billet.save();
//       const newBillet = Billets.find({});
//       console.log(req.body.billet);
//       res.redirect("/billets/new");
//     } else {
//       let billet = new Billets(req.body.billet);
//       console.log(billet.sectionSize);
//       await billet.save();
//       const newBillet = Billets.find({});
//       console.log(req.body.billet);
//       res.redirect("/billets/list");
//     }
//   })
// );

router.put(
    "/:id",
    catchAsync(async(req, res, next) => {
        const { id } = req.params;
        await Billets.findByIdAndUpdate(id, {...req.body.billet });
        res.redirect("/billets/list");
    })
);
// DELETE HEAT
router.delete(
    "/:id/heat",
    catchAsync(async(req, res) => {
        const { id } = req.params;
        console.log(req.params);
        console.log("deleted Heat");
        await Billets.findByIdAndDelete(id);
        res.redirect("/billets/list");
    })
);
// DELETE HEAT
router.post(
    "/deleteSelected",
    catchAsync(async(req, res) => {
        console.log("hihihi");
        const selectedBillets = req.body.selectedBillets;

        try {
            // Perform the deletion logic for multiple billets using selectedBillets array
            for (const id of selectedBillets) {
                console.log(id);
                await Billets.findByIdAndDelete(id);
            }

            res.redirect("/billets/list"); // Redirect to the billets page after deletion
        } catch (error) {
            console.error(error);
            res.status(500).json({ success: false, error: "Internal Server Error" });
        }
    })
);
module.exports = router;