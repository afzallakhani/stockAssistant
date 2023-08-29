const express = require("express");
const router = express.Router();
const path = require("path");
const { promisify } = require("util");
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
const Billets = require("../models/billetList");
const ItemCategories = require("../models/itemCategories");
const Images = require("../models/images");
const multerStorage = require("../utils/multerStorage");
const bodyParser = require("body-parser");

const validateItem = require("../utils/validateItem");
const events = require("events");
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
