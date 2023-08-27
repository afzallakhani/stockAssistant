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
const ItemCategories = require("../models/itemCategories");
const Images = require("../models/images");
const multerStorage = require("../utils/multerStorage");
const bodyParser = require("body-parser");

const validateItem = require("../utils/validateItem");
const events = require("events");
const eventEmitter = new events.EventEmitter();

let upload = multer({ storage: multerStorage });
router.get(
  "/allparties",
  catchAsync(async (req, res) => {
    const parties = await Party.find({});
    res.render("party/allparties", { parties });
  })
);
// router.get("/allparties", (req, res) => {
//   //   partyPath = path.join(__dirname, "views/party");
//   res.render("party/allParties");
// });

// Add New Party
router.post(
  "/",
  upload.fields([]),
  catchAsync(async (req, res, next) => {
    // if (!req.body.item) throw new ExpressError("Invalid Item Data", 400);

    let party = new Party(req.body.party);
    await party.save();
    console.log(party);
    res.redirect("/partymaster/allparties");
  })
);

// Delete Party
router.delete(
  "/:id",
  catchAsync(async (req, res) => {
    const { id } = req.params;
    await Party.findByIdAndDelete(id);
    res.redirect("/partymaster/allparties");
  })
);
// EDIT PARTY
router.put(
  "/:id",
  catchAsync(async (req, res, next) => {
    const { id } = req.params;
    await Party.findByIdAndUpdate(id, { ...req.body.party });
    res.redirect("/partymaster/allparties");
  })
);
router.get("/new", (req, res) => {
  res.render("party/new");
});
// router.get("/:id/edit", (req, res) => {
//   res.render("party/edit");
// });

router.get(
  "/:id/edit",
  catchAsync(async (req, res) => {
    const { id } = req.params;
    const party = await Party.findById(id);
    console.log(party);
    res.render("party/edit", { party });
  })
);

module.exports = router;
