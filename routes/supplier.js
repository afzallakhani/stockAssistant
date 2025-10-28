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
const Supplier = require("../models/supplier");
const Items = require("../models/elafStock");
const ItemCategories = require("../models/itemCategories");
const Images = require("../models/images");
const Billets = require("../models/billetList");
const multerStorage = require("../utils/multerStorage");
const bodyParser = require("body-parser");

const validateItem = require("../utils/validateItem");
const events = require("events");
const eventEmitter = new events.EventEmitter();

let upload = multer({ storage: multerStorage });
router.get(
  "/allsuppliers",
  catchAsync(async (req, res) => {
    const suppliers = await Supplier.find({});
    res.render("supplier/allsuppliers", { suppliers });
  })
);

router.get(
  "/new",
  catchAsync(async (req, res) => {
    const itemCategories = await ItemCategories.find({});
    res.render("supplier/new", { itemCategories });
  })
);
// --- LIVE SEARCH API ---
router.get("/search", async (req, res) => {
  try {
    const keyword = req.query.q ? req.query.q.trim() : "";
    if (!keyword) return res.json([]);

    const regex = new RegExp(keyword, "i"); // case-insensitive

    const suppliers = await Supplier.find({
      $or: [
        { supplierName: regex },
        { supplierItemCategory: regex },
        { supplierCity: regex },
        { supplierState: regex },
      ],
    }).limit(10);

    res.json(suppliers);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Server error" });
  }
});
// 🧭 FULL SEARCH RESULT PAGE
router.get("/result", async (req, res) => {
  const keyword = req.query.q ? req.query.q.trim() : "";

  // if no query, just show all suppliers
  if (!keyword) {
    const suppliers = await Supplier.find({});
    return res.render("supplier/allsuppliers", { suppliers });
  }

  const regex = new RegExp(keyword, "i");

  const suppliers = await Supplier.find({
    $or: [
      { supplierName: regex },
      { supplierItemCategory: regex },
      { supplierCity: regex },
      { supplierState: regex },
    ],
  });

  res.render("supplier/allsuppliers", { suppliers });
});
// 🟢 View single supplier page
router.get("/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const supplier = await Supplier.findById(id);

    if (!supplier) {
      return res
        .status(404)
        .render("error", { err: { message: "Supplier not found" } });
    }

    res.render("supplier/search", { supplier });
  } catch (err) {
    console.error(err);
    res.status(500).render("error", { err });
  }
});
// Add New Supplier
router.post(
  "/",
  upload.fields([]),
  catchAsync(async (req, res, next) => {
    // if (!req.body.item) throw new ExpressError("Invalid Item Data", 400);

    let supplier = new Supplier(req.body.supplier);
    await supplier.save();
    res.redirect("/supplier/allsuppliers");
  })
);
router.put(
  "/:id",
  catchAsync(async (req, res, next) => {
    const { id } = req.params;
    console.log(id);

    await Supplier.findByIdAndUpdate(id, { ...req.body.supplier });
    res.redirect("/supplier/allsuppliers");
  })
);
router.get(
  "/:id/edit",
  catchAsync(async (req, res) => {
    const { id } = req.params;
    const supplier = await Supplier.findById(id);
    console.log(supplier);
    const itemCategories = await ItemCategories.find({});
    res.render("supplier/edit", { supplier, itemCategories });
  })
);
// Delete Supplier
router.delete(
  "/:id",
  catchAsync(async (req, res) => {
    const { id } = req.params;
    await Supplier.findByIdAndDelete(id);
    res.redirect("/supplier/allsuppliers");
  })
);
module.exports = router;
