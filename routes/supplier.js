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
    catchAsync(async(req, res) => {
        const suppliers = await Supplier.find({});
        res.render("supplier/allsuppliers", { suppliers });
    })
);
router.get(
    "/new",
    catchAsync(async(req, res) => {
        const itemCategories = await ItemCategories.find({});
        res.render("supplier/new", { itemCategories });
    })
);
// Add New Supplier
router.post(
    "/",
    upload.fields([]),
    catchAsync(async(req, res, next) => {
        // if (!req.body.item) throw new ExpressError("Invalid Item Data", 400);

        let supplier = new Supplier(req.body.supplier);
        await supplier.save();
        res.redirect("/supplier/allsuppliers");
    })
);

module.exports = router;