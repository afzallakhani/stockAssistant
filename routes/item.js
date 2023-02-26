const express = require("express");
const router = express.Router();
const path = require("path");
const { promisify } = require("util");
const { urlencoded, query } = require("express");
// const unlinkAsync = promisify(fs.unlink);
const catchAsync = require("../utils/catchAsync");
const methodOverride = require("method-override");
const multer = require("multer");
const fs = require("fs");
require("dotenv/config");
const ExpressError = require("../utils/ExpressError");
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
  "/",
  catchAsync(async (req, res) => {
    const items = await Items.find({}).populate("itemImage");
    const images = await Images.find({});
    res.render("items/allItems", { items });
  })
);

router.get(
  "/new",
  catchAsync(async (req, res) => {
    const itemCategories = await ItemCategories.find({});
    res.render("items/new", { itemCategories });
  })
);

router.get(
  "/category",
  catchAsync(async (req, res) => {
    const category = await ItemCategories.find({});
    res.render("items/category", { category });
  })
);

// searched items display
router.get("/search", async (req, res) => {
  const queryString = req.query.item;
  const query = queryString.itemName;

  // for (let query of queryString) {
  //     //     console.log(typeof queryString);
  //     console.log(query.itemName);
  // }

  const item = await Items.find({
    $or: [
      { itemName: { $regex: query, $options: "i" } },
      { itemCategoryName: { $regex: query, $options: "i" } },
    ],
  }).populate("itemImage");

  res.render("items/search", { item });
});

// add new item to database
router.post(
  "/",
  upload.single("item[itemImage]"),
  validateItem,
  catchAsync(async (req, res, next) => {
    // if (!req.body.item) throw new ExpressError("Invalid Item Data", 400);

    let item = new Items(req.body.item);
    let image = new Images();
    image.contentType = req.file.mimetype;
    image.data = fs.readFileSync(
      path.join(__dirname, "..", "views", "images", req.file.filename)
    );
    image.path = req.file.path;
    image.name = req.file.originalname;
    const imagePath = req.file.path;
    await item.save();
    const currentItem = await Items.findById(item.id);
    currentItem.itemImage.push(image);
    await image.save();
    await currentItem.save();
    await fs.unlink(imagePath, (err) => {
      if (err) {
        console.error(err);
        return;
      } else {
        return;
      }
    });
    res.redirect("/items");
  })
);

// add item categories
router.post(
  "/category",
  upload.fields([]),
  catchAsync(async (req, res, next) => {
    // if (!req.body.item) throw new ExpressError("Invalid Item Data", 400);

    let category = new ItemCategories(req.body.category);
    await category.save();
    res.redirect("/items/new");
  })
);
router.get(
  "/:id/edit",
  catchAsync(async (req, res, next) => {
    const item = await Items.findById(req.params.id).populate("itemImage");
    res.render("items/edit", { item });
    // next(e);
  })
);
router.put(
  "/:id",
  upload.single("item[itemImage]"),
  catchAsync(async (req, res, next) => {
    const { id } = req.params;
    const item = await Items.findById(id).populate("itemImage");
    const imageId = item.itemImage[0].id;
    let image = await Images.findById(imageId);
    if (req.file) {
      image.contentType = req.file.mimetype;
      image.data = fs.readFileSync(
        path.join(__dirname, "..", "views", "images", req.file.filename)
      );
      image.path = req.file.path;
      image.name = req.file.originalname;
      await Items.findByIdAndUpdate(id, { ...req.body.item });
      await Images.findByIdAndUpdate(imageId, image);
      const imagePath = req.file.path;
      await fs.unlink(imagePath, (err) => {
        if (err) {
          console.error(err);
          return;
        } else {
          return;
        }
      });
    } else {
      await Items.findByIdAndUpdate(id, { ...req.body.item });
    }
    res.redirect("/items");
  })
);

router.delete(
  "/:id",
  catchAsync(async (req, res) => {
    const { id } = req.params;
    await Items.findByIdAndDelete(id);
    res.redirect("/items");
  })
);
module.exports = router;
