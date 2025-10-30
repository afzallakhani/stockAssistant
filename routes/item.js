const express = require("express");
const router = express.Router();
const path = require("path");
const { promisify } = require("util");
const { urlencoded, query, json } = require("express");
// const unlinkAsync = promisify(fs.unlink);
const catchAsync = require("../utils/catchAsync");
const methodOverride = require("method-override");
const multer = require("multer");
const puppeteer = require("puppeteer");

const fs = require("fs");

require("dotenv/config");
const ExpressError = require("../utils/ExpressError");
const ejs = require("ejs");

const Party = require("../models/partyMaster");
const Supplier = require("../models/supplier");
const Items = require("../models/elafStock");
const Transaction = require("../models/transaction");

const ItemCategories = require("../models/itemCategories");
const Images = require("../models/images");
const Billets = require("../models/billetList");
const multerStorage = require("../utils/multerStorage");
const bodyParser = require("body-parser");
const mongoose = require("mongoose");

const validateItem = require("../utils/validateItem");
const events = require("events");
const eventEmitter = new events.EventEmitter();

let upload = multer({ storage: multerStorage });

// router.get(
//   "/",
//   catchAsync(async (req, res) => {
//     const items = await Items.find({}).populate("itemImage");
//     const images = await Images.find({});
//     const itemCategories = await ItemCategories.find({});
//     const itemSuppliers = await Supplier.find({});
//     console.log({
//       items: items.length,
//       itemCategories: itemCategories.length,
//       itemSuppliers: itemSuppliers.length,
//     });
//     res.render("items/allItems", { items, itemCategories, itemSuppliers });
//   })
// );
router.get(
  "/",
  catchAsync(async (req, res) => {
    const items = await Items.find({}).populate("itemImage");
    const images = await Images.find({});
    const itemCategories = await ItemCategories.find({});
    const itemSuppliers = await Supplier.find({});

    console.log({
      items: items.length,
      itemCategories: itemCategories.length,
      itemSuppliers: itemSuppliers.length,
    });

    res.render("items/allItems", {
      items,
      itemCategories: itemCategories || [],
      itemSuppliers: itemSuppliers || [],
      query: req.query || {}, // ✅ Fix added
    });
  })
);

router.get("/", (req, res) => {
  console.log("hi");
  res.render("home");
});
router.get(
  "/spares",
  catchAsync(async (req, res) => {
    const items = await Items.find({}).populate("itemImage");
    const images = await Images.find({});
    console.log("hi");
    res.render("items/spares");
  })
);
router.get(
  "/new",
  catchAsync(async (req, res) => {
    const itemCategories = await ItemCategories.find({});
    const itemSuppliers = await Supplier.find({}, "supplierName supplierCity");
    console.log(itemSuppliers);
    res.render("items/new", { itemCategories, itemSuppliers });
  })
);
// router.get(
//     "/outwards",
//     catchAsync(async(req, res) => {
//         const itemCategories = await ItemCategories.find({});
//         res.render("items/new", { itemCategories });
//     })
// );
// Assuming you have an Express router setup
// and an 'Item' model for your items in the database.

// GET route to display the outwards form
router.get("/outwards", async (req, res) => {
  try {
    const items = await Items.find({}).populate("itemImage");
    const itemCategories = await ItemCategories.find({});
    const itemSuppliers = await Supplier.find({}, "supplierName supplierCity");

    res.render("items/outwards", {
      items,
      itemCategories,
      itemSuppliers,
    });
  } catch (error) {
    console.error("Error fetching items for outwards log:", error);
    res.status(500).send("Error fetching items.");
  }
});

// GET: Lend items page
router.get("/lend", async (req, res) => {
  try {
    const items = await Items.find({});
    res.render("items/lend", { items });
  } catch (error) {
    console.error("Error loading lend page:", error);
    res.status(500).send("Error loading lend page");
  }
});

// POST: Handle lend transaction
router.post(
  "/lend",
  catchAsync(async (req, res) => {
    const { items } = req.body;

    for (const id in items) {
      const lendQty = parseInt(items[id].lendQty, 10);
      if (lendQty > 0) {
        const item = await Items.findById(id);
        if (item) {
          const stockBefore = item.itemQty;
          const stockAfter = stockBefore - lendQty;

          item.itemQty = stockAfter;
          await item.save();

          await new Transaction({
            itemId: item._id,
            type: "lend",
            quantity: lendQty,
            stockBefore,
            stockAfter,
          }).save();
        }
      }
    }

    res.redirect("/items/transactions");
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
// router.get("/search", async (req, res) => {
//   const queryString = req.query.item;
//   const query = queryString.itemName;

//   // for (let query of queryString) {
//   //     //     console.log(typeof queryString);
//   //     console.log(query.itemName);
//   // }

//   const item = await Items.find({
//     $or: [
//       { itemName: { $regex: query, $options: "i" } },
//       { itemCategoryName: { $regex: query, $options: "i" } },
//     ],
//   }).populate("itemImage");

//   res.render("items/search", { item });
// });
router.get("/search", async (req, res) => {
  const query =
    req.query && req.query.item && req.query.item.itemName
      ? req.query.item.itemName.trim()
      : "";

  const items = await Items.find({
    $or: [
      { itemName: { $regex: query, $options: "i" } },
      { itemCategoryName: { $regex: query, $options: "i" } },
      { itemSupplier: { $regex: query, $options: "i" } },
      { itemDescription: { $regex: query, $options: "i" } },
    ],
  }).populate("itemImage");

  res.render("items/search", { items });
});

// add new item to database
// router.post(
//   "/",
//   upload.array("item[itemImage]"),
//   validateItem,
//   catchAsync(async (req, res, next) => {
//     // if (!req.body.item) throw new ExpressError("Invalid Item Data", 400);

//     let item = new Items(req.body.item);
//     let image = new Images();
//     image.contentType = req.file.mimetype;
//     image.data = fs.readFileSync(
//       path.join(__dirname, "..", "views", "images", req.file.filename)
//     );
//     image.path = req.file.path;
//     image.name = req.file.originalname;
//     const imagePath = req.file.path;
//     await item.save();
//     const currentItem = await Items.findById(item.id);
//     currentItem.itemImage.push(image);
//     await image.save();
//     await currentItem.save();
//     await fs.unlink(imagePath, (err) => {
//       if (err) {
//         console.error(err);
//         return;
//       } else {
//         return;
//       }
//     });
//     const initialStock = item.itemQty;
//     await new Transaction({
//       itemId: item._id,
//       type: "initial",
//       quantity: initialStock,
//       stockBefore: 0, // Stock is 0 before creation
//       stockAfter: initialStock,
//     }).save();
//     res.redirect("/items");
//   })
// );
router.post(
  "/",
  upload.array("item[itemImage]"),
  validateItem,
  catchAsync(async (req, res) => {
    let item = new Items(req.body.item);
    await item.save();

    if (req.files && req.files.length > 0) {
      for (const file of req.files) {
        const image = new Images({
          contentType: file.mimetype,
          data: fs.readFileSync(
            path.join(__dirname, "..", "views", "images", file.filename)
          ),
          path: file.path,
          name: file.originalname,
        });
        await image.save();
        item.itemImage.push(image);
        fs.unlink(file.path, () => {});
      }
      await item.save();
    }

    const initialStock = item.itemQty;
    await new Transaction({
      itemId: item._id,
      type: "initial",
      quantity: initialStock,
      stockBefore: 0,
      stockAfter: initialStock,
    }).save();

    res.redirect("/items");
  })
);
router.get(
  "/:id/view",
  catchAsync(async (req, res) => {
    const item = await Items.findById(req.params.id).populate("itemImage");
    if (!item) return res.status(404).send("Item not found");
    res.render("items/viewItem", { item });
  })
);

// GET route to display the inwards form
router.get(
  "/inwards",
  catchAsync(async (req, res) => {
    const items = await Items.find({}).populate("itemImage");
    const itemCategories = await ItemCategories.find({});
    const itemSuppliers = await Supplier.find({}, "supplierName supplierCity");

    res.render("items/inwards", {
      items,
      itemCategories,
      itemSuppliers,
    });
  })
);

// POST route to handle inward stock updates
router.post(
  "/inward-stock",
  catchAsync(async (req, res) => {
    const { items } = req.body;

    for (const itemData of items) {
      const inwardsQty = parseInt(itemData.inwardsQty, 10);

      if (inwardsQty > 0) {
        const item = await Items.findById(itemData.id);
        if (item) {
          const stockBefore = item.itemQty;
          const stockAfter = stockBefore + inwardsQty;

          // Update the item's quantity
          item.itemQty = stockAfter;
          await item.save();

          // Create a transaction log
          await new Transaction({
            itemId: item._id,
            type: "inward",
            quantity: inwardsQty,
            stockBefore: stockBefore,
            stockAfter: stockAfter,
          }).save();
        }
      }
    }

    res.redirect("/items");
  })
);

// GET route to view all transactions
// Replace the existing GET "/transactions" route
// router.get(
//     "/transactions",
//     catchAsync(async(req, res) => {
//         // Fetch all items for the filter dropdown
//         const allItems = await Items.find({});

//         // Build the filter query based on request parameters
//         const filter = {};
//         if (req.query.item) {
//             filter.itemId = req.query.item;
//         }

//         if (req.query.startDate) {
//             filter.createdAt = {
//                 ...filter.createdAt,
//                 $gte: new Date(req.query.startDate),
//             };
//         }

//         if (req.query.endDate) {
//             // Add one day to the end date to include all transactions on that day
//             let endDate = new Date(req.query.endDate);
//             endDate.setDate(endDate.getDate() + 1);
//             filter.createdAt = {...filter.createdAt, $lt: endDate };
//         }

//         // Fetch transactions using the filter
//         const transactions = await Transaction.find(filter)
//             .populate("itemId")
//             .sort({ createdAt: -1 });

//         // Render the page, passing the transactions, all items for the dropdown, and the query parameters
//         res.render("items/transactions", {
//             transactions,
//             allItems,
//             query: req.query,
//         });
//     })
// );
router.get(
  "/transactions",
  catchAsync(async (req, res) => {
    // Fetch all items for the filter dropdown
    const allItems = await Items.find();

    // Build the filter query based on request parameters
    const filter = {};
    if (req.query.item) {
      filter.itemId = req.query.item;
    }
    if (req.query.startDate) {
      filter.createdAt = {
        ...filter.createdAt,
        $gte: new Date(req.query.startDate),
      };
    }
    if (req.query.endDate) {
      // Add one day to the end date to include all transactions on that day
      let endDate = new Date(req.query.endDate);
      endDate.setDate(endDate.getDate() + 1);
      filter.createdAt = { ...filter.createdAt, $lt: endDate };
    }

    // Fetch transactions using the filter, sorted by most recent
    const transactions = await Transaction.find(filter)
      .populate("itemId")
      .sort({ createdAt: -1 });

    // --- New Grouping Logic ---
    // Group transactions by date
    const groupedTransactions = transactions.reduce((acc, tx) => {
      // Create a date key in 'YYYY-MM-DD' format for reliable grouping
      const dateKey = new Date(tx.createdAt).toISOString().split("T")[0];
      if (!acc[dateKey]) {
        acc[dateKey] = [];
      }
      acc[dateKey].push(tx);
      return acc;
    }, {});

    // Get an array of date keys and sort them from newest to oldest
    const sortedDates = Object.keys(groupedTransactions).sort(
      (a, b) => new Date(b) - new Date(a)
    );

    // Render the page, passing the new grouped data structures
    res.render("items/transactions", {
      groupedTransactions,
      sortedDates,
      allItems,
      query: req.query,
    });
  })
);
router.get(
  "/insights",
  catchAsync(async (req, res) => {
    const { itemId, startDate, endDate } = req.query;
    const allItems = await Items.find({}).sort({ itemName: 1 });
    let insights = null;
    let transactions = [];

    // Only run the query if all filters are present
    if (itemId && startDate && endDate) {
      const start = new Date(startDate);
      const end = new Date(endDate);
      // Adjust end date to include all transactions on that day
      end.setHours(23, 59, 59, 999);

      // 1. Fetch individual transactions for the detailed table
      transactions = await Transaction.find({
        itemId: itemId,
        createdAt: { $gte: start, $lte: end },
      })
        .populate("itemId")
        .sort({ createdAt: -1 });

      // 2. Use MongoDB Aggregation Pipeline to calculate insights
      const aggregationResult = await Transaction.aggregate([
        // Stage 1: Filter documents for the specific item and date range
        {
          $match: {
            itemId: new mongoose.Types.ObjectId(itemId),
            createdAt: { $gte: start, $lte: end },
          },
        },
        // Stage 2: Group by transaction type and sum the quantities
        {
          $group: {
            _id: "$type", // Group by 'inward', 'outward', 'initial', etc.
            totalQuantity: { $sum: "$quantity" },
          },
        },
      ]);

      // 3. Process the aggregation results to create a simple insights object
      let totalInward = 0;
      let totalOutward = 0;
      aggregationResult.forEach((result) => {
        if (result._id === "inward") {
          totalInward = result.totalQuantity;
        } else if (result._id === "outward") {
          totalOutward = result.totalQuantity;
        }
      });

      // Total Consumption is the sum of all 'outward' transactions in the period
      const totalConsumption = totalOutward;

      insights = {
        totalInward,
        totalOutward,
        totalConsumption,
      };
    }

    // Render the insights page with the data
    res.render("items/insights", {
      allItems,
      insights,
      transactions,
      query: req.query, // Pass query params back to pre-fill the form
    });
  })
);
// router.get(
//   "/consumption",
//   catchAsync(async (req, res) => {
//     const { itemId, category, supplier, startDate, endDate, mode } = req.query;

//     const items = await Items.find({}).sort({ itemName: 1 });
//     const categories = await ItemCategories.find({}).sort({
//       itemCategoryName: 1,
//     });
//     const suppliers = await Supplier.find({}).sort({ supplierName: 1 });

//     // 🔹 Match outward transactions (consumption only)
//     var baseMatch = { type: "outward" };

//     // 🔹 Add date filters
//     if (startDate || endDate) {
//       baseMatch.createdAt = {};
//       if (startDate) baseMatch.createdAt.$gte = new Date(startDate);
//       if (endDate)
//         baseMatch.createdAt.$lte = new Date(
//           new Date(endDate).setHours(23, 59, 59, 999)
//         );
//     }

//     // 🔹 Aggregate base data
//     var transactions = await Transaction.aggregate([
//       { $match: baseMatch },
//       {
//         $lookup: {
//           from: "allitems",
//           localField: "itemId",
//           foreignField: "_id",
//           as: "item",
//         },
//       },
//       { $unwind: "$item" },
//       {
//         $match: Object.assign(
//           {},
//           itemId ? { "item._id": new mongoose.Types.ObjectId(itemId) } : {},
//           category ? { "item.itemCategoryName": category } : {},
//           supplier ? { "item.itemSupplier": supplier } : {}
//         ),
//       },
//       {
//         $group: {
//           _id: "$item._id",
//           itemName: { $first: "$item.itemName" },
//           category: { $first: "$item.itemCategoryName" },
//           supplier: { $first: "$item.itemSupplier" },
//           totalUsed: { $sum: "$quantity" },
//           lastUsed: { $max: "$createdAt" },
//         },
//       },
//       { $sort: { itemName: 1 } },
//     ]);

//     // 🔹 Adjust totals if mode = "sinceLastInward"
//     if (mode === "sinceLastInward") {
//       for (var i = 0; i < transactions.length; i++) {
//         var tx = transactions[i];
//         var lastInward = await Transaction.findOne({
//           itemId: tx._id,
//           type: "inward",
//         })
//           .sort({ createdAt: -1 })
//           .lean();

//         if (lastInward) {
//           var afterInward = await Transaction.aggregate([
//             {
//               $match: {
//                 itemId: tx._id,
//                 type: "outward",
//                 createdAt: { $gte: lastInward.createdAt },
//               },
//             },
//             {
//               $group: { _id: null, totalUsed: { $sum: "$quantity" } },
//             },
//           ]);
//           tx.totalUsed = afterInward.length > 0 ? afterInward[0].totalUsed : 0;
//         }
//       }
//     }

//     // ---- Compute Current Stock + Average Usage ----
//     var msPerDay = 1000 * 60 * 60 * 24;
//     var now = new Date();
//     var totalUsedSum = 0;
//     var totalStock = 0;
//     var totalAvg = 0;
//     var enriched = [];

//     for (var i = 0; i < transactions.length; i++) {
//       var tx = transactions[i];
//       var item = await Items.findById(tx._id).lean();
//       var stock = item && item.itemQty ? item.itemQty : 0;

//       // Determine period range
//       var end = endDate
//         ? new Date(new Date(endDate).setHours(23, 59, 59, 999))
//         : now;
//       var start = null;

//       if (startDate) {
//         start = new Date(startDate);
//       } else if (mode === "sinceLastInward") {
//         var lastInward2 = await Transaction.findOne({
//           itemId: tx._id,
//           type: "inward",
//         })
//           .sort({ createdAt: -1 })
//           .lean();
//         if (lastInward2 && lastInward2.createdAt)
//           start = new Date(lastInward2.createdAt);
//       }

//       if (!start) {
//         var firstOutward = await Transaction.findOne({
//           itemId: tx._id,
//           type: "outward",
//         })
//           .sort({ createdAt: 1 })
//           .lean();
//         if (firstOutward && firstOutward.createdAt)
//           start = new Date(firstOutward.createdAt);
//       }

//       if (!start) start = tx.lastUsed ? new Date(tx.lastUsed) : new Date();

//       var diffMs = end.getTime() - start.getTime();
//       var daysRange = Math.round(diffMs / msPerDay);
//       if (isNaN(daysRange) || daysRange < 1) daysRange = 1;

//       var avgUsage = tx.totalUsed > 0 ? tx.totalUsed / daysRange : 0;

//       totalUsedSum += tx.totalUsed || 0;
//       totalStock += stock;
//       totalAvg += avgUsage;

//       enriched.push({
//         _id: tx._id,
//         itemName: tx.itemName,
//         category: tx.category,
//         supplier: tx.supplier,
//         totalUsed: tx.totalUsed || 0,
//         avgUsage: avgUsage.toFixed(2),
//         currentStock: stock,
//         lastUsed: tx.lastUsed,
//       });
//     }

//     var totalItems = enriched.length;
//     var avgUsageOverall =
//       totalItems > 0 ? (totalAvg / totalItems).toFixed(2) : "0.00";

//     // 🔹 Pass everything to view
//     res.render("items/consumption", {
//       items,
//       categories,
//       suppliers,
//       transactions: enriched,
//       summary: {
//         totalUsedSum,
//         totalStock,
//         avgUsageOverall,
//       },
//       query: req.query,
//     });
//   })
// );
// router.get(
//   "/consumption",
//   catchAsync(async (req, res) => {
//     var itemId = req.query.itemId;
//     var category = req.query.category;
//     var supplier = req.query.supplier;
//     var startDate = req.query.startDate;
//     var endDate = req.query.endDate;
//     var mode = req.query.mode;

//     const items = await Items.find({}).sort({ itemName: 1 });
//     const categories = await ItemCategories.find({}).sort({
//       itemCategoryName: 1,
//     });
//     const suppliers = await Supplier.find({}).sort({ supplierName: 1 });

//     var baseMatch = { type: "outward" };
//     if (startDate || endDate) {
//       baseMatch.createdAt = {};
//       if (startDate) baseMatch.createdAt.$gte = new Date(startDate);
//       if (endDate) {
//         var e = new Date(endDate);
//         e.setHours(23, 59, 59, 999);
//         baseMatch.createdAt.$lte = e;
//       }
//     }

//     // Fetch outward transactions
//     var transactions = await Transaction.aggregate([
//       { $match: baseMatch },
//       {
//         $lookup: {
//           from: "allitems",
//           localField: "itemId",
//           foreignField: "_id",
//           as: "item",
//         },
//       },
//       { $unwind: "$item" },
//       {
//         $match: Object.assign(
//           {},
//           itemId ? { "item._id": new mongoose.Types.ObjectId(itemId) } : {},
//           category ? { "item.itemCategoryName": category } : {},
//           supplier ? { "item.itemSupplier": supplier } : {}
//         ),
//       },
//       {
//         $group: {
//           _id: "$item._id",
//           itemName: { $first: "$item.itemName" },
//           category: { $first: "$item.itemCategoryName" },
//           supplier: { $first: "$item.itemSupplier" },
//           totalUsed: { $sum: "$quantity" },
//           lastUsed: { $max: "$createdAt" },
//         },
//       },
//       { $sort: { itemName: 1 } },
//     ]);

//     // Mode: sinceLastInward
//     if (mode === "sinceLastInward") {
//       for (var i = 0; i < transactions.length; i++) {
//         var tx = transactions[i];
//         var lastInward = await Transaction.findOne({
//           itemId: tx._id,
//           type: "inward",
//         })
//           .sort({ createdAt: -1 })
//           .lean();

//         if (lastInward) {
//           var afterInward = await Transaction.aggregate([
//             {
//               $match: {
//                 itemId: tx._id,
//                 type: "outward",
//                 createdAt: { $gte: lastInward.createdAt },
//               },
//             },
//             { $group: { _id: null, totalUsed: { $sum: "$quantity" } } },
//           ]);

//           tx.totalUsed = afterInward.length > 0 ? afterInward[0].totalUsed : 0;
//         }
//       }
//     }

//     // Fetch stock data
//     const itemStocks = {};
//     const allItems = await Items.find({}).select("_id itemQty itemName");
//     allItems.forEach(function (itm) {
//       itemStocks[itm._id.toString()] = itm.itemQty;
//     });

//     // Fetch billet data to calculate open/close heats
//     const allBillets = await Billets.find({});
//     const openHeats = allBillets.filter(function (b) {
//       return !b.ce || b.ce === "";
//     }).length;
//     const closeHeats = allBillets.filter(function (b) {
//       return b.ce && b.ce !== "";
//     }).length;
//     const totalHeats = openHeats + closeHeats;

//     // Calculate overall average usage per day and per heat
//     var totalUsedOverall = 0;
//     transactions.forEach(function (tx) {
//       totalUsedOverall += tx.totalUsed;
//       tx.currentStock = itemStocks[tx._id.toString()] || 0;
//     });

//     var avgPerDay = 0;
//     if (startDate && endDate) {
//       var d1 = new Date(startDate);
//       var d2 = new Date(endDate);
//       var diffDays = Math.round((d2 - d1) / (1000 * 60 * 60 * 24)) + 1;
//       if (diffDays > 0) avgPerDay = (totalUsedOverall / diffDays).toFixed(2);
//     }

//     var avgPerHeat =
//       totalHeats > 0 ? (totalUsedOverall / totalHeats).toFixed(2) : 0;

//     res.render("items/consumption", {
//       items,
//       categories,
//       suppliers,
//       transactions,
//       query: req.query,
//       avgPerDay,
//       avgPerHeat,
//       openHeats,
//       closeHeats,
//       totalHeats,
//     });
//   })
// );
router.get(
  "/consumption",
  catchAsync(async (req, res) => {
    var itemId = req.query.itemId;
    var category = req.query.category;
    var supplier = req.query.supplier;
    var startDate = req.query.startDate;
    var endDate = req.query.endDate;
    var mode = req.query.mode;
    var heatType = req.query.heatType;

    const items = await Items.find({}).sort({ itemName: 1 });
    const categories = await ItemCategories.find({}).sort({
      itemCategoryName: 1,
    });
    const suppliers = await Supplier.find({}).sort({ supplierName: 1 });

    var baseMatch = { type: "outward" };

    // ================================
    //  🔹 Determine date range based on mode
    // ================================
    var dateRangeStart = null;
    var dateRangeEnd = endDate ? new Date(endDate) : new Date();

    if (mode === "sinceLastInward" && itemId) {
      var lastInward = await Transaction.findOne({
        itemId: itemId,
        type: "inward",
      })
        .sort({ createdAt: -1 })
        .lean();

      if (lastInward) {
        dateRangeStart = new Date(lastInward.createdAt);
      }
    } else if (mode === "" && itemId) {
      // Since Start mode
      var firstTransaction = await Transaction.findOne({
        itemId: itemId,
      })
        .sort({ createdAt: 1 })
        .lean();

      if (firstTransaction) {
        dateRangeStart = new Date(firstTransaction.createdAt);
      }
    }

    // If manual date filters are provided, override
    if (startDate) dateRangeStart = new Date(startDate);
    if (endDate) {
      dateRangeEnd = new Date(endDate);
      dateRangeEnd.setHours(23, 59, 59, 999);
    }

    if (dateRangeStart) {
      baseMatch.createdAt = { $gte: dateRangeStart, $lte: dateRangeEnd };
    }

    // ================================
    //  🔹 Fetch Outward Transactions
    // ================================
    var transactions = await Transaction.aggregate([
      { $match: baseMatch },
      {
        $lookup: {
          from: "allitems",
          localField: "itemId",
          foreignField: "_id",
          as: "item",
        },
      },
      { $unwind: "$item" },
      {
        $match: Object.assign(
          {},
          itemId ? { "item._id": new mongoose.Types.ObjectId(itemId) } : {},
          category ? { "item.itemCategoryName": category } : {},
          supplier ? { "item.itemSupplier": supplier } : {}
        ),
      },
      {
        $group: {
          _id: "$item._id",
          itemName: { $first: "$item.itemName" },
          category: { $first: "$item.itemCategoryName" },
          supplier: { $first: "$item.itemSupplier" },
          totalUsed: { $sum: "$quantity" },
          lastUsed: { $max: "$createdAt" },
        },
      },
      { $sort: { itemName: 1 } },
    ]);

    // ================================
    //  🔹 Handle Mode: sinceLastInward (adjust totals)
    // ================================
    if (mode === "sinceLastInward") {
      for (var i = 0; i < transactions.length; i++) {
        var tx = transactions[i];
        var lastInward = await Transaction.findOne({
          itemId: tx._id,
          type: "inward",
        })
          .sort({ createdAt: -1 })
          .lean();

        if (lastInward) {
          var afterInward = await Transaction.aggregate([
            {
              $match: {
                itemId: tx._id,
                type: "outward",
                createdAt: { $gte: lastInward.createdAt },
              },
            },
            { $group: { _id: null, totalUsed: { $sum: "$quantity" } } },
          ]);

          tx.totalUsed = afterInward.length > 0 ? afterInward[0].totalUsed : 0;
        }
      }
    }

    // ================================
    //  🔹 Current Stock Info
    // ================================
    const itemStocks = {};
    const allItems = await Items.find({}).select("_id itemQty itemName");
    allItems.forEach(function (itm) {
      itemStocks[itm._id.toString()] = itm.itemQty;
    });

    // ================================
    //  🔹 Fetch Billets for same date range & heat type
    // ================================
    var billetFilter = {};

    if (dateRangeStart && dateRangeEnd) {
      billetFilter.createdAt = {
        $gte: dateRangeStart,
        $lte: dateRangeEnd,
      };
    }

    // Heat Type
    if (heatType === "open") {
      billetFilter.$or = [{ ce: null }, { ce: "" }];
    } else if (heatType === "close") {
      billetFilter.ce = { $nin: [null, ""] };
    }

    const billets = await Billets.find(billetFilter);
    const openHeats = billets.filter(function (b) {
      return !b.ce || b.ce === "";
    }).length;
    const closeHeats = billets.filter(function (b) {
      return b.ce && b.ce !== "";
    }).length;
    const totalHeats = billets.length;

    // ================================
    //  🔹 Calculate Averages
    // ================================
    var totalUsedOverall = 0;
    transactions.forEach(function (tx) {
      totalUsedOverall += tx.totalUsed;
      tx.currentStock = itemStocks[tx._id.toString()] || 0;
    });

    var avgPerDay = 0;
    if (dateRangeStart && dateRangeEnd) {
      var diffDays =
        Math.round(
          (dateRangeEnd.getTime() - dateRangeStart.getTime()) /
            (1000 * 60 * 60 * 24)
        ) + 1;
      if (diffDays > 0) avgPerDay = (totalUsedOverall / diffDays).toFixed(2);
    }

    var avgPerHeat =
      totalHeats > 0 ? (totalHeats / totalUsedOverall).toFixed(2) : 0;

    // ================================
    //  🔹 Render View
    // ================================
    res.render("items/consumption", {
      items,
      categories,
      suppliers,
      transactions,
      query: req.query,
      avgPerDay,
      avgPerHeat,
      openHeats,
      closeHeats,
      totalHeats,
    });
  })
);

// *** THIS IS THE ROUTE TO FIX THE ERROR ***
// It correctly defines the DELETE method for your transactions
router.delete(
  "/transactions/:id",
  catchAsync(async (req, res) => {
    const { id } = req.params;
    const transaction = await Transaction.findById(id);

    if (!transaction) {
      throw new ExpressError("Transaction not found", 404);
    }
    if (transaction.type === "initial") {
      throw new ExpressError("Cannot reverse an initial stock entry.", 400);
    }

    const item = await Items.findById(transaction.itemId);
    if (item) {
      if (transaction.type === "inward") {
        // Revert inward → reduce stock
        item.itemQty -= transaction.quantity;
      } else if (transaction.type === "outward") {
        // Revert outward → increase stock
        item.itemQty += transaction.quantity;
      } else if (transaction.type === "lend") {
        // Revert lend → increase stock (because item was lent out)
        item.itemQty += transaction.quantity;
      } else if (transaction.type === "return") {
        // Revert return → decrease stock (because returned item was added back)
        item.itemQty -= transaction.quantity;
      }

      await item.save();
    }

    await Transaction.findByIdAndDelete(id);
    res.redirect("/items/transactions");
  })
);
// ➕ Log "return" transaction when item is brought back
router.post(
  "/transactions/:id/return",
  catchAsync(async (req, res) => {
    const { id } = req.params;
    const { quantity } = req.body;

    const lendTx = await Transaction.findById(id).populate("itemId");
    if (!lendTx) throw new ExpressError("Lend transaction not found", 404);
    if (lendTx.type !== "lend")
      throw new ExpressError("Only lend transactions can be returned", 400);

    const item = lendTx.itemId;
    const returnQty = parseInt(quantity) || lendTx.quantity; // default to full quantity

    // ✅ Update stock
    item.itemQty += returnQty;
    await item.save();

    // ✅ Log return transaction
    await new Transaction({
      itemId: item._id,
      type: "return",
      quantity: returnQty,
      stockBefore: item.itemQty - returnQty,
      stockAfter: item.itemQty,
      borrower: lendTx.borrower || null,
    }).save();

    // ✅ Mark the lend transaction as returned
    lendTx.returned = true;
    await lendTx.save();

    res.redirect("/items/transactions");
  })
);

// --- Other routes ---
router.get(
  "/new",
  catchAsync(async (req, res) => {
    const itemCategories = await ItemCategories.find({});
    res.render("items/new", { itemCategories });
  })
);

// router.get(
//     "/transactions",
//     catchAsync(async(req, res) => {
//         const transactions = await Transaction.find({})
//             .populate("itemId") // Get item details
//             .sort({ createdAt: -1 }); // Show latest first
//         res.render("items/transactions", { transactions });
//     })
// );

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

router.delete(
  "/category/:id",
  catchAsync(async (req, res) => {
    const { id } = req.params;
    await ItemCategories.findByIdAndDelete(id);
    res.redirect("/items/category");
  })
);
router.get(
  "/:id/edit",
  catchAsync(async (req, res, next) => {
    const item = await Items.findById(req.params.id).populate("itemImage");
    const itemCategories = await ItemCategories.find({});
    const itemSuppliers = await Supplier.find({}, "supplierName supplierCity");
    res.render("items/edit", { item, itemCategories, itemSuppliers });
    // next(e);
  })
);
router.put(
  "/:id",
  upload.array("item[itemImage]"),
  catchAsync(async (req, res) => {
    const { id } = req.params;
    const item = await Items.findById(id).populate("itemImage");

    // ✅ Update basic item fields
    Object.assign(item, req.body.item);
    await item.save();

    // ✅ Handle new image uploads (add on top of existing)
    if (req.files && req.files.length > 0) {
      for (const file of req.files) {
        const image = new Images({
          contentType: file.mimetype,
          data: fs.readFileSync(
            path.join(__dirname, "..", "views", "images", file.filename)
          ),
          path: file.path,
          name: file.originalname,
        });
        await image.save();
        item.itemImage.push(image);
        fs.unlink(file.path, () => {}); // remove temp file
      }
      await item.save();
    }

    res.redirect(`/items`);
  })
);

// Replace your existing POST "/update-stock" route with this

router.post(
  "/update-stock",
  catchAsync(async (req, res) => {
    const { items } = req.body;

    for (const itemData of items) {
      const outwardsQty = parseInt(itemData.outwardsQty, 10);

      // Only process if there's a quantity to move outwards
      if (outwardsQty > 0) {
        // Find the item to get its current stock
        const item = await Items.findById(itemData.id);
        if (item) {
          const stockBefore = item.itemQty;
          const stockAfter = stockBefore - outwardsQty;

          // Update the item's quantity
          item.itemQty = stockAfter;
          await item.save();

          // Create a transaction log
          await new Transaction({
            itemId: item._id,
            type: "outward",
            quantity: outwardsQty,
            stockBefore: stockBefore,
            stockAfter: stockAfter,
          }).save();
        }
      }
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

// Add this route below other routes in items.js

router.get("/suggestions", async (req, res) => {
  try {
    const query = req.query.q || "";
    if (!query.trim()) return res.json([]);

    const items = await Items.find({
      $or: [
        { itemName: { $regex: query, $options: "i" } },
        { itemCategoryName: { $regex: query, $options: "i" } },
        { itemSupplier: { $regex: query, $options: "i" } },
        { itemDescription: { $regex: query, $options: "i" } },
      ],
    })
      .limit(8)
      .select("itemName itemCategoryName");

    res.json(items);
  } catch (err) {
    console.error("Error fetching suggestions:", err);
    res.status(500).json({ error: "Server error" });
  }
});

// const path = require("path");
// const fs = require("fs");

// ------------------------------------
// 🔹 STOCK REPORT PDF
// ------------------------------------
router.get(
  "/report/stock",
  catchAsync(async (req, res) => {
    var category = req.query.category;
    var supplier = req.query.supplier;

    var match = {};
    if (category && category !== "all") match.itemCategoryName = category;
    if (supplier && supplier !== "all") match.itemSupplier = supplier;

    const items = await Items.find(match).sort({ itemName: 1 }).lean();

    // render HTML first
    const html = await ejs.renderFile(
      path.join(__dirname, "../views/items/pdf_stock.ejs"),
      { items, category, supplier },
      { async: false }
    );

    // generate PDF
    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });

    const pdfPath = path.join(
      __dirname,
      `../public/reports/Stock_Report_${Date.now()}.pdf`
    );

    await page.pdf({
      path: pdfPath,
      format: "A4",
      landscape: true,
      printBackground: true,
    });

    await browser.close();

    res.download(pdfPath, () => fs.unlinkSync(pdfPath));
  })
);

// ------------------------------------
// 🔹 CONSUMPTION REPORT PDF
// ------------------------------------
router.get(
  "/report/consumption",
  catchAsync(async (req, res) => {
    var category = req.query.category;
    var supplier = req.query.supplier;
    var startDate = req.query.startDate;
    var endDate = req.query.endDate;
    var mode = req.query.mode;

    var match = { type: "outward" };

    // Date logic
    var dateRangeStart = null;
    var dateRangeEnd = endDate ? new Date(endDate) : new Date();

    if (mode === "sinceLastInward") {
      const lastInward = await Transaction.findOne({ type: "inward" })
        .sort({ createdAt: -1 })
        .lean();
      if (lastInward) dateRangeStart = new Date(lastInward.createdAt);
    } else if (mode === "sinceLastOutward") {
      const lastOutward = await Transaction.findOne({ type: "outward" })
        .sort({ createdAt: -1 })
        .lean();
      if (lastOutward) dateRangeStart = new Date(lastOutward.createdAt);
    } else if (mode === "sinceStart") {
      const firstTx = await Transaction.findOne({})
        .sort({ createdAt: 1 })
        .lean();
      if (firstTx) dateRangeStart = new Date(firstTx.createdAt);
    }

    if (startDate) dateRangeStart = new Date(startDate);
    if (endDate) {
      dateRangeEnd = new Date(endDate);
      dateRangeEnd.setHours(23, 59, 59, 999);
    }

    if (dateRangeStart) {
      match.createdAt = { $gte: dateRangeStart, $lte: dateRangeEnd };
    }

    // Filter by category/supplier
    const transactions = await Transaction.aggregate([
      { $match: match },
      {
        $lookup: {
          from: "allitems",
          localField: "itemId",
          foreignField: "_id",
          as: "item",
        },
      },
      { $unwind: "$item" },
      {
        $match: Object.assign(
          {},
          category && category !== "all"
            ? { "item.itemCategoryName": category }
            : {},
          supplier && supplier !== "all"
            ? { "item.itemSupplier": supplier }
            : {}
        ),
      },
      {
        $group: {
          _id: "$item._id",
          itemName: { $first: "$item.itemName" },
          category: { $first: "$item.itemCategoryName" },
          supplier: { $first: "$item.itemSupplier" },
          totalUsed: { $sum: "$quantity" },
          lastUsed: { $max: "$createdAt" },
        },
      },
      { $sort: { itemName: 1 } },
    ]);

    // Render HTML
    const html = await ejs.renderFile(
      path.join(__dirname, "../views/items/pdf_consumption.ejs"),
      {
        transactions,
        category,
        supplier,
        mode,
        startDate,
        endDate,
      },
      { async: false }
    );

    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });

    const pdfPath = path.join(
      __dirname,
      `../public/reports/Consumption_Report_${Date.now()}.pdf`
    );

    await page.pdf({
      path: pdfPath,
      format: "A4",
      landscape: true,
      printBackground: true,
    });

    await browser.close();

    res.download(pdfPath, () => fs.unlinkSync(pdfPath));
  })
);

module.exports = router;
