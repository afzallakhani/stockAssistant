const express = require("express");
const router = express.Router();
const path = require("path");
const { promisify } = require("util");
const { urlencoded, query, json } = require("express");
// const unlinkAsync = promisify(fs.unlink);
const runBackup = require("../utils/backupHelper");
const getMonthlyConsumption = require("../utils/getMonthlyConsumption");

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
const DB_NAME = "stockAssistant";
const BACKUP_DIR = path.join(__dirname, "../backups");
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
router.get("/backup-now", (req, res) => {
  runBackup("manual");
  res.send("Backup triggered successfully!");
});
// GET item stock & avg monthly consumption

router.get("/:id/po-info", async (req, res) => {
  const itemId = req.params.id;

  const item = await Items.findById(itemId).lean();
  if (!item) {
    return res.json({});
  }

  /* ================= CURRENT STOCK ================= */
  const txs = await Transaction.find({ itemId: item._id });

  let inward = 0;
  let outward = 0;

  txs.forEach((tx) => {
    if (tx.type === "inward") inward += tx.quantity;
    if (tx.type === "outward" || tx.type === "lend") outward += tx.quantity;
  });

  /* ================= AVG / MONTH (UTIL) ================= */
  const monthly = await getMonthlyConsumption(item);

  res.json({
    unit: item.itemUnit,
    currentStock: item.itemQty,
    avgMonthly: monthly.perMonth,
  });
});

// // ✅ Manual Backup Trigger (this is what "Backup Now" button calls)
// router.post("/utility/backup", (req, res) => {
//   runBackup("manual");
//   res.redirect("/items/utility");
// });

// // ✅ Restore from a selected backup folder
// router.post("/utility/restore", (req, res) => {
//   const { backupFolder } = req.body;
//   if (!backupFolder) return res.send("⚠️ No backup folder selected!");

//   const restorePath = path.join(BACKUP_DIR, backupFolder, DB_NAME);
//   const cmd = `mongorestore --db=${DB_NAME} --gzip "${restorePath}" --drop`;

//   console.log(`♻️ Restoring from ${restorePath}...`);
//   exec(cmd, (error) => {
//     if (error) {
//       console.error(`❌ Restore failed: ${error.message}`);
//       return res.send("❌ Restore failed. Check console for details.");
//     }
//     console.log(`✅ Restore completed from ${backupFolder}`);
//     res.send(`✅ Database restored successfully from backup: ${backupFolder}`);
//   });
// });

// // ✅ Delete a selected backup folder
// router.post("/utility/delete", (req, res) => {
//   const { backupFolder } = req.body;
//   if (!backupFolder) return res.send("⚠️ No backup selected!");
//   const deletePath = path.join(BACKUP_DIR, backupFolder);
//   fs.rmSync(deletePath, { recursive: true, force: true });
//   res.redirect("/items/utility");
// });
// Manual Backup Trigger
router.post("/utility/backup", (req, res) => {
  runBackup("manual");
  req.flash("success", "✅ Backup started successfully!");
  res.redirect("/items/utility");
});

// Restore Backup
router.post("/utility/restore", (req, res) => {
  const file = req.body.backupFile;
  const BACKUP_PATH = path.join(__dirname, "../backups");
  const fullPath = path.join(BACKUP_PATH, file);

  const cmd = `"C:\\Program Files\\MongoDB\\Tools\\100\\bin\\mongorestore.exe" --gzip --archive="${fullPath}" --drop`;

  exec(cmd, (err) => {
    if (err) {
      req.flash("error", "Restore failed");
      return res.redirect("/items/utility");
    }
    req.flash("success", "Database restored successfully");
    res.redirect("/items/utility");
  });
});

// Delete Backup
router.post("/utility/delete", (req, res) => {
  const file = req.body.backupFile;
  const BACKUP_PATH = path.join(__dirname, "../backups");

  try {
    fs.unlinkSync(path.join(BACKUP_PATH, file));
    req.flash("success", "Backup deleted successfully");
  } catch (err) {
    req.flash("error", "Failed to delete backup");
  }

  res.redirect("/items/utility");
});

// Utility Dashboard Page

router.get("/utility", (req, res) => {
  const BACKUP_PATH = path.join(__dirname, "../backups");

  let backups = [];

  if (fs.existsSync(BACKUP_PATH)) {
    backups = fs
      .readdirSync(BACKUP_PATH)
      .filter((f) => f.endsWith(".archive.gz"))
      .map((f) => {
        const stat = fs.statSync(path.join(BACKUP_PATH, f));
        return {
          name: f,
          date: stat.mtime,
        };
      })
      .sort((a, b) => b.date - a.date);
  }

  res.render("items/utility", {
    backups,
    success: req.flash("success"),
    error: req.flash("error"),
  });
});

// router.get(
//   "/",
//   catchAsync(async (req, res) => {
//     const items = await Items.find({}).populate("itemImage");
//     for (let item of items) {
//       const dc = await getMonthlyConsumption(item);
//       item.monthlyConsumption = dc.perMonth;
//       item.stockDaysLeft =
//         dc.perMonth > 0 ? Math.floor((item.itemQty / dc.perMonth) * 30) : "∞";
//     }
//     const images = await Images.find({});
//     const itemCategories = await ItemCategories.find({});
//     const itemSuppliers = await Supplier.find({});

//     console.log({
//       items: items.length,
//       itemCategories: itemCategories.length,
//       itemSuppliers: itemSuppliers.length,
//     });

//     res.render("items/allItems", {
//       items,
//       itemCategories: itemCategories || [],
//       itemSuppliers: itemSuppliers || [],
//       query: req.query || {}, // ✅ Fix added
//     });
//   }),
// );
router.get(
  "/",
  catchAsync(async (req, res) => {
    const page = parseInt(req.query.page) || 1;
    const limit = 10;
    const skip = (page - 1) * limit;

    // 🔹 Total count for pagination
    const totalItems = await Items.countDocuments();

    // 🔹 Fetch only 10 items
    const items = await Items.find({})
      .populate("itemImage")
      .select(
        "itemName itemCategoryName itemSupplier itemQty itemUnit itemDescription createdAt itemImage",
      )
      .skip(skip)
      .limit(limit)
      .lean();

    // 🔹 Aggregation for consumption (still fast)
    const consumptionData = await Transaction.aggregate([
      { $match: { type: "OUTWARD" } },
      {
        $group: {
          _id: "$item",
          totalQty: { $sum: "$qty" },
        },
      },
    ]);

    const consumptionMap = {};
    for (let i = 0; i < consumptionData.length; i++) {
      consumptionMap[consumptionData[i]._id.toString()] =
        consumptionData[i].totalQty;
    }

    // 🔹 Attach data
    for (let i = 0; i < items.length; i++) {
      const id = items[i]._id.toString();
      const totalOutward = consumptionMap[id] || 0;

      items[i].monthlyConsumption = totalOutward;

      if (totalOutward > 0) {
        items[i].stockDaysLeft = Math.floor(
          (items[i].itemQty / totalOutward) * 30,
        );
      } else {
        items[i].stockDaysLeft = "∞";
      }

      items[i].formattedItemDate = new Date(
        items[i].createdAt,
      ).toLocaleDateString("en-GB");
    }

    const totalPages = Math.ceil(totalItems / limit);

    const itemCategories = await ItemCategories.find({})
      .select("itemCategoryName")
      .lean();

    const itemSuppliers = await Supplier.find({}).select("supplierName").lean();

    res.render("items/allItems", {
      items,
      itemCategories,
      itemSuppliers,
      currentPage: page,
      totalPages,
      query: req.query || {},
    });
  }),
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
  }),
);
router.get(
  "/new",
  catchAsync(async (req, res) => {
    const itemCategories = await ItemCategories.find({});
    const itemSuppliers = await Supplier.find({}, "supplierName supplierCity");
    console.log(itemSuppliers);
    res.render("items/new", { itemCategories, itemSuppliers });
  }),
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
    const items = await Items.find({}).populate("itemImage").lean();
    const itemCategories = await ItemCategories.find({});
    const itemSuppliers = await Supplier.find({}, "supplierName supplierCity");
    items.forEach((item) => {
      if (item.itemImage && item.itemImage.length > 0) {
        item.base64Image = `data:image/${
          item.itemImage[0].contentType
        };base64,${item.itemImage[0].data.toString("base64")}`;
      } else {
        item.base64Image =
          "https://via.placeholder.com/60x60.png?text=No+Image";
      }
    });
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
router.get(
  "/lend",
  catchAsync(async (req, res) => {
    const items = await Items.find({}).populate("itemImage").lean();
    const itemCategories = await ItemCategories.find({});
    const itemSuppliers = await Supplier.find({}, "supplierName supplierCity");
    items.forEach((item) => {
      if (item.itemImage && item.itemImage.length > 0) {
        item.base64Image = `data:image/${
          item.itemImage[0].contentType
        };base64,${item.itemImage[0].data.toString("base64")}`;
      } else {
        item.base64Image =
          "https://via.placeholder.com/60x60.png?text=No+Image";
      }
    });
    res.render("items/lend", {
      items,
      itemCategories,
      itemSuppliers,
    });
  }),
);

// POST: Handle lend transaction
router.post(
  "/lend",
  catchAsync(async (req, res) => {
    const { items } = req.body;

    for (const itemData of items) {
      const lendQty = parseInt(itemData.lendQty, 10);
      if (lendQty > 0) {
        const item = await Items.findById(itemData.id);
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
            remarks: itemData.remarks || "",
          }).save();
        }
      }
    }

    res.redirect("/items/transactions");
  }),
);

router.get(
  "/category",
  catchAsync(async (req, res) => {
    const category = await ItemCategories.find({});
    res.render("items/category", { category });
  }),
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
            path.join(__dirname, "..", "views", "images", file.filename),
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
  }),
);
router.get(
  "/:id/view",
  catchAsync(async (req, res) => {
    const item = await Items.findById(req.params.id).populate("itemImage");
    if (!item) return res.status(404).send("Item not found");
    res.render("items/viewItem", { item });
  }),
);

// GET route to display the inwards form
router.get(
  "/inwards",
  catchAsync(async (req, res) => {
    const items = await Items.find({}).populate("itemImage").lean();
    const itemCategories = await ItemCategories.find({});
    const itemSuppliers = await Supplier.find({}, "supplierName supplierCity");
    items.forEach((item) => {
      if (item.itemImage && item.itemImage.length > 0) {
        item.base64Image = `data:image/${
          item.itemImage[0].contentType
        };base64,${item.itemImage[0].data.toString("base64")}`;
      } else {
        item.base64Image =
          "https://via.placeholder.com/60x60.png?text=No+Image";
      }
    });
    res.render("items/inwards", {
      items,
      itemCategories,
      itemSuppliers,
    });
  }),
);

// POST route to handle inward stock updates
router.post(
  "/inwards",
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
            stockBefore,
            stockAfter,
            remarks: itemData.remarks || "",
          }).save();
        }
      }
    }

    res.redirect("/items");
  }),
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
    // Fetch all items for the filter dropdow
    const page = parseInt(req.query.page || 1);
    const limit = 50;

    const allItems = await Items.find().select("itemName _id");

    // Build the filter query based on request parameters
    const filter = {};
    if (req.query.itemId) {
      filter.itemId = req.query.itemId;
    }
    if (req.query.type) {
      filter.type = req.query.type;
    }

    if (req.query.startDate || req.query.endDate) {
      filter.createdAt = {};
      if (req.query.startDate) {
        filter.createdAt.$gte = new Date(req.query.startDate);
      }
      if (req.query.endDate) {
        const end = new Date(req.query.endDate);
        end.setHours(23, 59, 59, 999);
        filter.createdAt.$lte = end;
      }
    }

    // Fetch transactions using the filter, sorted by most recent
    const transactions = await Transaction.find(filter)
      .populate("itemId", "itemName")
      .sort({ createdAt: -1 })
      .skip((page - 1) * limit)
      .limit(500);
    const totalCount = await Transaction.countDocuments(filter);
    const totalPages = Math.ceil(totalCount / limit);

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
      (a, b) => new Date(b) - new Date(a),
    );

    // Render the page, passing the new grouped data structures
    // res.render("items/transactions", {
    //   groupedTransactions,
    //   sortedDates,
    //   query: req.query,
    //   items: allItems, // ✅ now available in EJS
    // });
    res.render("items/transactions", {
      groupedTransactions,
      sortedDates,
      query: req.query,
      items: allItems,
      page,
      totalPages,
    });
  }),
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
  }),
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
// latestrouter.get(
//     "/consumption",
//     catchAsync(async(req, res) => {
//         const { itemId, category, supplier, startDate, endDate, mode, heatType } =
//         req.query;

//         // ---------------------------------
//         // 🔹 Masters
//         // ---------------------------------
//         const items = await Items.find({}).sort({ _id: 1 });
//         const categories = await ItemCategories.find({});
//         const suppliers = await Supplier.find({});

//         // ---------------------------------
//         // 🔹 Date Range (Manual)
//         // ---------------------------------
//         const hasManualRange = startDate || endDate;
//         const rangeStart = startDate ? new Date(startDate) : null;
//         const rangeEnd = endDate ?
//             new Date(new Date(endDate).setHours(23, 59, 59, 999)) :
//             new Date();

//         // ---------------------------------
//         // 🔹 Base Match (initial aggregation)
//         // ---------------------------------
//         const baseMatch = { type: { $in: ["outward", "lend"] } };

//         if (hasManualRange) {
//             baseMatch.createdAt = {};
//             if (rangeStart) baseMatch.createdAt.$gte = rangeStart;
//             if (rangeEnd) baseMatch.createdAt.$lte = rangeEnd;
//         }

//         // ---------------------------------
//         // 🔹 Aggregate items (structure only)
//         // ---------------------------------
//         let transactions = await Transaction.aggregate([
//             { $match: baseMatch },
//             {
//                 $lookup: {
//                     from: "allitems",
//                     localField: "itemId",
//                     foreignField: "_id",
//                     as: "item",
//                 },
//             },
//             { $unwind: "$item" },
//             {
//                 $match: {
//                     ...(itemId && { "item._id": new mongoose.Types.ObjectId(itemId) }),
//                     ...(category && { "item.itemCategoryName": category }),
//                     ...(supplier && { "item.itemSupplier": supplier }),
//                 },
//             },
//             {
//                 $group: {
//                     _id: "$item._id",
//                     itemName: { $first: "$item.itemName" },
//                     category: { $first: "$item.itemCategoryName" },
//                     supplier: { $first: "$item.itemSupplier" },
//                     totalUsed: { $sum: "$quantity" }, // ⚠️ will be corrected below
//                     lastUsed: { $max: "$createdAt" },
//                 },
//             },
//             { $sort: { _id: 1 } },
//         ]);

//         // ---------------------------------
//         // 🔹 Initial & Last Inward Maps
//         // ---------------------------------
//         const initialMap = {};
//         const lastInwardMap = {};

//         const initials = await Transaction.aggregate([
//             { $match: { type: "initial" } },
//             { $sort: { createdAt: 1 } },
//             {
//                 $group: {
//                     _id: "$itemId",
//                     date: { $first: "$createdAt" },
//                     qty: { $first: "$quantity" },
//                 },
//             },
//         ]);

//         initials.forEach((r) => {
//             initialMap[r._id.toString()] = r;
//         });

//         const lastInwards = await Transaction.aggregate([
//             { $match: { type: "inward" } },
//             { $sort: { createdAt: -1 } },
//             {
//                 $group: {
//                     _id: "$itemId",
//                     date: { $first: "$createdAt" },
//                     qty: { $first: "$quantity" },
//                 },
//             },
//         ]);

//         lastInwards.forEach((r) => {
//             lastInwardMap[r._id.toString()] = r;
//         });

//         // ---------------------------------
//         // 🔹 Image / Unit / Stock Maps
//         // ---------------------------------
//         const itemDocs = await Items.find({})
//             .populate("itemImage")
//             .select("_id itemImage itemUnit itemQty createdAt")
//             .lean();

//         const imageMap = {};
//         const unitMap = {};
//         const stockMap = {};
//         const itemCreatedMap = {}; // ✅ ADD
//         const itemQtyMap = {}; // ✅ ADD
//         itemDocs.forEach((i) => {
//             imageMap[i._id.toString()] =
//                 i.itemImage && i.itemImage.length ?
//                 "data:image/" +
//                 i.itemImage[0].contentType +
//                 ";base64," +
//                 i.itemImage[0].data.toString("base64") :
//                 "https://via.placeholder.com/400x300.png?text=No+Image";

//             unitMap[i._id.toString()] = i.itemUnit || "";
//             stockMap[i._id.toString()] = i.itemQty || 0;
//         });

//         // ---------------------------------
//         // 🔹 PER-ITEM PERIOD + CORRECT TOTAL USED
//         // ---------------------------------
//         for (const tx of transactions) {
//             const id = tx._id.toString();

//             let periodStart = null;
//             const periodEnd = rangeEnd;

//             if (hasManualRange) {
//                 periodStart = rangeStart;
//             } else if (mode === "sinceLastInward") {
//                 periodStart = lastInwardMap[id] ? lastInwardMap[id].date : null;
//             } else {
//                 if (initialMap[id]) {
//                     periodStart = initialMap[id].date;
//                 } else {
//                     // 🔥 FALLBACK: item created date
//                     periodStart = itemCreatedMap[id] || null;
//                 }
//             }

//             // 🔥 RECALCULATE TOTAL USED FOR THIS ITEM
//             if (periodStart) {
//                 const sumResult = await Transaction.aggregate([{
//                         $match: {
//                             itemId: tx._id,
//                             type: { $in: ["outward", "lend"] },
//                             createdAt: { $gte: periodStart, $lte: periodEnd },
//                         },
//                     },
//                     { $group: { _id: null, total: { $sum: "$quantity" } } },
//                 ]);

//                 tx.totalUsed = sumResult.length > 0 ? sumResult[0].total : 0;
//             }

//             tx.periodStart = periodStart;
//             tx.periodEnd = periodEnd;

//             tx.base64Image = imageMap[id];
//             tx.unit = unitMap[id];
//             tx.currentStock = stockMap[id];
//             tx.initialDate = initialMap[id] ?
//                 initialMap[id].date :
//                 itemCreatedMap[id] || null;

//             tx.initialQty = initialMap[id] ? initialMap[id].qty : itemQtyMap[id] || 0;

//             tx.lastInwards = lastInwardMap[id] ? lastInwardMap[id].date : null;
//             tx.lastInwardQty = lastInwardMap[id] ? lastInwardMap[id].qty : 0;
//             // ---------------------------------
//             // 🔹 MONTHLY CONSUMPTION LOGIC (FIX)
//             // ---------------------------------

//             let months = 0;

//             if (periodStart && periodEnd) {
//                 months =
//                     (periodEnd.getFullYear() - periodStart.getFullYear()) * 12 +
//                     (periodEnd.getMonth() - periodStart.getMonth()) +
//                     Math.max(1, (periodEnd.getDate() - periodStart.getDate()) / 30);
//             }

//             months = Math.max(1, Number(months.toFixed(2)));

//             tx.consumptionMonths = months;

//             // 🔹 Avg per month
//             tx.avgPerMonth =
//                 months > 0 ? Number((tx.totalUsed / months).toFixed(2)) : 0;

//             // 🔹 Stock months left
//             tx.stockMonthsLeft =
//                 tx.avgPerMonth > 0 ?
//                 Number((tx.currentStock / tx.avgPerMonth).toFixed(1)) :
//                 null;

//             // 🔹 Stock out date
//             if (tx.stockMonthsLeft) {
//                 const outDate = new Date();
//                 outDate.setMonth(outDate.getMonth() + Math.ceil(tx.stockMonthsLeft));
//                 tx.stockOutDate = outDate;
//             } else {
//                 tx.stockOutDate = null;
//             }
//         }

//         // ---------------------------------
//         // 🔹 HEAT SUMMARY (USES FIXED TOTALS)
//         // ---------------------------------
//         const billetFilter = {};

//         if (hasManualRange) {
//             billetFilter.createdAt = { $gte: rangeStart, $lte: rangeEnd };
//         }

//         if (heatType === "open") {
//             billetFilter.$or = [{ ce: null }, { ce: "" }];
//         } else if (heatType === "close") {
//             billetFilter.ce = { $nin: [null, ""] };
//         }

//         const billets = await Billets.find(billetFilter);

//         const openHeats = billets.filter((b) => !b.ce).length;
//         const closeHeats = billets.filter((b) => b.ce).length;
//         const totalHeats = billets.length;

//         let totalUsedOverall = 0;
//         transactions.forEach((t) => {
//             totalUsedOverall += Number(t.totalUsed || 0);
//         });

//         const avgPerHeat =
//             totalHeats > 0 ? (totalUsedOverall / totalHeats).toFixed(2) : 0;

//         // ---------------------------------
//         // 🔹 Render
//         // ---------------------------------
//         res.render("items/consumption", {
//             items,
//             categories,
//             suppliers,
//             transactions,
//             query: req.query,

//             totalHeats,
//             openHeats,
//             closeHeats,
//             avgPerHeat,
//         });
//     })
// );
router.get(
  "/consumption",
  catchAsync(async (req, res) => {
    const { itemId, category, supplier, startDate, endDate, mode, heatType } =
      req.query;

    // ---------------------------------
    // 🔹 Masters (ITEM-DRIVEN)
    // ---------------------------------
    const items = await Items.find({})
      .populate("itemImage")
      .sort({ _id: 1 })
      .lean();

    const categories = await ItemCategories.find({});
    const suppliers = await Supplier.find({});

    // ---------------------------------
    // 🔹 Date Range
    // ---------------------------------
    const hasManualRange = startDate || endDate;
    const rangeStart = startDate ? new Date(startDate) : null;
    const rangeEnd = endDate
      ? new Date(new Date(endDate).setHours(23, 59, 59, 999))
      : new Date();

    // ---------------------------------
    // 🔹 Initial & Last Inward Maps
    // ---------------------------------
    const initialMap = {};
    const lastInwardMap = {};

    const initials = await Transaction.aggregate([
      { $match: { type: "initial" } },
      { $sort: { createdAt: 1 } },
      {
        $group: {
          _id: "$itemId",
          date: { $first: "$createdAt" },
          qty: { $first: "$quantity" },
        },
      },
    ]);

    initials.forEach((r) => {
      initialMap[r._id.toString()] = r;
    });

    const lastInwards = await Transaction.aggregate([
      { $match: { type: "inward" } },
      { $sort: { createdAt: -1 } },
      {
        $group: {
          _id: "$itemId",
          date: { $first: "$createdAt" },
          qty: { $first: "$quantity" },
        },
      },
    ]);

    lastInwards.forEach((r) => {
      lastInwardMap[r._id.toString()] = r;
    });

    // ---------------------------------
    // 🔹 Build Consumption Rows
    // ---------------------------------
    const transactions = [];

    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      const id = item._id.toString();

      // 🔹 Filters
      if (itemId && id !== itemId) continue;
      if (category && item.itemCategoryName !== category) continue;
      if (supplier && item.itemSupplier !== supplier) continue;

      // ---------------------------------
      // 🔹 Period Start Logic (NODE 15 SAFE)
      // ---------------------------------
      let periodStart = null;
      const periodEnd = rangeEnd;

      if (hasManualRange) {
        periodStart = rangeStart;
      } else if (mode === "sinceLastInward") {
        if (lastInwardMap[id] && lastInwardMap[id].date) {
          periodStart = lastInwardMap[id].date;
        } else {
          periodStart = item.createdAt;
        }
      } else {
        if (initialMap[id] && initialMap[id].date) {
          periodStart = initialMap[id].date;
        } else {
          periodStart = item.createdAt;
        }
      }

      // ---------------------------------
      // 🔹 Total Used
      // ---------------------------------
      let totalUsed = 0;

      if (periodStart) {
        const sum = await Transaction.aggregate([
          {
            $match: {
              itemId: item._id,
              type: { $in: ["outward", "lend"] },
              createdAt: { $gte: periodStart, $lte: periodEnd },
            },
          },
          { $group: { _id: null, total: { $sum: "$quantity" } } },
        ]);

        if (sum.length > 0) totalUsed = sum[0].total;
      }
      // ---------------------------------
      // 🔹 Last Outward / Lend (FIX)
      // ---------------------------------
      let lastUsed = null;
      let lastOutwardQty = 0;

      const lastOutward = await Transaction.findOne({
        itemId: item._id,
        type: { $in: ["outward", "lend"] },
      })
        .sort({ createdAt: -1 })
        .select("createdAt quantity")
        .lean();

      if (lastOutward) {
        lastUsed = lastOutward.createdAt;
        lastOutwardQty = lastOutward.quantity;
      }

      // ---------------------------------
      // 🔹 Monthly Consumption
      // ---------------------------------
      let months = 1;

      if (periodStart && periodEnd) {
        months =
          (periodEnd.getFullYear() - periodStart.getFullYear()) * 12 +
          (periodEnd.getMonth() - periodStart.getMonth()) +
          (periodEnd.getDate() - periodStart.getDate()) / 30;

        months = Math.max(1, Number(months.toFixed(2)));
      }

      const avgPerMonth = Number((totalUsed / months).toFixed(2));

      // ---------------------------------
      // 🔹 Stock Calculations
      // ---------------------------------
      let stockMonthsLeft = null;
      let stockOutDate = null;

      if (avgPerMonth > 0) {
        stockMonthsLeft = Number((item.itemQty / avgPerMonth).toFixed(1));

        stockOutDate = new Date();
        stockOutDate.setMonth(
          stockOutDate.getMonth() + Math.ceil(stockMonthsLeft),
        );
      }
      // ---------------------------------
      // 🔹 Reorder Qty (1 FULL MONTH)
      // ---------------------------------
      let reorderQty = 0;

      if (avgPerMonth > 0) {
        reorderQty = Number(avgPerMonth.toFixed(1));
      }

      // ---------------------------------
      // 🔹 Image
      // ---------------------------------
      let base64Image = "https://via.placeholder.com/400x300.png?text=No+Image";

      if (item.itemImage && item.itemImage.length > 0) {
        base64Image =
          "data:image/" +
          item.itemImage[0].contentType +
          ";base64," +
          item.itemImage[0].data.toString("base64");
      }

      // ---------------------------------
      // 🔹 Push Row
      // ---------------------------------
      transactions.push({
        _id: item._id,
        itemName: item.itemName,
        category: item.itemCategoryName,
        supplier: item.itemSupplier,

        base64Image,
        unit: item.itemUnit || "",
        currentStock: item.itemQty || 0,

        initialDate:
          initialMap[id] && initialMap[id].date
            ? initialMap[id].date
            : item.createdAt,

        initialQty:
          initialMap[id] && initialMap[id].qty
            ? initialMap[id].qty
            : item.itemQty || 0,

        lastInwards:
          lastInwardMap[id] && lastInwardMap[id].date
            ? lastInwardMap[id].date
            : null,

        lastInwardQty:
          lastInwardMap[id] && lastInwardMap[id].qty
            ? lastInwardMap[id].qty
            : 0,

        periodStart,
        periodEnd,
        lastUsed,
        lastOutwardQty,
        totalUsed,
        avgPerMonth,
        stockMonthsLeft,
        stockOutDate,
        reorderQty,
      });
    }

    // ---------------------------------
    // 🔹 Heat Summary (UNCHANGED)
    // ---------------------------------
    const billetFilter = {};

    if (hasManualRange) {
      billetFilter.createdAt = { $gte: rangeStart, $lte: rangeEnd };
    }

    if (heatType === "open") {
      billetFilter.$or = [{ ce: null }, { ce: "" }];
    } else if (heatType === "close") {
      billetFilter.ce = { $nin: [null, ""] };
    }

    const billets = await Billets.find(billetFilter);

    const openHeats = billets.filter((b) => !b.ce).length;
    const closeHeats = billets.filter((b) => b.ce).length;
    const totalHeats = billets.length;

    let totalUsedOverall = 0;
    for (let i = 0; i < transactions.length; i++) {
      totalUsedOverall += Number(transactions[i].totalUsed || 0);
    }

    const avgPerHeat =
      totalHeats > 0 ? (totalUsedOverall / totalHeats).toFixed(2) : 0;

    // ---------------------------------
    // 🔹 Render
    // ---------------------------------
    res.render("items/consumption", {
      items,
      categories,
      suppliers,
      transactions,
      query: req.query,

      totalHeats,
      openHeats,
      closeHeats,
      avgPerHeat,
    });
  }),
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
  }),
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
  }),
);

// --- Other routes ---
router.get(
  "/new",
  catchAsync(async (req, res) => {
    const itemCategories = await ItemCategories.find({});
    res.render("items/new", { itemCategories });
  }),
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
  }),
);

router.delete(
  "/category/:id",
  catchAsync(async (req, res) => {
    const { id } = req.params;
    await ItemCategories.findByIdAndDelete(id);
    res.redirect("/items/category");
  }),
);
router.get(
  "/:id/edit",
  catchAsync(async (req, res, next) => {
    const item = await Items.findById(req.params.id).populate("itemImage");
    const itemCategories = await ItemCategories.find({});
    const itemSuppliers = await Supplier.find({}, "supplierName supplierCity");
    res.render("items/edit", { item, itemCategories, itemSuppliers });
    // next(e);
  }),
);
router.put(
  "/:id",
  upload.array("item[itemImage]"),
  catchAsync(async (req, res) => {
    const { id } = req.params;
    const item = await Items.findById(id).populate("itemImage");

    // ✅ Update text fields
    Object.assign(item, req.body.item);

    // ✅ DELETE selected old images
    if (req.body.deleteImages) {
      const idsToDelete = Array.isArray(req.body.deleteImages)
        ? req.body.deleteImages
        : [req.body.deleteImages];

      // Remove from DB
      await Images.deleteMany({ _id: { $in: idsToDelete } });

      // Remove references from item
      item.itemImage = item.itemImage.filter(
        (img) => !idsToDelete.includes(img._id.toString()),
      );
    }

    // ✅ ADD / REPLACE uploaded images
    if (req.files && req.files.length > 0) {
      // (Optional) clear all existing images if you want full replacement
      // await Images.deleteMany({ _id: { $in: item.itemImage } });
      // item.itemImage = [];

      for (const file of req.files) {
        const image = new Images({
          contentType: file.mimetype,
          data: fs.readFileSync(
            path.join(__dirname, "..", "views", "images", file.filename),
          ),
          path: file.path,
          name: file.originalname,
        });
        await image.save();
        item.itemImage.push(image);
        fs.unlink(file.path, () => {}); // cleanup temp file
      }
    }

    await item.save();
    // req.flash("success", "Item updated successfully!");

    res.redirect("/items");
  }),
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
            stockBefore,
            stockAfter,
            remarks: itemData.remarks || "",
          }).save();
        }
      }
    }

    res.redirect("/items");
  }),
);

router.delete(
  "/:id",
  catchAsync(async (req, res) => {
    const { id } = req.params;
    await Items.findByIdAndDelete(id);
    res.redirect("/items");
  }),
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
      { async: false },
    );

    // generate PDF
    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });

    const pdfPath = path.join(
      __dirname,
      `../public/reports/Stock_Report_${Date.now()}.pdf`,
    );

    await page.pdf({
      path: pdfPath,
      format: "A4",
      landscape: true,
      printBackground: true,
    });

    await browser.close();

    res.download(pdfPath, () => fs.unlinkSync(pdfPath));
  }),
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
      baseMatch.createdAt = {
        $gte: dateRangeStart,
        $lte: dateRangeEnd,
      };
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
            : {},
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
      { async: false },
    );

    const browser = await puppeteer.launch({ headless: true });
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });

    const pdfPath = path.join(
      __dirname,
      `../public/reports/Consumption_Report_${Date.now()}.pdf`,
    );

    await page.pdf({
      path: pdfPath,
      format: "A4",
      landscape: true,
      printBackground: true,
    });

    await browser.close();

    res.download(pdfPath, () => fs.unlinkSync(pdfPath));
  }),
);

router.post("/utility/backup-drive", (req, res) => {
  runBackup("manual-drive");
  req.flash("success", "☁️ Google Drive backup started!");
  res.redirect("/items/utility");
});
router.post("/category/ajax", upload.none(), async (req, res) => {
  try {
    const category = new ItemCategories({
      itemCategoryName: req.body.itemCategoryName,
    });
    await category.save();

    res.json({ success: true, category });
  } catch (err) {
    res.json({ success: false, error: err.message });
  }
});

module.exports = router;
