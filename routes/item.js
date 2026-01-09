const express = require("express");
const router = express.Router();
const path = require("path");
const { promisify } = require("util");
const { urlencoded, query, json } = require("express");
// const unlinkAsync = promisify(fs.unlink);
const runBackup = require("../utils/backupHelper");
const getDailyConsumption = require("../utils/dailyConsumption");

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

router.get(
    "/",
    catchAsync(async(req, res) => {
        const items = await Items.find({}).populate("itemImage");
        for (let item of items) {
            const dc = await getDailyConsumption(item);
            item.dailyConsumption = dc.perDay;
            item.stockDaysLeft =
                dc.perDay > 0 ? (item.itemQty / dc.perDay).toFixed(1) : "∞";
        }
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
    catchAsync(async(req, res) => {
        const items = await Items.find({}).populate("itemImage");
        const images = await Images.find({});
        console.log("hi");
        res.render("items/spares");
    })
);
router.get(
    "/new",
    catchAsync(async(req, res) => {
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
router.get("/outwards", async(req, res) => {
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
    catchAsync(async(req, res) => {
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
    })
);

// POST: Handle lend transaction
router.post(
    "/lend",
    catchAsync(async(req, res) => {
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
    })
);

router.get(
    "/category",
    catchAsync(async(req, res) => {
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
router.get("/search", async(req, res) => {
    const query =
        req.query && req.query.item && req.query.item.itemName ?
        req.query.item.itemName.trim() :
        "";

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
    catchAsync(async(req, res) => {
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
    catchAsync(async(req, res) => {
        const item = await Items.findById(req.params.id).populate("itemImage");
        if (!item) return res.status(404).send("Item not found");
        res.render("items/viewItem", { item });
    })
);

// GET route to display the inwards form
router.get(
    "/inwards",
    catchAsync(async(req, res) => {
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
    })
);

// POST route to handle inward stock updates
router.post(
    "/inwards",
    catchAsync(async(req, res) => {
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
    catchAsync(async(req, res) => {
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
            (a, b) => new Date(b) - new Date(a)
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
    })
);
router.get(
    "/insights",
    catchAsync(async(req, res) => {
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
    catchAsync(async(req, res) => {
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

        var baseMatch = {
            type: { $in: ["outward", "lend"] }, // consumption types
        };

        // ================================
        // 🔹 DATE RANGE LOGIC (FINAL)
        // =============================

        // END date → today if not selected
        let dateRangeEnd = endDate ?
            new Date(new Date(endDate).setHours(23, 59, 59, 999)) :
            new Date();

        // START date priority:
        // 1️⃣ Manual startDate
        // 2️⃣ Since last inward (per item handled later)
        // 3️⃣ Since start (global earliest outward)

        let dateRangeStart = null;

        // 1️⃣ Manual date range
        if (startDate) {
            dateRangeStart = new Date(startDate);
        }

        // 2️⃣ Since start (global fallback)
        if (!dateRangeStart) {
            const firstOutward = await Transaction.findOne({
                    type: "outward",
                })
                .sort({ createdAt: 1 })
                .lean();

            if (firstOutward) {
                dateRangeStart = new Date(firstOutward.createdAt);
            }
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
                $match: Object.assign({},
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
                    var afterInward = await Transaction.aggregate([{
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
        const allItems = await Items.find({})
            .select("_id itemQty itemName itemUnit")
            .lean();
        allItems.forEach(function(itm) {
            itemStocks[itm._id.toString()] = itm.itemQty;
        });
        const unitMap = {};
        allItems.forEach((itm) => {
            unitMap[itm._id.toString()] = itm.itemUnit || "";
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
        const openHeats = billets.filter(function(b) {
            return !b.ce || b.ce === "";
        }).length;
        const closeHeats = billets.filter(function(b) {
            return b.ce && b.ce !== "";
        }).length;
        const totalHeats = billets.length;
        const allItemsWithImages = await Items.find({})
            .populate("itemImage")
            .lean();
        const imageMap = {};
        allItemsWithImages.forEach((itm) => {
            if (itm.itemImage && itm.itemImage.length > 0) {
                imageMap[itm._id.toString()] = `data:image/${
          itm.itemImage[0].contentType
        };base64,${itm.itemImage[0].data.toString("base64")}`;
            } else {
                imageMap[itm._id.toString()] =
                    "https://via.placeholder.com/60x60.png?text=No+Image";
            }
        });
        // ================================
        // 🔹 Initial Stock Info (NO N+1)
        // ================================
        const initialMap = {};

        const initials = await Transaction.aggregate([
            { $match: { type: "initial" } },
            { $sort: { createdAt: 1 } },
            {
                $group: {
                    _id: "$itemId",
                    initialDate: { $first: "$createdAt" },
                    initialQty: { $first: "$quantity" },
                },
            },
        ]);

        initials.forEach((row) => {
            initialMap[row._id.toString()] = {
                date: row.initialDate,
                qty: row.initialQty,
            };
        });

        // ===============================
        // 🔹 Last Inward DATE + QTY (NO N+1)
        // ===============================
        const lastInwardMap = {};

        const lastInwards = await Transaction.aggregate([
            { $match: { type: "inward" } },
            { $sort: { createdAt: -1 } },
            {
                $group: {
                    _id: "$itemId",
                    lastInwardDate: { $first: "$createdAt" },
                    lastInwardQty: { $first: "$quantity" },
                },
            },
        ]);

        lastInwards.forEach((row) => {
            lastInwardMap[row._id.toString()] = {
                date: row.lastInwardDate,
                qty: row.lastInwardQty,
            };
        });
        // ===============================
        // 🔹 Last Outward DATE + QTY (NO N+1)
        // ===============================
        const lastOutwardMap = {};

        const lastOutwards = await Transaction.aggregate([
            { $match: { type: { $in: ["outward", "lend"] } } },
            { $sort: { createdAt: -1 } },
            {
                $group: {
                    _id: "$itemId",
                    lastOutwardDate: { $first: "$createdAt" },
                    lastOutwardQty: { $first: "$quantity" },
                },
            },
        ]);

        lastOutwards.forEach((row) => {
            lastOutwardMap[row._id.toString()] = {
                date: row.lastOutwardDate,
                qty: row.lastOutwardQty,
            };
        });

        // ================================
        //  🔹 Calculate Per-Item Consumption
        // ================================
        const leadTime = 7;
        const enrichedTransactions = [];

        for (const rawTx of transactions) {
            const tx = {...rawTx };

            // 🔹 Image
            tx.base64Image =
                imageMap[tx._id.toString()] ||
                "https://via.placeholder.com/60x60.png?text=No+Image";

            // 🔹 Stock
            tx.currentStock = itemStocks[tx._id.toString()] || 0;

            let start = dateRangeStart;
            let end = dateRangeEnd || new Date();
            // 🔹 Last Inward Info
            const inwardInfo = lastInwardMap[tx._id.toString()] || {};

            tx.lastInwards = inwardInfo.date || null;
            tx.lastInwardQty = inwardInfo.qty || 0;
            // last outwards info
            const outwardInfo = lastOutwardMap[tx._id.toString()] || {};

            tx.lastUsed = outwardInfo.date || tx.lastUsed || null;
            tx.lastOutwardQty = outwardInfo.qty || 0;

            // 🔹 Unit
            tx.unit = unitMap[tx._id.toString()] || "";

            if (!start) {
                const firstOutward = await Transaction.findOne({
                        itemId: tx._id,
                        type: "outward",
                    })
                    .sort({ createdAt: 1 })
                    .lean();

                if (firstOutward) start = new Date(firstOutward.createdAt);
            }

            if (!start) {
                tx.avgPerDay = "0.00";
                tx.stockDaysLeft = "∞";
                tx.reorderQty = 0;
                tx.reorderStatus = "safe";
                enrichedTransactions.push(tx);
                continue;
            }
            const msPerDay = 1000 * 60 * 60 * 24;

            const diffDays = Math.max(
                1,
                Math.ceil((dateRangeEnd.getTime() - start.getTime()) / msPerDay)
            );

            // 🔹 Convert to months (approx)
            const diffMonths = Math.max(1, diffDays / 30);

            // 🔹 Monthly average consumption
            const monthlyAvg = tx.totalUsed / diffMonths;
            // 🔹 Expected Stock-Out Date
            if (monthlyAvg > 0 && tx.currentStock > 0) {
                const monthsLeftExact = tx.currentStock / monthlyAvg;
                const daysLeft = Math.floor(monthsLeftExact * 30);

                const stockOutDate = new Date();
                stockOutDate.setDate(stockOutDate.getDate() + daysLeft);

                tx.stockOutDate = stockOutDate;
            } else {
                tx.stockOutDate = null;
            }

            tx.avgPerMonth = monthlyAvg.toFixed(2);
            // 🔹 Initial stock info
            const init = initialMap[tx._id.toString()] || {};
            tx.initialDate = init.date || null;
            tx.initialQty = init.qty || 0;

            // 🔹 Stock life in months
            if (tx.currentStock === 0) {
                tx.stockMonthsLeft = 0;
            } else if (monthlyAvg > 0) {
                tx.stockMonthsLeft = Math.floor(tx.currentStock / monthlyAvg);
            } else {
                // No consumption but stock exists
                tx.stockMonthsLeft = "∞";
            }

            // 🔹 Reorder quantity = 1 month requirement
            tx.reorderQty = Math.ceil(monthlyAvg);

            if (tx.stockMonthsLeft === 0) {
                tx.reorderStatus = "danger";
            } else if (tx.stockMonthsLeft === 1) {
                tx.reorderStatus = "warning";
            } else {
                tx.reorderStatus = "safe";
            }

            enrichedTransactions.push(tx);
        }
        // ================================
        // 🔹 SUMMARY (DATE-AWARE & MODE-AWARE)
        // ================================
        let totalUsedOverall = 0;
        let summaryStartDates = [];

        enrichedTransactions.forEach((tx) => {
            totalUsedOverall += Number(tx.totalUsed || 0);

            if (tx._effectiveStart instanceof Date && !isNaN(tx._effectiveStart)) {
                summaryStartDates.push(tx._effectiveStart);
            }
        });

        // 🔹 Determine summary start
        let summaryStart = null;

        // Priority 1️⃣ Manual startDate
        if (startDate) {
            summaryStart = new Date(startDate);
        }
        // Priority 2️⃣ Since Last Inward → earliest per-item inward
        else if (mode === "sinceLastInward" && summaryStartDates.length > 0) {
            summaryStart = new Date(
                Math.min(...summaryStartDates.map((d) => d.getTime()))
            );
        }
        // Priority 3️⃣ Since Start → earliest outward
        else if (summaryStartDates.length > 0) {
            summaryStart = new Date(
                Math.min(...summaryStartDates.map((d) => d.getTime()))
            );
        }

        // 🔹 Summary end date
        const summaryEnd = dateRangeEnd || new Date();

        // 🔹 Consumption days
        let consumptionDays = 0;
        if (summaryStart && summaryEnd) {
            consumptionDays = Math.max(
                1,
                Math.ceil((summaryEnd - summaryStart) / (1000 * 60 * 60 * 24))
            );
        }

        // 🔹 Avg Life / Heat
        const avgPerHeat =
            totalHeats > 0 ? (totalUsedOverall / totalHeats).toFixed(2) : 0;
        res.render("items/consumption", {
            items,
            categories,
            suppliers,
            transactions: enrichedTransactions,
            query: req.query,

            avgPerHeat,
            openHeats,
            closeHeats,
            totalHeats,

            periodStart: summaryStart, // ✅ FIXED
            periodEnd: summaryEnd, // ✅ FIXED
            consumptionDays, // ✅ FIXED
        });

        // // ================================
        // // 🔹 Avg Life / Heat (SAFE)
        // // ================================
        // let totalUsedOverall = 0;

        // enrichedTransactions.forEach((tx) => {
        //   totalUsedOverall += Number(tx.totalUsed || 0);
        // });
        // console.log("Last inward map size:", Object.keys(lastInwardMap).length);

        // const avgPerHeat =
        //   totalHeats > 0 ? (totalUsedOverall / totalHeats).toFixed(2) : 0;
        // // 🔹 Consumption Period Summary
        // const periodStart = dateRangeStart;
        // const periodEnd = dateRangeEnd;

        // let consumptionDays = 0;
        // if (periodStart && periodEnd) {
        //   consumptionDays = Math.max(
        //     1,
        //     Math.ceil(
        //       (periodEnd.getTime() - periodStart.getTime()) / (1000 * 60 * 60 * 24)
        //     )
        //   );
        // }

        // res.render("items/consumption", {
        //   items,
        //   categories,
        //   suppliers,
        //   transactions: enrichedTransactions,
        //   query: req.query,

        //   avgPerHeat,
        //   openHeats,
        //   closeHeats,
        //   totalHeats,
        //   periodStart,
        //   periodEnd,
        //   consumptionDays,
        // });
    })
);

// *** THIS IS THE ROUTE TO FIX THE ERROR ***
// It correctly defines the DELETE method for your transactions
router.delete(
    "/transactions/:id",
    catchAsync(async(req, res) => {
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
    catchAsync(async(req, res) => {
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
    catchAsync(async(req, res) => {
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
    catchAsync(async(req, res, next) => {
        // if (!req.body.item) throw new ExpressError("Invalid Item Data", 400);

        let category = new ItemCategories(req.body.category);
        await category.save();
        res.redirect("/items/new");
    })
);

router.delete(
    "/category/:id",
    catchAsync(async(req, res) => {
        const { id } = req.params;
        await ItemCategories.findByIdAndDelete(id);
        res.redirect("/items/category");
    })
);
router.get(
    "/:id/edit",
    catchAsync(async(req, res, next) => {
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
    catchAsync(async(req, res) => {
        const { id } = req.params;
        const item = await Items.findById(id).populate("itemImage");

        // ✅ Update text fields
        Object.assign(item, req.body.item);

        // ✅ DELETE selected old images
        if (req.body.deleteImages) {
            const idsToDelete = Array.isArray(req.body.deleteImages) ?
                req.body.deleteImages :
                [req.body.deleteImages];

            // Remove from DB
            await Images.deleteMany({ _id: { $in: idsToDelete } });

            // Remove references from item
            item.itemImage = item.itemImage.filter(
                (img) => !idsToDelete.includes(img._id.toString())
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
                        path.join(__dirname, "..", "views", "images", file.filename)
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
    })
);

// Replace your existing POST "/update-stock" route with this

router.post(
    "/update-stock",
    catchAsync(async(req, res) => {
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
    })
);

router.delete(
    "/:id",
    catchAsync(async(req, res) => {
        const { id } = req.params;
        await Items.findByIdAndDelete(id);
        res.redirect("/items");
    })
);

// Add this route below other routes in items.js

router.get("/suggestions", async(req, res) => {
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
    catchAsync(async(req, res) => {
        var category = req.query.category;
        var supplier = req.query.supplier;

        var match = {};
        if (category && category !== "all") match.itemCategoryName = category;
        if (supplier && supplier !== "all") match.itemSupplier = supplier;

        const items = await Items.find(match).sort({ itemName: 1 }).lean();

        // render HTML first
        const html = await ejs.renderFile(
            path.join(__dirname, "../views/items/pdf_stock.ejs"), { items, category, supplier }, { async: false }
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
    catchAsync(async(req, res) => {
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
                $match: Object.assign({},
                    category && category !== "all" ?
                    { "item.itemCategoryName": category } :
                    {},
                    supplier && supplier !== "all" ?
                    { "item.itemSupplier": supplier } :
                    {}
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
            path.join(__dirname, "../views/items/pdf_consumption.ejs"), {
                transactions,
                category,
                supplier,
                mode,
                startDate,
                endDate,
            }, { async: false }
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

router.post("/utility/backup-drive", (req, res) => {
    runBackup("manual-drive");
    req.flash("success", "☁️ Google Drive backup started!");
    res.redirect("/items/utility");
});
router.post("/category/ajax", upload.none(), async(req, res) => {
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