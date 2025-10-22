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

router.get(
    "/",
    catchAsync(async(req, res) => {
        const items = await Items.find({}).populate("itemImage");
        const images = await Images.find({});
        res.render("items/allItems", { items });
    })
);
router.get("/", (req, res) => {
    res.render("home");
});

router.get(
    "/new",
    catchAsync(async(req, res) => {
        const itemCategories = await ItemCategories.find({});
        res.render("items/new", { itemCategories });
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
        const items = await Items.find({}); // Fetch all items from the database
        res.render("items/outwards", { items }); // Render the form and pass the items
    } catch (error) {
        console.error("Error fetching items for outwards log:", error);
        res.status(500).send("Error fetching items.");
    }
});

router.get(
    "/category",
    catchAsync(async(req, res) => {
        const category = await ItemCategories.find({});
        res.render("items/category", { category });
    })
);

// searched items display
router.get("/search", async(req, res) => {
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
    catchAsync(async(req, res, next) => {
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
        const initialStock = item.itemQty;
        await new Transaction({
            itemId: item._id,
            type: "initial",
            quantity: initialStock,
            stockBefore: 0, // Stock is 0 before creation
            stockAfter: initialStock,
        }).save();
        res.redirect("/items");
    })
);
// GET route to display the inwards form
router.get(
    "/inwards",
    catchAsync(async(req, res) => {
        const items = await Items.find({}); // Fetch all items
        res.render("items/inwards", { items });
    })
);

// POST route to handle inward stock updates
router.post(
    "/inward-stock",
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
    catchAsync(async(req, res) => {
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
            filter.createdAt = {...filter.createdAt, $lt: endDate };
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
                item.itemQty -= transaction.quantity;
            } else if (transaction.type === "outward") {
                item.itemQty += transaction.quantity;
            }
            await item.save();
        }

        await Transaction.findByIdAndDelete(id);
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
        res.render("items/edit", { item, itemCategories });
        // next(e);
    })
);
router.put(
    "/:id",
    upload.single("item[itemImage]"),
    catchAsync(async(req, res, next) => {
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
            await Items.findByIdAndUpdate(id, {...req.body.item });
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
            await Items.findByIdAndUpdate(id, {...req.body.item });
        }
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
    catchAsync(async(req, res) => {
        const { id } = req.params;
        await Items.findByIdAndDelete(id);
        res.redirect("/items");
    })
);
module.exports = router;