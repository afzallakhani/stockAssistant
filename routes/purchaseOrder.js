const express = require("express");
const router = express.Router();

const PurchaseOrder = require("../models/purchaseOrder");
const Supplier = require("../models/supplier");
const Items = require("../models/elafStock");
const Transaction = require("../models/transaction");

const catchAsync = require("../utils/catchAsync");
const generatePONumber = require("../utils/generatePONumber");

const puppeteer = require("puppeteer");
const ejs = require("ejs");
const path = require("path");

/* ===============================
   SUPPLIER ITEMS API  (🚨 MUST BE FIRST)
================================ */
router.get(
    "/supplier/:supplierId/items",
    catchAsync(async(req, res) => {
        const { supplierId } = req.params;

        // 1️⃣ Get supplier
        const supplier = await Supplier.findById(supplierId);
        if (!supplier) {
            return res.json([]);
        }

        // 2️⃣ Find items USING SUPPLIER NAME (IMPORTANT)
        const supplierName = supplier.supplierName.trim();

        const items = await Items.find({
            itemSupplier: {
                $regex: new RegExp(supplierName, "i"),
            },
        });

        // 3️⃣ Avg monthly consumption (last 90 days)
        const since = new Date();
        since.setDate(since.getDate() - 90);

        const itemIds = items.map((i) => i._id);

        const consumption = await Transaction.aggregate([{
                $match: {
                    item: { $in: itemIds },
                    type: "OUTWARD",
                    createdAt: { $gte: since },
                },
            },
            {
                $group: {
                    _id: "$item",
                    totalUsed: { $sum: "$quantity" },
                },
            },
        ]);

        const consumptionMap = {};
        consumption.forEach((c) => {
            consumptionMap[c._id.toString()] = Math.round(c.totalUsed / 3);
        });

        // 4️⃣ Response
        res.json(
            items.map((item) => ({
                _id: item._id,
                name: item.itemName,
                unit: item.itemUnit,
                currentStock: item.itemQty || 0,
                avgMonthlyConsumption: consumptionMap[item._id.toString()] || 0,
            })),
        );
    }),
);

/* ===============================
   CREATE PO FORM
================================ */
router.get(
    "/new",
    catchAsync(async(req, res) => {
        const suppliers = await Supplier.find({});
        res.render("purchaseOrders/new", { suppliers });
    }),
);

/* ===============================
   CREATE PO (POST)
================================ */
router.post(
    "/",
    catchAsync(async(req, res) => {
        const { supplier, items } = req.body;

        const cleanedItems = items
            .filter((i) => i.quantity && Number(i.quantity) > 0)
            .map((i) => ({
                item: i.item,
                quantity: Number(i.quantity),
                unit: i.unit,
            }));

        if (!cleanedItems.length) {
            req.flash("error", "Please enter quantity for at least one item");
            return res.redirect("/purchase-orders/new");
        }

        const po = new PurchaseOrder({
            poNumber: generatePONumber(),
            supplier,
            items: cleanedItems,
        });

        await po.save();
        res.redirect(`/purchase-orders/${po._id}`);
    }),
);

/* ===============================
   SHOW PO
================================ */
router.get(
    "/:id",
    catchAsync(async(req, res) => {
        const po = await PurchaseOrder.findById(req.params.id)
            .populate("supplier")
            .populate("items.item");

        res.render("purchaseOrders/show", { po });
    }),
);

/* ===============================
   PDF EXPORT
================================ */
router.get(
    "/:id/pdf",
    catchAsync(async(req, res) => {
        const po = await PurchaseOrder.findById(req.params.id)
            .populate("supplier")
            .populate("items.item");

        const browser = await puppeteer.launch();
        const page = await browser.newPage();

        const html = await ejs.renderFile(
            path.join(__dirname, "../views/purchaseOrders/pdf.ejs"), { po },
        );

        await page.setContent(html, { waitUntil: "networkidle0" });

        const pdf = await page.pdf({
            format: "A4",
            printBackground: true,
        });

        await browser.close();

        res.set({
            "Content-Type": "application/pdf",
            "Content-Disposition": `attachment; filename=${po.poNumber}.pdf`,
        });

        res.send(pdf);
    }),
);

module.exports = router;