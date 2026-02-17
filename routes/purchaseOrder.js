const express = require("express");
const router = express.Router();

const PurchaseOrder = require("../models/purchaseOrder");
const Supplier = require("../models/supplier");
const Items = require("../models/elafStock");
const Transaction = require("../models/transaction");
const getMonthlyConsumption = require("../utils/getMonthlyConsumption");

const catchAsync = require("../utils/catchAsync");
const generatePONumber = require("../utils/generatePONumber");

const puppeteer = require("puppeteer");
const ejs = require("ejs");
const path = require("path");
/* ===============================
   LIST ALL POs
================================ */
router.get(
  "/",
  catchAsync(async (req, res) => {
    const pos = await PurchaseOrder.find({}).sort({ createdAt: -1 });

    res.render("purchaseOrders/index", { pos });
  }),
);

/* ===============================
   SUPPLIER ITEMS API  (🚨 MUST BE FIRST)
================================ */

router.get(
  "/supplier/:supplierId/items",
  catchAsync(async (req, res) => {
    const { supplierId } = req.params;

    const supplier = await Supplier.findById(supplierId);
    if (!supplier) return res.json([]);

    const supplierName = supplier.supplierName.trim();

    const items = await Items.find({
      itemSupplier: { $regex: new RegExp(supplierName, "i") },
    }).lean();

    const response = [];

    for (let i = 0; i < items.length; i++) {
      const dc = await getMonthlyConsumption(items[i]);

      response.push({
        _id: items[i]._id,
        name: items[i].itemName,
        unit: items[i].itemUnit,
        currentStock: items[i].itemQty || 0,
        avgMonthlyConsumption: dc.perMonth || 0,
      });
    }

    res.json(response);
  }),
);
/* ===============================
   ALL ITEMS API
================================ */
router.get(
  "/items/all",
  catchAsync(async (req, res) => {
    const items = await Items.find({});

    const since = new Date();
    since.setDate(since.getDate() - 90);

    const itemIds = items.map((i) => i._id);

    const consumption = await Transaction.aggregate([
      {
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
  catchAsync(async (req, res) => {
    const suppliers = await Supplier.find({});
    const items = await Items.find({});
    res.render("purchaseOrders/new", { suppliers, items });
  }),
);

/* ===============================
   CREATE PO (POST)
================================ */
router.post(
  "/",
  catchAsync(async (req, res) => {
    const {
      supplierMode,
      supplierId,
      supplierName,
      supplierAddress,
      supplierGst,
      items,
      terms,
    } = req.body;

    /* ================= SUPPLIER ================= */
    let supplierData;

    if (supplierMode === "custom") {
      supplierData = {
        supplierType: "custom",
        name: supplierName || "—",
        address: supplierAddress || "",
        gst: supplierGst || "",
      };
    } else {
      const supplier = await Supplier.findById(supplierId);

      if (supplier) {
        supplierData = {
          supplierType: "master",
          name: supplier.supplierName,
          supplierRef: supplier._id, // ✅ VERY IMPORTANT

          address: supplier.supplierAddress || "",
          gst: supplier.supplierGst || "",
        };
      } else {
        supplierData = {
          supplierType: "unknown",
          name: "Unknown Supplier",
          address: "",
          gst: "",
        };
      }
    }

    /* ================= ITEMS ================= */
    let poItems = [];
    let subTotal = 0;

    for (let i of items) {
      let itemData = {};

      if (i.itemType === "master") {
        const dbItem = await Items.findById(i.itemId);

        const qty = Number(i.qty) || 0;
        const rate = Number(i.rate) || 0;

        itemData = {
          itemType: "master",
          itemRef: dbItem._id,
          name: dbItem.itemName,
          description: i.description || "",
          qty: qty,
          unit: i.unit || dbItem.itemUnit,
          rate: rate,
          amount: qty * rate,
        };
      } else {
        const qty = Number(i.qty) || 0;
        const rate = Number(i.rate) || 0;

        itemData = {
          itemType: "custom",
          name: i.name,
          description: i.description || "",
          qty: qty,
          unit: i.unit || "",
          rate: rate,
          amount: qty * rate,
        };
      }

      subTotal += itemData.amount;
      poItems.push(itemData);
    }

    /* ================= SAVE PO ================= */
    const po = new PurchaseOrder({
      poNumber: generatePONumber(),
      supplier: supplierData,
      items: poItems,
      terms: terms,
      subTotal: subTotal,
    });

    await po.save();

    res.redirect(`/purchase-orders/${po._id}`);
  }),
);
/* ===============================
   DELETE PO
================================ */
router.delete(
  "/:id",
  catchAsync(async (req, res) => {
    const po = await PurchaseOrder.findById(req.params.id);

    if (!po) return res.redirect("/purchase-orders");

    if (po.status === "final") {
      req.flash("error", "Finalized PO cannot be deleted.");
      return res.redirect("/purchase-orders");
    }

    await PurchaseOrder.findByIdAndDelete(req.params.id);
    res.redirect("/purchase-orders");
  }),
);

/* ===============================
   SHOW PO
================================ */
router.get(
  "/:id",
  catchAsync(async (req, res) => {
    const po = await PurchaseOrder.findById(req.params.id);

    res.render("purchaseOrders/show", { po });
  }),
);
/* ===============================
   EDIT PO (FORM)
================================ */
router.get(
  "/:id/edit",
  catchAsync(async (req, res) => {
    const po = await PurchaseOrder.findById(req.params.id);
    if (!po) return res.redirect("/purchase-orders");

    if (po.status === "final") {
      req.flash("error", "Finalized PO cannot be edited.");
      return res.redirect(`/purchase-orders/${po._id}`);
    }

    const suppliers = await Supplier.find({});
    const items = await Items.find({});

    res.render("purchaseOrders/edit", {
      po,
      suppliers,
      items,
    });
  }),
);

/* ===============================
   PDF EXPORT
================================ */
router.get(
  "/:id/pdf",
  catchAsync(async (req, res) => {
    const po = await PurchaseOrder.findById(req.params.id);

    const browser = await puppeteer.launch();
    const page = await browser.newPage();

    const html = await ejs.renderFile(
      path.join(__dirname, "../views/purchaseOrders/pdf.ejs"),
      { po },
    );

    await page.setContent(html, { waitUntil: "networkidle2" });
    await page.evaluate(async () => {
      const images = Array.from(document.images);
      await Promise.all(
        images.map((img) => {
          if (img.complete) return;
          return new Promise((resolve) => {
            img.onload = img.onerror = resolve;
          });
        }),
      );
    });

    const pdf = await page.pdf({
      format: "A4",
      printBackground: true,
      displayHeaderFooter: true,

      headerTemplate: `
    <div style="font-size:10px; width:100%; text-align:center;"></div>
  `,

      footerTemplate: `
    <div style="
      font-size:10px;
      width:100%;
      text-align:right;
      padding-right:30px;
    ">
      Page <span class="pageNumber"></span> of <span class="totalPages"></span>
    </div>
  `,
      margin: {
        top: "0mm",
        bottom: "18mm",
        left: "10mm",
        right: "10mm",
      },
    });

    await browser.close();

    res.set({
      "Content-Type": "application/pdf",
      "Content-Disposition": `attachment; filename=${po.poNumber}.pdf`,
    });

    res.send(pdf);
  }),
);

router.post(
  "/:id/finalize",
  catchAsync(async (req, res) => {
    await PurchaseOrder.findByIdAndUpdate(req.params.id, {
      status: "final",
    });
    res.redirect(`/purchase-orders/${req.params.id}`);
  }),
);

/* ===============================
   UPDATE PO
================================ */
router.put(
  "/:id",
  catchAsync(async (req, res) => {
    const po = await PurchaseOrder.findById(req.params.id);
    if (!po) return res.redirect("/purchase-orders");

    if (po.status === "final") {
      req.flash("error", "Finalized PO cannot be edited.");
      return res.redirect(`/purchase-orders/${po._id}`);
    }

    const {
      supplierMode,
      supplierId,
      supplierName,
      supplierAddress,
      supplierGst,
      items,
      terms,
    } = req.body;

    /* -------- SUPPLIER SNAPSHOT -------- */
    let supplierData;

    if (supplierMode === "custom") {
      supplierData = {
        supplierType: "custom",
        name: supplierName || "—",
        address: supplierAddress || "",
        gst: supplierGst || "",
      };
    } else {
      const supplier = await Supplier.findById(supplierId);
      supplierData = supplier
        ? {
            supplierType: "master",
            name: supplier.supplierName,
            supplierRef: supplier._id, // ✅ VERY IMPORTANT

            address: supplier.supplierAddress || "",
            gst: supplier.supplierGst || "",
          }
        : po.supplier;
    }

    /* -------- ITEMS SNAPSHOT -------- */
    let poItems = [];
    let subTotal = 0;

    for (let i of items) {
      const qty = Number(i.qty) || 0;
      const rate = Number(i.rate) || 0;

      let itemData;

      if (i.itemType === "master") {
        const dbItem = await Items.findById(i.itemId);
        itemData = {
          itemType: "master",
          itemRef: dbItem._id,
          name: dbItem.itemName,
          description: i.description || "",
          qty,
          unit: i.unit || dbItem.itemUnit,
          rate,
          amount: qty * rate,
        };
      } else {
        itemData = {
          itemType: "custom",
          name: i.name,
          description: i.description || "",
          qty,
          unit: i.unit || "",
          rate,
          amount: qty * rate,
        };
      }

      subTotal += itemData.amount;
      poItems.push(itemData);
    }

    po.supplier = supplierData;
    po.items = poItems;
    po.terms = terms;
    po.subTotal = subTotal;

    await po.save();

    res.redirect(`/purchase-orders/${po._id}`);
  }),
);

module.exports = router;
