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

    var baseMatch = {
      type: { $in: ["outward", "lend"] }, // consumption types
    };

    // ================================
    // 🔹 DATE RANGE LOGIC (FINAL)
    // =============================

    // END date → today if not selected
    let dateRangeEnd = endDate
      ? new Date(new Date(endDate).setHours(23, 59, 59, 999))
      : new Date();

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
    const allItems = await Items.find({})
      .select("_id itemQty itemName itemUnit")
      .lean();
    allItems.forEach(function (itm) {
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
    const openHeats = billets.filter(function (b) {
      return !b.ce || b.ce === "";
    }).length;
    const closeHeats = billets.filter(function (b) {
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
      const tx = { ...rawTx };

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
