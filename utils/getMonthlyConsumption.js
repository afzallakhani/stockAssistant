const Transaction = require("../models/transaction");

module.exports = async function getMonthlyConsumption(
  item,
  endDate = new Date(),
) {
  // 🔹 Find FIRST outward/lend transaction date
  const firstTx = await Transaction.findOne({
    itemId: item._id,
    type: { $in: ["outward", "lend"] },
  })
    .sort({ createdAt: 1 })
    .lean();

  // If no usage yet
  if (!firstTx) {
    return {
      totalUsed: 0,
      months: 1,
      perMonth: 0,
    };
  }

  const startDate = firstTx.createdAt;

  // 🔹 Get all usage transactions
  const txs = await Transaction.find({
    itemId: item._id,
    type: { $in: ["outward", "lend"] },
    createdAt: { $gte: startDate, $lte: endDate },
  });

  const totalUsed = txs.reduce((sum, t) => sum + t.quantity, 0);

  // 🔹 Calculate months difference properly
  const diffMs = endDate - startDate;
  const diffDays = diffMs / (1000 * 60 * 60 * 24);

  const months = diffDays / 30;

  const safeMonths = Math.max(1, Number(months.toFixed(2)));

  return {
    totalUsed,
    months: safeMonths,
    perMonth: Number((totalUsed / safeMonths).toFixed(2)),
  };
};
