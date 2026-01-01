const Transaction = require("../models/transaction");

module.exports = async function getDailyConsumption(
  item,
  endDate = new Date()
) {
  const startDate = item.createdAt;

  const txs = await Transaction.find({
    itemId: item._id,
    type: { $in: ["outward", "lend"] },
    createdAt: { $gte: startDate, $lte: endDate },
  });

  const totalUsed = txs.reduce((sum, t) => sum + t.quantity, 0);

  const days = Math.max(
    1,
    Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24))
  );

  return {
    totalUsed,
    days,
    perDay: Number((totalUsed / days).toFixed(2)),
  };
};
