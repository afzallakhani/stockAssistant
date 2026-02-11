// const Transaction = require("../models/transaction");

// module.exports = async function getDailyConsumption(
//   item,
//   endDate = new Date()
// ) {
//   const startDate = item.createdAt;

//   const txs = await Transaction.find({
//     itemId: item._id,
//     type: { $in: ["outward", "lend"] },
//     createdAt: { $gte: startDate, $lte: endDate },
//   });

//   const totalUsed = txs.reduce((sum, t) => sum + t.quantity, 0);

//   const days = Math.max(
//     1,
//     Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24))
//   );

//   return {
//     totalUsed,
//     days,
//     perDay: Number((totalUsed / days).toFixed(2)),
//   };
// };
const Transaction = require("../models/transaction");

module.exports = async function getMonthlyConsumption(
  item,
  endDate = new Date(),
) {
  const startDate = item.createdAt;

  const txs = await Transaction.find({
    itemId: item._id,
    type: { $in: ["outward", "lend"] },
    createdAt: { $gte: startDate, $lte: endDate },
  });

  const totalUsed = txs.reduce((sum, t) => sum + t.quantity, 0);

  // 🔹 Calculate months difference (fractional)
  const months =
    (endDate.getFullYear() - startDate.getFullYear()) * 12 +
    (endDate.getMonth() - startDate.getMonth()) +
    Math.max(1, (endDate.getDate() - startDate.getDate()) / 30);

  const safeMonths = Math.max(1, Number(months.toFixed(2)));

  return {
    totalUsed,
    months: safeMonths,
    perMonth: Number((totalUsed / months).toFixed(2)),
  };
};
