module.exports = function generatePONumber() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  const t = Date.now().toString().slice(-4);

  return `PO-${y}${m}${d}-${t}`;
};
