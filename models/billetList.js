const mongoose = require("mongoose");
const Schema = mongoose.Schema;
const BilletTc = require("./billetTc");

const BilletSchema = new mongoose.Schema({
  sectionSize: String,
  gradeName: {
    type: String,
    uppercase: true,
  },
  createdAt: {
    type: Date,
    default: Date.now, // Corrected: use Date.now without parentheses
  },
  heatNo: {
    type: String,
    uppercase: true,
    unique: true,
    ref: "BilletTc",
  },
  c: String,
  mn: String,
  p: String,
  s: String,
  si: String,
  cr: String,
  mo: String,
  ni: String,
  al: String,
  cu: String,
  v: String,
  ce: String,
  // fullLengthQty: String,
  // fullPisLength: [String],
  // shortLengthQty: [String],
  // shortPisLength: [String],
});

BilletSchema.virtual("formattedHeatDate").get(function () {
  return this.createdAt.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
  });
});

module.exports = mongoose.model("BilletList", BilletSchema);
