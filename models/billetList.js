// const { date, String } = require("joi");
const mongoose = require("mongoose");
const Schema = mongoose.Schema;
const BilletTc = require("./billetTc");
const BilletSchema = new mongoose.Schema({
  sectionSize: String,
  gradeName: {
    type: String,
    uppercase: true,
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

module.exports = mongoose.model("BilletList", BilletSchema);
// module.exports = mongoose.SchemaTypes("billetList", BilletSchema);
// module.exports.productionDate = BilletSchema;
