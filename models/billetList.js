// const { date, Number } = require("joi");
const mongoose = require("mongoose");
const Schema = mongoose.Schema;

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
  },
  productionDate: String,
  c: Number,
  mn: Number,
  p: Number,
  s: Number,
  si: Number,
  cr: Number,
  mo: Number,
  ni: Number,
  al: Number,
  cu: Number,
  v: Number,
  ce: Number,
  fullLengthQty: Number,
  fullPisLength: [Number],
  shortLengthQty: [Number],
  shortPisLength: [Number],
});

module.exports = mongoose.model("billetList", BilletSchema);
// module.exports = mongoose.SchemaTypes("billetList", BilletSchema);
// module.exports.productionDate = BilletSchema;
