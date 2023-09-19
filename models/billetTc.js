const mongoose = require("mongoose");
const Schema = mongoose.Schema;

const BilletList = require("./billetList");
const { func } = require("joi");

const BilletTc = new mongoose.Schema({
  billNo: {
    type: String,
    uppercase: true,
  },
  tcNo: {
    type: String,
    default: function () {
      return this.billNo;
    },
  },

  tcDate: String,
  poDate: String,
  poNo: String,
  totalQtyMts: Number,
  totalPcs: Number,
  vehicleNo: {
    type: String,
    uppercase: true,
  },
  buyerName: {
    type: String,
    uppercase: true,
  },
  heatNo: [
    {
      pcs: String,
      type: mongoose.Schema.Types.ObjectId,
      ref: "BilletList",
    },
  ],
  colorCode: [String],
});

module.exports = mongoose.model("BilletTc", BilletTc);
