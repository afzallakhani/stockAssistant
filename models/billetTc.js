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
    uppercase: true,
    default: function () {
      return this.billNo;
    },
  },

  tcDate: { type: Date, default: Date.now() },
  poDate: {
    type: Date,
    default: Date.now(),
  },
  poNo: {
    type: String,
    uppercase: true,
  },
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
  colorCode: {
    type: String,
    uppercase: true,
  },
  createdAt: {
    type: Date,
    default: Date.now,
  },
});
BilletTc.virtual("formattedDate").get(function () {
  return this.createdAt.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  });
});
BilletTc.virtual("formattedTcDate").get(function () {
  return this.tcDate.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
  });
});
BilletTc.virtual("formattedPoDate").get(function () {
  return this.poDate.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
  });
});
module.exports = mongoose.model("BilletTc", BilletTc);
