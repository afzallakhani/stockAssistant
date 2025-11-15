const { required } = require("joi");
const mongoose = require("mongoose");

const SupplierSchema = new mongoose.Schema({
  // itemImage: String,
  supplierName: {
    type: String,
    uppercase: true,
    unique: true,
    trim: true,
    required: true,
  },
  supplierCity: {
    type: String,
    uppercase: true,
    trim: true,
    required: true,
  },
  supplierState: {
    type: String,
    uppercase: true,
    trim: true,
    required: true,
  },

  supplierItemCategory: {
    type: String,
    uppercase: true,
    trim: true,
    required: false,
  },
  supplierEmail: {
    type: String,
    lowercase: true,
    trim: true,
  },
  supplierPhone: {
    type: String,
    uppercase: true,
    trim: true,
  },
  supplierGst: {
    type: String,
    uppercase: true,
    trim: true,
  },
});

module.exports = mongoose.model("supplier", SupplierSchema);
