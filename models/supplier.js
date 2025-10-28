const mongoose = require("mongoose");

const SupplierSchema = new mongoose.Schema({
  // itemImage: String,
  supplierName: {
    type: String,
    uppercase: true,
    unique: true,
    trim: true,
  },
  supplierCity: {
    type: String,
    uppercase: true,
    trim: true,
  },
  supplierState: {
    type: String,
    uppercase: true,
    trim: true,
  },

  supplierItemCategory: {
    type: String,
    uppercase: true,
    trim: true,
  },
  supplierEmail: {
    type: String,
    lowercase: true,
    unique: true,
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
    unique: true,
    trim: true,
  },
});

module.exports = mongoose.model("supplier", SupplierSchema);
