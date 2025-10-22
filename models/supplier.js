const mongoose = require("mongoose");

const SupplierSchema = new mongoose.Schema({
    // itemImage: String,
    supplierName: {
        type: String,
        uppercase: true,
        unique: true,
    },
    supplierCity: {
        type: String,
        uppercase: true,
    },
    supplierState: {
        type: String,
        uppercase: true,
    },

    supplierItemCategory: {
        type: String,
        uppercase: true,
    },
    supplierEmail: {
        type: String,
        uppercase: true,
    },
    supplierPhone: {
        type: String,
        uppercase: true,
    },
    supplierGst: {
        type: String,
        uppercase: true,
    },
});

module.exports = mongoose.model("supplier", SupplierSchema);