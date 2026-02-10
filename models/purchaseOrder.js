const mongoose = require("mongoose");
const Schema = mongoose.Schema;

const POItemSchema = new Schema({
    itemType: {
        type: String,
        enum: ["master", "custom"],
        required: true,
    },

    // If selected from DB
    itemRef: {
        type: Schema.Types.ObjectId,
        ref: "allItems",
    },

    // Snapshot / Custom fields
    name: String,
    description: String,

    qty: {
        type: Number,
        required: true,
    },
    unit: String,

    rate: {
        type: Number,
        required: true,
    },

    amount: Number,
});

const SupplierSnapshotSchema = new Schema({
    supplierType: {
        type: String,
        enum: ["master", "custom"],
        required: true,
    },

    // If selected from DB
    supplierRef: {
        type: Schema.Types.ObjectId,
        ref: "supplier",
    },

    // Snapshot / Custom
    name: String,
    address: String,
    gst: String,
    phone: String,
    email: String,
});

const PurchaseOrderSchema = new Schema({
    poNumber: String,

    supplier: SupplierSnapshotSchema,

    items: [POItemSchema],

    terms: String,

    subTotal: Number,

    status: {
        type: String,
        enum: ["draft", "final"],
        default: "draft",
    },
}, {
    timestamps: true,
}, );

module.exports = mongoose.model("PurchaseOrder", PurchaseOrderSchema);