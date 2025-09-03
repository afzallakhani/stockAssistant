const mongoose = require("mongoose");
const Schema = mongoose.Schema;

const transactionSchema = new Schema({
    itemId: {
        type: Schema.Types.ObjectId,
        ref: "allItems", // Reference to your Item model
        required: true,
    },
    type: {
        type: String,
        enum: ["initial", "inward", "outward", "adjustment"], // Transaction types
        required: true,
    },
    quantity: {
        type: Number,
        required: true,
    },
    stockBefore: {
        type: Number,
        required: true,
    },
    stockAfter: {
        type: Number,
        required: true,
    },
    createdAt: {
        type: Date,
        default: Date.now,
    },
});

module.exports = mongoose.model("Transaction", transactionSchema);