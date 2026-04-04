const mongoose = require("mongoose");
const Schema = mongoose.Schema;

// const transactionSchema = new Schema({
//     itemId: {
//         type: Schema.Types.ObjectId,
//         ref: "allItems", // Reference to your Item model
//         required: true,
//     },
//     type: {
//         type: String,
//         enum: ["initial", "inward", "outward", "adjustment"], // Transaction types
//         required: true,
//     },
//     quantity: {
//         type: Number,
//         required: true,
//     },
//     stockBefore: {
//         type: Number,
//         required: true,
//     },
//     stockAfter: {
//         type: Number,
//         required: true,
//     },
//     createdAt: {
//         type: Date,
//         default: Date.now,
//     },
// });

// const transactionSchema = new Schema({
//   itemId: {
//     type: Schema.Types.ObjectId,
//     ref: "allItems",
//     required: true,
//   },
//   type: {
//     type: String,
//     enum: ["initial", "inward", "outward", "lend", "return", "adjustment"], // ✅ Added lend & return
//     required: true,
//   },
//   quantity: {
//     type: Number,
//     required: true,
//   },
//   stockBefore: Number,
//   stockAfter: Number,
//   borrower: String,

//   returned: {
//     type: Boolean,
//     default: false, // ✅ new field
//   },
//   createdAt: {
//     type: Date,
//     default: Date.now,
//   },
// });

const transactionSchema = new Schema({
  itemId: {
    type: Schema.Types.ObjectId,
    ref: "allItems",
    required: true,
  },
  type: {
    type: String,
    enum: ["initial", "inward", "outward", "lend", "return", "adjustment"],
    required: true,
  },
  quantity: {
    type: Number,
    required: true,
  },
  stockBefore: Number,
  stockAfter: Number,
  borrower: String,
  returned: {
    type: Boolean,
    default: false,
  },
  remarks: {
    type: String,
    trim: true,
    default: "",
  },
  createdAt: {
    type: Date,
    default: Date.now,
  },
});
// 🔥 Indexes for fast filtering
transactionSchema.index({ itemId: 1 });
transactionSchema.index({ type: 1 });
transactionSchema.index({ createdAt: -1 });

// ✅ Best compound index for your filter page
transactionSchema.index({ itemId: 1, type: 1, createdAt: -1 });

module.exports = mongoose.model("Transaction", transactionSchema);
