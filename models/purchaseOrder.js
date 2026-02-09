const mongoose = require("mongoose");
const Schema = mongoose.Schema;

const PurchaseOrderSchema = new Schema({
  poNumber: {
    type: String,
    required: true,
    unique: true,
  },

  supplier: {
    type: Schema.Types.ObjectId,
    ref: "supplier",
    required: true,
  },

  items: [
    {
      item: {
        type: Schema.Types.ObjectId,
        ref: "allItems",
        required: true,
      },
      quantity: {
        type: Number,
        required: true,
      },
      unit: String,
      remark: String,
    },
  ],

  status: {
    type: String,
    enum: ["CREATED", "SENT", "RECEIVED", "CANCELLED"],
    default: "CREATED",
  },

  createdAt: {
    type: Date,
    default: Date.now,
  },
});

module.exports = mongoose.model("PurchaseOrder", PurchaseOrderSchema);
