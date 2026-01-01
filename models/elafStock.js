const mongoose = require("mongoose");
const Schema = mongoose.Schema;
const Images = require("./images");
const Transaction = require("./transaction"); // Import the Transaction model

const ItemSchema = new mongoose.Schema({
  itemName: {
    type: String,
    uppercase: true,
  },
  itemSupplier: {
    type: String,
    uppercase: true,
  },
  itemUnit: {
    type: String,
    uppercase: true,
  },
  itemQty: Number,
  itemDescription: {
    type: String,
    uppercase: true,
  },
  life: {
    // Add the new 'life' field
    type: String,
    uppercase: true,
    default: 0,
  },
  itemImage: [
    {
      type: Schema.Types.ObjectId,
      ref: "Images",
    },
  ],
  // itemImage: String,
  itemCategoryName: String,
  createdAt: {
    type: Date,
    default: Date.now, // Corrected: use Date.now without parentheses
  },
});

ItemSchema.post("findOneAndDelete", async function (item) {
  if (item) {
    await Images.deleteMany({
      _id: {
        $in: item.itemImage,
      },
    });
    // Add this line to delete associated transactions
    await Transaction.deleteMany({ itemId: item._id });
  }
});

ItemSchema.virtual("formattedItemDate").get(function () {
  return this.createdAt.toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
  });
});
// 🔥 Indexes for item filtering & reports
ItemSchema.index({ itemCategoryName: 1 });
ItemSchema.index({ itemSupplier: 1 });
ItemSchema.index({ createdAt: -1 });

// Optional (helps sorting & dropdowns)
ItemSchema.index({ itemName: 1 });

// const Items = mongoose.model("allItems", ItemSchema);

module.exports = mongoose.model("allItems", ItemSchema);
