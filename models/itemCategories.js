const mongoose = require("mongoose");
// const Schema = mongoose.Schema;
const CategorySchema = new mongoose.Schema({
  // itemImage: String,
  itemCategoryName: String,
});

// module.exports = mongoose.model("itemCategories", CategorySchema);
const ItemCategories = mongoose.model("ItemCategories", CategorySchema);
module.exports = ItemCategories;
