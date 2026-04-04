const mongoose = require("mongoose");
const Items = require("./models/elafStock"); // adjust path if needed

mongoose.connect("mongodb://localhost:27017/stockAssistant", {
  useNewUrlParser: true,
  useUnifiedTopology: true,
});

(async () => {
  try {
    // Update items where category name is "CCM REFACTORIES"
    const result = await Items.updateMany(
      { itemCategoryName: "CCM REFACTORIES" },
      { $set: { itemCategoryName: "CCM REFRACTORIES" } }
    );

    console.log(`✅ Updated ${result.modifiedCount} items successfully`);
  } catch (err) {
    console.error("❌ Error updating category names:", err);
  } finally {
    mongoose.connection.close();
  }
})();
