// Add this in a temporary script, e.g. updateSuppliers.js

const mongoose = require("mongoose");
const Items = require("./models/elafStock"); // adjust path if needed

mongoose.connect("mongodb://localhost:27017/stockAssistant", {
  useNewUrlParser: true,
  useUnifiedTopology: true,
});

(async () => {
  try {
    const result = await Items.updateMany(
      {},
      { $set: { itemSupplier: "RHI MAGNESITA INDIA LIMITED" } }
    );

    console.log(`✅ Updated ${result.modifiedCount} items with itemSupplier`);
  } catch (err) {
    console.error("❌ Error updating suppliers:", err);
  } finally {
    mongoose.connection.close();
  }
})();
