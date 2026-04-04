const mongoose = require("mongoose");
const Items = require("./models/elafStock");
const Supplier = require("./models/supplier");

async function migrateSuppliers() {
  await mongoose.connect("mongodb://127.0.0.1:27017/stockAssistant");

  const items = await Items.find({});

  for (let item of items) {
    if (typeof item.itemSupplier === "string") {
      let name = item.itemSupplier.trim().toUpperCase();

      // 🔥 REMOVE CITY IN BRACKETS
      name = name.replace(/\(.*?\)/g, "").trim();

      let supplier = await Supplier.findOne({ supplierName: name });

      if (supplier) {
        item.itemSupplier = supplier._id;
        await item.save();
        console.log(`Updated ${item.itemName}`);
      } else {
        console.log(`Supplier NOT FOUND: ${name}`);
      }
    }
  }

  console.log("Migration completed");
  process.exit();
}

migrateSuppliers();
