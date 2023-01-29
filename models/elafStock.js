const mongoose = require("mongoose");
const Schema = mongoose.Schema;
const Images = require("./images");

const ItemSchema = new mongoose.Schema({
    itemName: String,
    itemUnit: String,
    itemQty: Number,
    itemDescription: String,
    itemImage: [{
        type: Schema.Types.ObjectId,
        ref: "Images",
    }, ],
    // itemImage: String,
    itemCategoryName: String,
});

ItemSchema.post("findOneAndDelete", async function(item) {
    if (item) {
        await Images.deleteMany({
            _id: {
                $in: item.itemImage,
            },
        });
    }
});

// const Items = mongoose.model("allItems", ItemSchema);

module.exports = mongoose.model("allItems", ItemSchema);