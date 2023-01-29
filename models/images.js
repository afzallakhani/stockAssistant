const mongoose = require("mongoose");
const Schema = mongoose.Schema;

const imageSchema = new mongoose.Schema({
    data: Buffer,
    contentType: String,
    path: String,
    name: String,
});

const Images = mongoose.model("Images", imageSchema);
module.exports = Images;