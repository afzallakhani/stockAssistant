const express = require("express");
const bodyParser = require("body-parser");
const path = require("path");
const mongoose = require("mongoose");
const Items = require("./models/elafStock");
const { urlencoded } = require("express");
const ejsMate = require("ejs-mate");
const methodOverride = require("method-override");
const multer = require("multer");
const fs = require("fs");
const { find } = require("./models/elafStock");
require("dotenv/config");

mongoose.connect("mongodb://localhost:27017/stockAssistant", {
    useNewUrlParser: true,
    useCreateIndex: true,
    useUnifiedTopology: true,
});

const db = mongoose.connection;
db.on("error", console.error.bind(console, "connection error:"));
db.once("open", () => {
    console.log("Database Connected");
    console.log(path.join(__dirname + "/images"));
});

const app = express();

app.engine("ejs", ejsMate);
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));

app.use(express.urlencoded({ extended: true }));
app.use(methodOverride("_method"));
app.use(bodyParser.json());

var storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, "views/images");
    },
    filename: (req, file, cb) => {
        cb(
            null,
            file.fieldname + "-" + Date.now() + path.extname(file.originalname)
        );
    },
});

var upload = multer({ storage: storage });

app.get("/", (req, res) => {
    res.render("home");
});
app.get("/items", async(req, res) => {
    const items = await Items.find({});
    res.render("items/allItems", { items });
});

app.get("/uploads", (req, res) => {
    res.render("items/uploads");
});

app.post("/uploads", upload.single("image"), (req, res) => {
    // res.json({ image: req.file });
    console.log(req.file);
});
app.get("/items/new", (req, res) => {
    res.render("items/new");
});

// app.get('/items/:id', (req, res, next)=>{
//     res.send(req.body)
// })

app.post("/items", upload.single("item[itemImage]"), async(req, res, next) => {
    let item = new Items(req.body.item);
    // item.itemImage.contentType = req.file.mimetype;
    // item.itemImage.data = fs.readFileSync(
    //     path.join(__dirname + "/views/images" + req.file.filename)
    // );
    item.itemImage = req.file.filename;
    console.log(item);
    await item.save();
    res.redirect("/items");
});

app.get("/items/:id/edit", async(req, res) => {
    const item = await Items.findById(req.params.id);
    res.render("items/edit", { item });
});
app.put("/items/:id", async(req, res, next) => {
    const { id } = req.params;
    const item = await Items.findByIdAndUpdate(id, {
        ...req.body.item,
    });
    res.redirect("/items");
});
app.delete("/items/:id", async(req, res, next) => {
    const item = await Items.findByIdAndDelete(req.params.id);
    res.redirect("/items");
});

app.listen(3000, () => {
    console.log("App Running On Port 3000");
});