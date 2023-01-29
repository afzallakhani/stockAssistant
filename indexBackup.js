// const express = require("express");
// const bodyParser = require("body-parser");
// const path = require("path");
// const mongoose = require("mongoose");
// const Items = require("./models/elafStock");
// const Images = require("./models/images");
// const { urlencoded } = require("express");
// const ejsMate = require("ejs-mate");
// const methodOverride = require("method-override");
// const multer = require("multer");
// const fs = require("fs");
// require("dotenv/config");
// const { promisify } = require("util");
// const unlinkAsync = promisify(fs.unlink);

// mongoose.connect("mongodb://localhost:27017/stockAssistant", {
//     useNewUrlParser: true,
//     useCreateIndex: true,
//     useUnifiedTopology: true,
//     useFindAndModify: false,
// });

// const db = mongoose.connection;
// db.on("error", console.error.bind(console, "connection error:"));
// db.once("open", () => {
//     console.log("Database Connected");
// });

// const app = express();

// app.engine("ejs", ejsMate);
// app.set("view engine", "ejs");
// app.set("views", path.join(__dirname, "views"));

// app.use(express.urlencoded({ extended: true }));
// app.use(methodOverride("_method"));
// app.use(bodyParser.json());

// var storage = multer.diskStorage({
//     destination: (req, file, cb) => {
//         cb(null, "views/images");
//     },
//     filename: (req, file, cb) => {
//         cb(null, file.fieldname + "-" + Date.now());
//     },
// });

// var upload = multer({ storage: storage });

// app.get("/", (req, res) => {
//     res.render("home");
// });

// app.get("/items", async(req, res) => {
//     const items = await Items.find({}).populate("itemImage");

//     const images = await Images.find({});
//     res.render("items/allItems", { items });
// });

// app.get("/uploads", (req, res) => {
//     res.render("items/uploads");
// });

// app.post("/uploads", upload.single("image"), (req, res) => {
//     console.log(req.file);
// });
// app.get("/items/new", (req, res) => {
//     res.render("items/new");
// });

// app.post("/items", upload.single("item[itemImage]"), async(req, res, next) => {
//     let item = new Items(req.body.item);
//     let image = new Images();
//     image.contentType = req.file.mimetype;
//     image.data = fs.readFileSync(
//         path.join(__dirname + "/views/images/" + req.file.filename)
//     );
//     image.path = req.file.path;
//     image.name = req.file.originalname;
//     const imagePath = req.file.path;
//     await item.save();
//     const currentItem = await Items.findById(item.id);
//     currentItem.itemImage.push(image);
//     await image.save();
//     await currentItem.save();
//     await fs.unlink(imagePath, (err) => {
//         if (err) {
//             console.error(err);
//             return;
//         } else {
//             return;
//         }
//     });
//     res.redirect("/items");
// });

// app.get("/items/:id/edit", async(req, res) => {
//     const item = await Items.findById(req.params.id).populate("itemImage");
//     res.render("items/edit", { item });
// });
// app.put(
//     "/items/:id",
//     upload.single("item[itemImage]"),
//     async(req, res, next) => {
//         const { id } = req.params;
//         const item = await Items.findById(id).populate("itemImage");
//         const imageId = item.itemImage[0].id;
//         let image = await Images.findById(imageId);
//         if (req.file) {
//             image.contentType = req.file.mimetype;
//             image.data = fs.readFileSync(
//                 path.join(__dirname + "/views/images/" + req.file.filename)
//             );
//             image.path = req.file.path;
//             image.name = req.file.originalname;
//             await Items.findByIdAndUpdate(id, {...req.body.item });
//             await Images.findByIdAndUpdate(imageId, image);
//             const imagePath = req.file.path;
//             await fs.unlink(imagePath, (err) => {
//                 if (err) {
//                     console.error(err);
//                     return;
//                 } else {
//                     return;
//                 }
//             });
//         } else {
//             await Items.findByIdAndUpdate(id, {...req.body.item });
//         }
//         res.redirect("/items");
//     }
// );

// app.delete("/items/:id", upload.none(), async(req, res, next) => {
//     const { id } = req.params;
//     const item = await Items.findById(id).populate("itemImage");
//     const imageId = item.itemImage[0].id;
//     if (imageId) {
//         await Images.findByIdAndDelete(imageId);
//     }
//     await Items.findByIdAndDelete(id);
//     res.redirect("/items");
// });

// app.listen(3000, () => {
//     console.log("App Running On Port 3000");
// });