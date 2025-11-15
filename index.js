const express = require("express");
const bodyParser = require("body-parser");
const path = require("path");
const mongoose = require("mongoose");
const runBackup = require("./utils/backupHelper");

// const Items = require("./models/elafStock");
const item = require("./routes/item");
const party = require("./routes/party");
const supplier = require("./routes/supplier");
const docx = require("docx");

const billets = require("./routes/billet");
// const Images = require("./models/images");
// const Joi = require("joi");
const ejsMate = require("ejs-mate");
const methodOverride = require("method-override");
const fs = require("fs");
require("dotenv/config");
const session = require("express-session");
const flash = require("connect-flash");

const ExpressError = require("./utils/ExpressError");

// const catchAsync = require("./utils/catchAsync");
// const { promisify } = require("util");
// const multerStorage = require("./utils/multerStorage");
// const unlinkAsync = promisify(fs.unlink);

mongoose.connect("mongodb://localhost:27017/stockAssistant", {
  useNewUrlParser: true,
  useCreateIndex: true,
  useUnifiedTopology: true,

  useFindAndModify: false,
});

const db = mongoose.connection;
db.on("error", console.error.bind(console, "connection error:"));
db.once("open", () => {
  console.log("Database Connected");
  console.log("MongoDB Version:", mongoose.version);
  const changeStream = db.watch();

  changeStream.on("change", (change) => {
    console.log("📦 Database change detected:", change.operationType);
    runBackup("auto");
  });

  console.log("👀 Auto-backup watcher started...");
});

const app = express();
console.log("Node.js Version:", process.version);
app.engine("ejs", ejsMate);
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(methodOverride("_method"));
// app.use(bodyParser.json());

// Session config
const sessionConfig = {
  secret: "supersecretbackupkey",
  resave: false,
  saveUninitialized: true,
  cookie: { httpOnly: true, maxAge: 1000 * 60 * 60 }, // 1 hour
};

app.use(session(sessionConfig));
app.use(flash());

// Make flash messages available to all views
app.use((req, res, next) => {
  res.locals.success = req.flash("success");
  res.locals.error = req.flash("error");
  next();
});

app.use("/items", item);
app.use("/partymaster", party);
app.use("/supplier", supplier);
app.use("/billets", billets);
app.use(express.static(path.join(__dirname, "public")));

app.get("/", (req, res) => {
  res.render("home");
});

// app.use((req, res) => {
//     res.status(404).send("NOT FOUND!");
// });
app.all("*", (req, res, next) => {
  next(new ExpressError("Page Not Found!", 404));
});

app.use((err, req, res, next) => {
  const { statusCode = 500, message = "Something Went Wrong!" } = err;
  if (!err.message) err.message = "Oh No! Something Went Wrong!";
  res.status(statusCode).render("error", { err });
});

app.listen(3000, "0.0.0.0", () => {
  console.log("App Running On Port 3000");
});
