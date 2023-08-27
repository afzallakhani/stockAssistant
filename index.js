const express = require("express");
const bodyParser = require("body-parser");
const path = require("path");
const mongoose = require("mongoose");
// const Items = require("./models/elafStock");
const item = require("./routes/item");
const party = require("./routes/party");
// const Images = require("./models/images");
// const Joi = require("joi");
const ejsMate = require("ejs-mate");
const methodOverride = require("method-override");
const fs = require("fs");
require("dotenv/config");
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
});

const app = express();

app.engine("ejs", ejsMate);
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(methodOverride("_method"));
// app.use(bodyParser.json());
app.use("/items", item);
app.use("/partymaster", party);
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

app.listen(3000, () => {
  console.log("App Running On Port 3000");
});
