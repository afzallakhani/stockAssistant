const mongoose = require("mongoose");
const Items = require("../models/elafStock");

mongoose.connect("mongodb://localhost:27017/stockAssistant", {
    useNewUrlParser: true,
    useCreateIndex: true,
    useUnifiedTopology: true,
});

const db = mongoose.connection;
db.on("error", console.error.bind(console, "connection error:"));
db.once("open", () => {
    console.log("Database Connected");
});

const seedItems = [{
        itemName: "Tundish Nozzle",
        itemUnit: "NOS",
        itemQty: 200,
        itemDescription: "15 MM Tundish Nozzle",
        itemImage: "https://www.google.com/search?q=tundish+nozzle&rlz=1C1CHBF_enIN923IN923&sxsrf=ALeKk01r-m9wShDZFnHV2rRvZSzkiQZ6nA:1615392791152&tbm=isch&source=iu&ictx=1&fir=Pa6Cv5J76uPGiM%252Cq21k3791AOZxIM%252C_&vet=1&usg=AI4_-kR6PCutBSmguMkCyb7QlyzcPlDkJA&sa=X&ved=2ahUKEwiG3PSLj6bvAhWoyDgGHR1RCtsQ9QF6BAgbEAE#imgrc=Pa6Cv5J76uPGiM",
        itemCategoryName: "CCM REFACTORY",
    },
    {
        itemName: "Ladle Nozzle",
        itemUnit: "NOS",
        itemQty: 50,
        itemDescription: "1QC Ladle Nozzle",
        itemImage: "https://www.google.com/imgres?imgurl=https%3A%2F%2F5.imimg.com%2Fdata5%2FQU%2FHX%2FMY-18283074%2Fladle-nozzle-500x500.jpg&imgrefurl=https%3A%2F%2Fwww.indiamart.com%2Fproddetail%2Fladle-nozzle-19041966791.html&tbnid=TctIOYvGhikJHM&vet=12ahUKEwiZ7ve_j6bvAhWJUisKHX2FBpwQMygAegUIARCWAQ..i&docid=aK11bg4pdSJnyM&w=500&h=500&q=ladlenozzle&ved=2ahUKEwiZ7ve_j6bvAhWJUisKHX2FBpwQMygAegUIARCWAQ",
        itemCategoryName: "CCM REFACTORY",
    },
    {
        itemName: "Nozzle Filling Compound - NFC",
        itemUnit: "MTS",
        itemQty: 5,
        itemDescription: "Nozzle Filling Compound.",
        itemImage: "https://2.wlimg.com/product_images/bc-full/dir_80/2387626/ladle-nozzle-filling-compound-732820.jpg",
        itemCategoryName: "CCM REFACTORY",
    },
];

// Items.insertMany(seedItems)
//     .then((res) => {
//         console.log(res);
//     })
//     .catch((e) => {
//         console.log(e);
//     });