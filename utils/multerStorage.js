const multer = require("multer");
// const unlinkAsync = promisify(fs.unlink);

var storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, "views/images");
    },
    filename: (req, file, cb) => {
        cb(null, file.fieldname + "-" + Date.now());
    },
});

module.exports = storage;