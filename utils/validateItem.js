// const { itemSchema } = require("../validationSchemas");

// const validateItem = (req, res, next) => {
//     const { error } = itemSchema.validate(req.body);
//     if (error) {
//         const msg = error.details.map((el) => el.message).join(",");
//         throw new ExpressError(msg, 400);
//     } else {
//         next();
//     }
// };

// module.exports = validateItem;
const Joi = require("joi");
const ExpressError = require("./ExpressError"); // Import ExpressError

// Define the validation schema using the code you provided
const itemSchema = Joi.object({
    item: Joi.object({
        itemName: Joi.string().required(),
        itemUnit: Joi.string().required(),
        life: Joi.string().allow(""),
        itemQty: Joi.number().required().min(0),
        itemDescription: Joi.string().required(),
        itemCategoryName: Joi.string().required(),
        // Note: 'createdAt' and 'itemImage' are not included here
        createdAt: Joi.date().required(), // Add this line to allow the date field

        // as they are handled separately (by mongoose/multer) and not part of the form submission data to be validated this way.
    }).required(),
});

// Create and export the middleware function
const validateItem = (req, res, next) => {
    const { error } = itemSchema.validate(req.body);

    if (error) {
        // If validation fails, create a clean error message from Joi's output
        const msg = error.details.map((el) => el.message).join(",");
        // Pass the error to the next error-handling middleware
        return next(new ExpressError(msg, 400));
    } else {
        // If validation succeeds, move to the next function (the route handler)
        next();
    }
};

module.exports = validateItem;