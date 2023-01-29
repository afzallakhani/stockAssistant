const Joi = require("joi");

module.exports.itemSchema = Joi.object({
    item: Joi.object({
        itemName: Joi.string().required(),
        itemUnit: Joi.string().required(),
        itemQty: Joi.number().required().min(0),
        itemDescription: Joi.string().required(),
        itemCategoryName: Joi.string().required(),
    }).required(),
});