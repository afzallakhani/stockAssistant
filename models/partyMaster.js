const mongoose = require("mongoose");

const PartySchema = new mongoose.Schema({
  // itemImage: String,
  partyName: String,
  partyAddress: String,
  partyType: String,
  partyItemCategory: String,
});

module.exports = mongoose.model("partyMaster", PartySchema);
