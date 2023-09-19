const mongoose = require("mongoose");

const PartySchema = new mongoose.Schema({
  // itemImage: String,
  partyName: {
    type: String,
    uppercase: true,
  },
  partyAddress: {
    type: String,
    uppercase: true,
  },
  partyType: {
    type: String,
    uppercase: true,
  },
  partyItemCategory: {
    type: String,
    uppercase: true,
  },
});

module.exports = mongoose.model("partyMaster", PartySchema);
