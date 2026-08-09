const { ClientSecretCredential } = require("@azure/identity");

const credential = new ClientSecretCredential(
  process.env.TENANT_ID,
  process.env.CLIENT_ID,
  process.env.CLIENT_SECRET
);

module.exports = credential;