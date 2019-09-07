const { pg } = require('pg')

module.exports = {
  PGquery: (text, params) => client.query(text, params)
};
