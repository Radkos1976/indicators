const pg_settings=require('../settings/PostegresDB.js');

const PSTGRconf = pg_settings.PSTGRconf;
const pg = require('pg');
var pool = new pg.Pool(PSTGRconf);

pool.on('error', function (err, client) {
  console.error('idle client error', err.message, err.stack);
});

module.exports = {
  Query: (text, params) => pool.query(text, params)
};
