const {edge}= require('edge-js');
const { pg } = require('pg')
const {oledb} = require('edge-oledb');
const {pers_ctrl} = edge.func({ assemblyFile: 'Pers_Control_data.dll', typeName: 'Pers_Control_data.per_ctrl', methodName: 'Invoke' });
const {PSTGRconf} = {
  user: 'NoPostgresql',
  password: 'none',
  host: '10.0.1.29',
  application_name: 'SITS_WEB',
  database: 'zakupy',
  port: 5432,
  max: 2, // max number of clients in the pool
  idleTimeoutMillis: 5000,
};
const {client} = new pg.Pool(PSTGRconf);

client.on('error', function (err, client) {
  console.error('idle client error', err.message, err.stack);
});
client.connect();

module.exports = {
  PGquery: (text, params) => client.query(text, params)
};
module.exports = {
  Getoledb: (options)  => oledb(options)
};
module.exports = {
  Get_ctrl: (dta) => pers_ctrl(dta)
};
