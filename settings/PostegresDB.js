const session_manage = {
  user: 'NoPostgresql',
  password: 'none',
  host: '192.168.0.106',
  application_name: 'Session_manage',
  database: 'zakupy',
  port: 5432,
  max: 1, // max number of clients in the pool
  idleTimeoutMillis: 5000,
};
const pSTGRconf = {
  user: 'NoPostgresql',
  password: 'none',
  host: '192.168.0.106',
  application_name: 'SITS_WEB',
  database: 'zakupy',
  port: 5432,
  max: 10, // max number of clients in the pool
  idleTimeoutMillis: 5000,
};
module.exports = {
  Session_manage: session_manage,
  PSTGRconf: pSTGRconf
};
