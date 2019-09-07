const pg = require('../connect/Pg_conn.js');
function get_PG(options) {
  return new Promise((resolve, reject) => {
    pg.Query(options, (error,result) => {
      if(error) {
        reject(error);}
      else {
        try {
          resolve(result);
        } catch (e) {
          reject({ message: e.message });
        }
      }
    });
  });
}
async function chk_is_allow() {
  let ifs_allow=true;
   try {
      const result = await get_PG("select in_progress from datatbles where table_name ='ifs_CONN'");
      ifs_allow=result.rows[0].in_progress;
   } catch (e) {
    console.log('Błąd ' + e);
      ifs_allow=true;
   } finally {
     return ifs_allow;
   }
}
module.exports = {
  Chk_is_allow:() => chk_is_allow()
};
