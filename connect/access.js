const  edge   = require('edge-js');
const  util = require('util');
const  oledb  = edge.func({ assemblyFile: 'Class_odbc.dll', typeName: 'Class_odbc.Startup', methodName: 'Invoke' });
function get_ODBC(options) {
  return new Promise((resolve, reject) => {
    oledb(options, (error,result) => {
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

async function get(options) {
  var data ={}
  const result=await get_ODBC(options);
  data.valid = true;
  data.records = result;
  return data
}

module.exports = {
  Getoledb: (options) => get(options)
};
