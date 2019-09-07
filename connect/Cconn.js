const edge = require('edge-js');
const pers_ctrl  = edge.func({ assemblyFile: 'Pers_Control_data.dll', typeName: 'Pers_Control_data.per_ctrl', methodName: 'Invoke' });
function get_ctls(dta) {
  return new Promise((resolve, reject) => {
    pers_ctrl(dta, (error,result) => {
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
module.exports = {
  Get_ctrl: (dta) => get_ctls(dta)
};
