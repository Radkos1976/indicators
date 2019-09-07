const cryptoJSON = require('crypto-json')

async function encrypt(input,options) {
  let data = {
    crypto:"NO",
    data: input
  };
  if (typeof(options)!='undefined') {
    try {
      const result=await cryptoJSON.encrypt(input,options.key);
      data = {
        crypto:"YES",
        data: result
      };
    } catch (e) {
      data = {
        crypto:"NO",
        data: input
      };
    } 
  };
  return data;
}
module.exports = {
  Encrypt: (input,options) => encrypt(input,options)
};
