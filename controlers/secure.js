const db = require('../connect/access.js')
const uid = require('uid-safe')
const Ads="Provider=Microsoft.ace.OLEDB.12.0;Data Source=user_access.accdb;";
// low level funcions
function get_Guid(options) {
  return new Promise((resolve, reject) => {
    uid(options, function (err, string) {
      if(err) {
        reject(err);
      } else {
        resolve(string);
      }
    });
  });
}
async function sigin_up(username,salt,verifier,email) {
  const count_usr=await user_exist(username);
  if (count_usr==0) {
    const options = {
      dsn :  Adsn,
      query : "INSERT INTO USERS ( login, salt, verifier, email ) values ( @login, @salt, @verifier, @email )",
      prepare: "true",
      Values : {
        Val_name1: '@login',
        Value1:username,
        Type1:'VarWChar',
        Len1:124,
        Val_name2: '@salt',
        Value2:salt,
        Type2:'VarWChar',
        Len2:4048,
        Val_name3: '@verifier',
        Value3:verifier,
        Type3:'VarWChar',
        Len3:6024,
        Val_name4: '@email',
        Value4:email,
        Type4:'VarWChar',
        Len4:124,
        }
      };
      const result = await db.Getoledb(options);
      return result
   } else {
      return 'User_exist';
  }
}
async function user_exist(username) {
  const options = {
    dsn :  Ads,
    query : "Select count(login) as exist from USERS where login=@login",
    prepare: "true",
    Values : {
       Val_name1: '@login',
       Value1: username,
       Type1:'VarWChar',
       Len1:50,
      }
  };
  const result = await db.Getoledb(options);
  return result
}
async function user_credent(username) {
  const options = {
    dsn :  Ads,
    query : "Select * from USERS where login=@login",
    prepare: "true",
    Values : {
       Val_name1: '@login',
       Value1:username,
       Type1:'VarWChar',
       Len1:124,
      }
  };
  const result = await db.Getoledb(options);
  return result
}
async function chk_access(username,view) {
  const options = {
    dsn :  Ads,
    query : "Select count(username) from usr_access where username=@username and view=@view ",
    prepare: "true",
    Values : {
       Val_name1: '@username',
       Value1:username,
       Type1:'VarWChar',
       Len1:124,
       Val_name2: '@view',
       Value2:view,
       Type2:'VarWChar',
       Len2:124,
      }
  };
  const result = await db.Getoledb(options);
  return result
}
async function set_access(username,view) {

}
// main functions
async function usr_credentials(username) {
  let exist = await user_exist(username);
  if (exist!='0') {
    let sess_uuid = await get_Guid(18);
    let data = await user_credent(username);
    if (!data.valid || data.records.length==0) {
      return data  ={
        msg : "Error in db",
        error : "yes"
        };
    } else {
    return {
      error : "no",
      uuid: sess_uuid,
      username: data.records[0].login,
      salt:data.records[0].salt,
      verifier:data.records[0].verifier,
      email:data.records[0].email
      }
    }
  } else {
    return {
        msg : "User_don't_exist",
        error : "yes",
    }
  }
}
module.exports = {
  Usr_credentials: (username) => usr_credentials(username)
};
