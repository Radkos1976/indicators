const Router = require('express-promise-router')
const srp= require('secure-remote-password/server')
const router = new Router()
const NodeCache = require( "node-cache" );
const tmp_session= new NodeCache({ stdTTL: 80, checkperiod: 90 });
const users = require('../controlers/secure.js');
router.post('/client_hello',async(req, res) => {
    const username = req.body.username;
    const cl_epherial = req.body.clientEphemeral;
    
    req.session.verify='';
    let data = await users.Usr_credentials(username);
    if (data.error!="yes") {
      const se_epherial = await srp.generateEphemeral(data.verifier);
      data['cl_epherial']=cl_epherial;
      data['serv_emp_public']=se_epherial.public;
      data['serv_emp_secret']=se_epherial.secret;
      tmp_session.set(data.uuid, data); //zapisz dane w cachu podręcznym sesji
      res.send({
        uuid: data.uuid,
        error : "no",
        emp: data.serv_emp_public,
        slt: data.salt,
      });
    } else {
      res.send(data);
    }
});
router.post('/server_hello',async(req,res) => {
  const uuid=req.body.uuid;
  const clientSession_proof=req.body.proof;
  const value  = await tmp_session.get(uuid);
  if(value != undefined) {
    const serverSession = await srp.deriveSession(value.serv_emp_secret, value.cl_epherial, value.salt, value.username, value.verifier, clientSession_proof);
    let tmp ={
      uuid:uuid,
      key:serverSession.key,
      proof:serverSession.proof
    };
    req.session.verify=tmp;
    res.send ({
      error : "no",
      proof:serverSession.proof
    });
  } else {
    req.session.verify='';
    res.send ({
      msg : "Error in auth",
      error : "yes"
    });
  }
})
router.get('/sess_unsign',async(req,res)=>{
  delete req.session.verify;
  res.send('done');
})
module.exports = router;
