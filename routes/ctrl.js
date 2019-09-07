const Router = require('express-promise-router')
const cn = require('../connect/Cconn.js')
const router = new Router()
const crypto = require('../controlers/crypto.js');

/* GET listing. */
router.get('/lstaudit/:id',async(req,res) => {
      const dta ={
        type_dta: 'Ctrl_detalis',
        person_id: req.params.id,
        mnth: ''
      };
      const result = await cn.Get_ctrl(dta);
      res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
      if (result) {
        res.send(await crypto.Encrypt(result,req.session.verify));
      } else {
        res.send({DATA_VALID: false,ERROR: "No results"});
      };
});
router.get('/pers_ctrl/:id',async (req,res) => {
      const dta={
        type_dta: 'Pers_ctrl',
        person_id: req.params.id.substring(7,11),
        mnth: req.params.id.substring(1,7)
      };
      const result = await cn.Get_ctrl(dta);
      res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
      if (result) {
        res.send(await crypto.Encrypt(result,req.session.verify));
      } else {
        res.send({DATA_VALID: false,ERROR: "No results"});
      }
});
router.get('/audit/:id',async (req,res) => {
    if( req.params.id.substring(0,12)=='Lst_ctrl_err' || req.params.id.substring(0,12)=='Emp_Ctrl_sta' || req.params.id.substring(0,12)=='PER_ctrl_enh')
    {
        var mnth = req.params.id.substring(12,18);
        var person_id = req.params.id.substring(18,26);
    }
    else
    {
        var mnth= req.params.id.substring(12,19);
        var person_id= ' ';
    }
    const dta={
        type_dta: req.params.id.substring(0,12),
        mnth: mnth,
        person_id: person_id
    }
      const result = await cn.Get_ctrl(dta);
      res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
      if (result) {
        res.send(await crypto.Encrypt(result,req.session.verify));
      } else {
        res.send({DATA_VALID: false,ERROR: "No results"});
      }
});

module.exports = router;
