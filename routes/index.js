var express = require('express');
var router = express.Router();

const access = require('./indicators');
const pers_ctr = require('./ctrl');


/* GET home page. */

router.get('/AUDIT_info', function (req ,res) {
    res.render('audit_info');
  })
router.get('/AUDIT_RESULTS', function (req ,res) {
  res.render('pers_ctrl');
})


module.exports = router;
