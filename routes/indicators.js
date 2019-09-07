const Router = require('express-promise-router');
const crypto = require('../controlers/crypto.js');
const wydaj = require('../controlers/wydajn.js');
const termin = require('../controlers/termin.js');
const wydaj_netto = require('../controlers/wydajn_netto.js');
const audit_5s =  require('../controlers/audit5s.js');
const objectives =  require('../controlers/area_objectives.js');
const NodeCache = require( "node-cache" );
const myCache = new NodeCache({ stdTTL: 800, checkperiod: 900 });
const router = new Router();
const path = require('path');
//router.use(function get_permissions (req, res, next) {
//  console.log(req.path);
//  next()
//})

router.get('/restricted/:id', async(req,res,next) =>{
 if(typeof(req.session.verify)!= 'undefined'){
   next();
 } else {
   res.render('error', {
     message:'Proszę o zalogowanie się do strony!',
     error: {status: 'Pobierane dane wymagają uwierzytelnienia'}
   });
 }
});
router.get('/logon', async(req, res) => {
  res.render('logon');
    });
router.get('/restricted/:id', async(req, res) => {
  res.sendFile(path.join(__dirname, 'restricted', req.params.id))
});
router.get('/indicators', async(req, res) => {
  res.render('index4');
    });
router.get('/200SW', async(req, res) => {
    res.render('index0');
    });
router.get('/500LI', async(req, res) => {
      res.render('index1');
    });
router.get('/600M', async(req, res) => {
        res.render('index5');
    });
router.get('/400ST', async(req, res) => {
      res.render('index2');
    });
router.get('/ALL', async(req, res) => {
      res.render('index3');
    });
router.get ('/line', async(req, res) =>{
      res.render('linia');
})
router.get('/AL', async(req, res) => {
      res.render('index');
    });
router.get('/audit_info', async (req, res, next) =>{
      if (typeof(req.session.verify)!= 'undefined') {
        next();
      } else {
        res.render('logon', {msg: 'Logowanie' })
      }
    });
router.get('/audit_info/', async(req, res) => {
      //CHECK IF RESTRICED access
       res.render('audit_info');
    });
router.get('/AUDIT_RESULTS', async(req, res) => {
       res.render('pers_ctrl');
    });
router.get('/data1/:id/:day?/:uuid?', async(req, res, next) => {
    const dep=req.params.id;
    let days=req.params.day;
    if (days ==undefined) {days=33;}
    const value  = await myCache.get('1_' + dep + '_' + days);
    if(value == undefined) {
      next();
    } else {
      if (typeof(req.session.verify)!= 'undefined'){res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
      } else {res.setHeader('Cache-Control', 'public, max-age=200');}
      res.send(await crypto.Encrypt(value,req.session.verify));
    }
})
router.get('/data1/:id/:day?/:uuid?', async(req, res) => {
  const dep=req.params.id;
  let days=req.params.day;
  if (days ==undefined) {days=33;}
  const quer = await IS_one(dep);
    if (quer==5) {
        var que = await termin.Day('400S0',days);
    } else {
        var que = await wydaj.Day(dep,days);
    }
    const  data  = que;
      if (data.error!="yes") {myCache.set('1_' + dep + '_' + days)};
      if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
      } else {res.setHeader('Cache-Control', 'public, max-age=200');}
      res.send(await crypto.Encrypt(data,req.session.verify));
})
router.get('/data2/:id/:week?', async(req, res,next) => {
    const dep=req.params.id;
    let weeks=req.params.week;
    if (weeks == undefined) {weeks=13;}
    const value  = await myCache.get('2_' + dep + '_' + weeks);
    if(value == undefined) {
      next();
    } else {
      if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
      } else {res.setHeader('Cache-Control', 'public, max-age=200');}
      res.send(await crypto.Encrypt(value,req.session.verify));
    }
})
router.get('/data2/:id/:week?', async(req, res) => {
  const dep=req.params.id;
  let weeks=req.params.week;
  if (weeks == undefined) {weeks=13;}
  const quer = await IS_one(dep);
    if (quer==5) {
        var que = await termin.Wk('400S0',weeks);
     } else {
        var que = await wydaj.Wk(dep,weeks);
     }
    const  data  = que;
    if (data.error!="yes") {myCache.set('2_' + dep + '_' + weeks, data)};
    if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    } else {res.setHeader('Cache-Control', 'public, max-age=200');}
    res.send(await crypto.Encrypt(data,req.session.verify));
})
router.get('/data3/:id/:days?', async(req, res,next) => {
  const dep=req.params.id;
  let days=req.params.days;
  if (days == undefined) {days=365;}
  const value  = await myCache.get('3_' + dep + '_' + days);
  if(value == undefined) {
    next();
  } else {
    if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
} else {res.setHeader('Cache-Control', 'public, max-age=200');}
    res.send(await crypto.Encrypt(value,req.session.verify));
  }
})
router.get('/data3/:id/:days?', async(req, res) => {
  const dep=req.params.id;
  let days=req.params.days;
  if (days == undefined) {days=365;}
  const quer = await IS_one(dep);
    if (quer==5) {
        var que = await termin.Mnth('400S0',days);
     } else {
        var que = await wydaj.Mnth(dep,days);
     }
    const  data  = que;
    if (data.error!="yes") {myCache.set('3_' + dep + '_' + days, data)};
    if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    } else {res.setHeader('Cache-Control', 'public, max-age=200');}
    res.send(await crypto.Encrypt(data,req.session.verify));
})
router.get('/data20/:id/:day?', async(req, res,next) => {
  const dep=req.params.id;
  let days=req.params.day;
  if (days == undefined) {days=33;}
  const value  = await myCache.get('20_' + dep + '_' + days);
  if(value == undefined) {
    next();
  } else {
    if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    } else {res.setHeader('Cache-Control', 'public, max-age=200');}
    res.send(await crypto.Encrypt(value,req.session.verify));
  }
})
router.get('/data20/:id/:day?', async(req, res) => {
  const dep=req.params.id;
  let days=req.params.day;
  if (days == undefined) {days=33;}
  const  data  = await termin.Day(dep,days);
    if (data.error!="yes") {myCache.set('20_' + dep + '_' + days)};
    if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    } else {res.setHeader('Cache-Control', 'public, max-age=200');}
    res.send(await crypto.Encrypt(data,req.session.verify));
})
router.get('/data21/:id/:week?', async(req, res,next) => {
  const dep=req.params.id;
  let weeks=req.params.week;
  if (weeks == undefined) {weeks=13;}
  const value  = await myCache.get('21_' + dep + '_' + weeks);
  if(value == undefined) {
      next();
  } else {
    if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    } else {res.setHeader('Cache-Control', 'public, max-age=200');}
    res.send(await crypto.Encrypt(value,req.session.verify));
  }
})
router.get('/data21/:id/:week?', async(req, res) => {
  const dep=req.params.id;
  let weeks=req.params.week;
  if (weeks == undefined) {weeks=13;}
  const data = await termin.Wk(dep,weeks);
  if (data.error!="yes") {myCache.set('21_' + dep + '_' + weeks, data)};
  if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
  } else {res.setHeader('Cache-Control', 'public, max-age=200');}
  res.send(await crypto.Encrypt(data,req.session.verify));
})
router.get('/data16/:id/:day?', async(req, res,next) => {
  const dep=req.params.id;
  let days=req.params.day;
  if (days == undefined) {days=33;}
  const value  = await myCache.get('16_' + dep + '_' + days);
  if(value == undefined) {
      next();
  } else {
    if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    } else {res.setHeader('Cache-Control', 'public, max-age=200');}
      res.send(await crypto.Encrypt(value,req.session.verify));
  }
})
router.get('/data16/:id/:day?', async(req, res) => {
  const dep=req.params.id;
  let days=req.params.day;
  if (days == undefined) {days=33;}
  const  data  = await wydaj_netto.Day(dep,days);
    if (data.error!="yes") {myCache.set('16_' + dep + '_' + days)};
    if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    } else {res.setHeader('Cache-Control', 'public, max-age=200');}
    res.send(await crypto.Encrypt(data,req.session.verify));
})
router.get('/data17/:id/:week?', async(req, res,next) => {
  const dep=req.params.id;
  let weeks=req.params.week;
  if (weeks == undefined) {weeks=13;}
  const value  = await myCache.get('17_' + dep + '_' + weeks);
  if(value == undefined) {
    next();
  } else {
    if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    } else {res.setHeader('Cache-Control', 'public, max-age=200');}
    res.send(await crypto.Encrypt(value,req.session.verify));
  }
})
router.get('/data17/:id/:week?', async(req, res) => {
  const dep=req.params.id;
  let weeks=req.params.week;
  if (weeks == undefined) {weeks=13;}
  const data = await wydaj_netto.Wk(dep,weeks);
  if (data.error!="yes") {myCache.set('17_' + dep + '_' + weeks, data)};
  if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
  } else {res.setHeader('Cache-Control', 'public, max-age=200');}
  res.send(await crypto.Encrypt(data,req.session.verify));
})
router.get('/data18/:id/:days?', async(req, res,next) => {
  const dep=req.params.id;
  let days=req.params.days;
  if (days == undefined) {days=365;}
  const value  = await myCache.get('18_' + dep + '_' + days);
  if(value == undefined) {
      next();
  } else {
    if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    } else {res.setHeader('Cache-Control', 'public, max-age=200');}
    res.send(await crypto.Encrypt(value,req.session.verify));
  }
})
router.get('/data18/:id/:days?', async(req, res) => {
  const dep=req.params.id;
  let days=req.params.days;
  if (days == undefined) {days=365;}
  const data = await wydaj_netto.Mnth(dep,days);
  if (data.error!="yes") {myCache.set('18_' + dep + '_' + days, data)};
  if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
  } else {res.setHeader('Cache-Control', 'public, max-age=200');}
  res.send(await crypto.Encrypt(data,req.session.verify));
})
router.get('/data22/:id/:days?', async(req, res,next) => {
  const dep=req.params.id;
  let days=req.params.days;
  if (days == undefined) {days=365;}
  const value  = await myCache.get('22_' + dep + '_' + days);
  if(value == undefined) {
      next();
  } else {
    if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    } else {res.setHeader('Cache-Control', 'public, max-age=200');}
    res.send(await crypto.Encrypt(value,req.session.verify));
  }
})
router.get('/data22/:id/:days?', async(req, res) => {
  const dep=req.params.id;
  let days=req.params.days;
  if (days == undefined) {days=365;}
  const data = await termin.Mnth(dep,days);
  if (data.error!="yes") {myCache.set('22_' + dep + '_' + days, data)};
  if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
  } else {res.setHeader('Cache-Control', 'public, max-age=200');}
  res.send(await crypto.Encrypt(data,req.session.verify));
})
router.get('/data5/:id/:day?', async(req, res,next) => {
  const dep=req.params.id;
  let days=req.params.day;
  if (days == undefined) {days=33;}
  const value  = await myCache.get('5_' + dep + '_' + days);
  if(value == undefined) {
    next();
  } else {
    if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    } else {res.setHeader('Cache-Control', 'public, max-age=200');}
    res.send(await crypto.Encrypt(value,req.session.verify));
  }
})
router.get('/data5/:id/:day?', async(req, res) => {
  const dep=req.params.id;
  let days=req.params.day;
  if (days == undefined) {days=33;}
  const  data  = await audit_5s.Day(dep,days);
    if (data.error!="yes") {myCache.set('5_' + dep + '_' + days)};
    if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    } else {res.setHeader('Cache-Control', 'public, max-age=200');}
    res.send(await crypto.Encrypt(data,req.session.verify));
})
router.get('/data6/:id/:week?', async(req, res,next) => {
  const dep=req.params.id;
  let weeks=req.params.week;
  if (weeks == undefined) {weeks=13;}
  const value  = await myCache.get('6_' + dep + '_' + weeks);
  if(value == undefined) {
    next();
  } else {
    if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
} else {res.setHeader('Cache-Control', 'public, max-age=200');}
    res.send(await crypto.Encrypt(value,req.session.verify));
  }
})
router.get('/data6/:id/:week?', async(req, res) => {
  const dep=req.params.id;
  let weeks=req.params.week;
  if (weeks == undefined) {weeks=13;}
  const data = await audit_5s.Wk(dep,weeks);
  if (data.error!="yes") {myCache.set('6_' + dep + '_' + weeks, data)};
  if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
  } else {res.setHeader('Cache-Control', 'public, max-age=200');}
  res.send(await crypto.Encrypt(data,req.session.verify));
})
router.get('/data4/:id', async(req, res,next) => {
    const dep=req.params.id;
    const value  = await myCache.get('4_' + dep);
    if(value == undefined) {
      next();
    } else {
      if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
      } else {res.setHeader('Cache-Control', 'public, max-age=200');}
      res.send(await crypto.Encrypt(value,req.session.verify));
    }
})
router.get('/data4/:id', async(req, res) => {
      const dep=req.params.id;
      const data = await objectives.Present_score(dep);
      //    console.log ("Przypisano "+ dep);
      if (data.error!="yes")  {myCache.set('4_' + dep, data)};
      if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
      } else {res.setHeader('Cache-Control', 'public, max-age=200');}
      res.send(await crypto.Encrypt(data,req.session.verify));
})
router.get('/data7/:id/:days?', async(req, res,next) => {
  const dep=req.params.id;
  let days=req.params.days;
  if (days == undefined) {days=365;}
  const value  = await myCache.get('7_' + dep + '_' + days);
  if(value == undefined) {
    next();
  } else {
    if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    } else {res.setHeader('Cache-Control', 'public, max-age=200');}
    res.send(await crypto.Encrypt(value,req.session.verify));
  }
})
router.get('/data7/:id/:days?', async(req, res) => {
  const dep=req.params.id;
  let days=req.params.days;
  if (days == undefined) {days=365;}
  const data = await audit_5s.Mnth(dep,days);
  if (data.error!="yes") {myCache.set('7_' + dep + '_' + days, data)};
  if (typeof(req.session.verify)!= 'undefined'){res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
  } else {res.setHeader('Cache-Control', 'public, max-age=200');}
  res.send(await crypto.Encrypt(data,req.session.verify));
})
router.get('/data10/:id', async(req, res, next) => {
  const dep=req.params.id;
  const value  = await myCache.get('10_' + dep);
  if(value == undefined) {
    next()
  } else {
    if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    } else {res.setHeader('Cache-Control', 'public, max-age=200');}
    res.send(await crypto.Encrypt(value,req.session.verify));
  }
})
router.get('/data10/:id', async(req, res) => {
    const dep=req.params.id;
    const data = await objectives.Score_barLeft(dep);
    //    console.log ("Przypisano "+ dep);
    if (data.error!="yes")  {myCache.set('10_' + dep, data)};
    if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    } else {res.setHeader('Cache-Control', 'public, max-age=200');}
    res.send(await crypto.Encrypt(data,req.session.verify));
})
router.get('/data13/:id/:days?', async(req, res,next) => {
    const dep=req.params.id;
    let days=req.params.days;
    if (days == undefined) {days=365;}
    const value  = await myCache.get('13_' + dep + '_' + days);
    if(value == undefined) {
      next();
    } else {
      if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
      } else {res.setHeader('Cache-Control', 'public, max-age=200');}
      res.send(await crypto.Encrypt(value,req.session.verify));
    }
})
router.get('/data13/:id/:days?', async(req, res) => {
  const dep=req.params.id;
  let days=req.params.days;
  if (days == undefined) {days=365;}
  const data = await objectives.Historical_score(dep,days);
  if (data.error!="yes") {myCache.set('13_' + dep + '_' + days, data)};
  if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
  } else {res.setHeader('Cache-Control', 'public, max-age=200');}
  res.send(await crypto.Encrypt(data,req.session.verify));
})
router.get('/data14/:id', async(req, res,next) => {
  const dep=req.params.id;
  const value  = await myCache.get('14_' + dep);
  if(value == undefined) {
      next();
  } else {
    if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    } else {res.setHeader('Cache-Control', 'public, max-age=200');}
    res.send(await crypto.Encrypt(value,req.session.verify));
  }
})
router.get('/data14/:id', async(req, res) => {
  const dep=req.params.id;
  const data = await objectives.Bar_OnTime(dep);
  //    console.log ("Przypisano "+ dep);
  if (data.error!="yes")  {myCache.set('14_' + dep, data)};
  if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
  } else {res.setHeader('Cache-Control', 'public, max-age=200');}
  res.send(await crypto.Encrypt(data,req.session.verify));
})
router.get('/data11/:id', async(req, res,next) => {
  const dep=req.params.id;
  const value  = await myCache.get('11_' + dep);
  if(value == undefined) {
      next();
  } else {
    if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    } else {res.setHeader('Cache-Control', 'public, max-age=200');}
    res.send(await crypto.Encrypt(value,req.session.verify));
  }
})
router.get('/data11/:id', async(req, res) => {
    const dep=req.params.id;
    const data = await objectives.Score_barRight(dep);
    //    console.log ("Przypisano "+ dep);
    if (data.error!="yes")  {myCache.set('11_' + dep, data)};
    if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    } else {res.setHeader('Cache-Control', 'public, max-age=200');}
    res.send(await crypto.Encrypt(data,req.session.verify));
})
router.get('/data12/:id', async(req, res,next) => {
  const dep=req.params.id;
  const value  = await myCache.get('12_' + dep);
  if(value == undefined) {
    next()
  } else {
    if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    } else {res.setHeader('Cache-Control', 'public, max-age=200');}
    res.send(await crypto.Encrypt(value,req.session.verify));
  }
})
router.get('/data12/:id', async(req, res) => {
  const dep=req.params.id;
  const data = await objectives.Score_barCentral(dep);
  //    console.log ("Przypisano "+ dep);
  if (data.error!="yes")  {myCache.set('12_' + dep, data)};
  if (typeof(req.session.verify)!= 'undefined') {res.setHeader('Cache-Control', 'private, no-cache, no-store, must-revalidate');
  } else {res.setHeader('Cache-Control', 'public, max-age=200');}
  res.send(await crypto.Encrypt(data,req.session.verify));
})
async function IS_one (depar) {
  var ar=';300M1;000M1;200S3;500L4;500L3;400S3;600W0;000M0';
  if (ar.indexOf(depar)!=-1) {
    return 1;
  } else {
    var as =';600M1;600M0;300M0';
    if (as.indexOf(depar)!=-1) {
      return 3;
    } else {
      var es =';100K0;100K1;100K2';
      if (es.indexOf(depar)!=-1) {
        return 4;
      } else {
        var s =';200S2';
        if (s.indexOf(depar)!=-1) {
          return 5;
        } else {
          var e=';100K3';
          if (e.indexOf(depar)!=-1) {
            return 6;
          } else {
        return 2;
        };
      };
    };
  };
 };
};
module.exports = router;
