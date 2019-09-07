const Router = require('express-promise-router');
const router = new Router();

router.get('/indicators/:id/:dep', async(req,res) => {
    let qr= await IS_one(req.params.dep);
    switch (req.params.id) {
      case 4:
        let ctx
        if (qr==4) {
          ctx ={
            visible=true,
            witdh=
            height=
            datasource=
          }
        }

        let tmp {
          _conn=1,
          _conn_adr="/data13/",

        }
        break;
      default:

    }
});
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
