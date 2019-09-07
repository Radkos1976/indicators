const db = require('../connect/access.js')
//const Json2csvParser = require('json2csv').Parser;
const fs = require('fs');
const Adsn="Provider=Microsoft.ace.OLEDB.12.0;Data Source=WYDAJ.accdb;Persist Security Info=False;Mode=Read;";
async function IS_one (DEPar) {
  var ar=';300M1;000M1;200S3;500L4;500L3;400S3;600W0;000M0';
  if (ar.indexOf(DEPar)!=-1) {
    return 1;
  } else {
    var as =';600M1;600M0;300M0';
    if (as.indexOf(DEPar)!=-1) {
      return 3;
    } else {
      var es =';100K0;100K1;100K2';
      if (es.indexOf(DEPar)!=-1) {
        return 4;
      } else {
        var s =';200S2';
        if (s.indexOf(DEPar)!=-1) {
          return 5;
        } else {
          var e=';100K3';
          if (e.indexOf(DEPar)!=-1) {
            return 6;
          } else {
        return 2;
        };
      };
    };
  };
 };
};
async function present_score(DEP) {
  const quer=await IS_one(DEP);
  if (quer==2 || quer==5) {
      var que= "Select ta.poziom5s,ta.poziom,ta.prog,ta.wydaj from (SELECT top 1 b.descr ,cdbl(dateserial(mid(a.month,1,4),mid(a.month,5,2),1)) as mont, round(a.wyn*100,2) AS Wynik FROM (SELECT b.month,mid(a.mpk,1,4) as mpk, avg(a.wynik) AS wyn FROM wynik5s AS a, (SELECT month, week FROM wydaj where month>format(now()-60,'yyyymm') GROUP BY month, week)  AS b WHERE a.tydzień=b.week GROUP BY b.month,mpk)  AS a ,(select * from dpt where dp_id=@DEP) AS b WHERE  a.mpk=b.mpk order by a.month desc) as Act,(select * from (select poziom5s,poziom,prog,wydaj,cdbl(Ważny_OD) as OD,cdbl(Ważny_DO) as DO from progi where dp_id=@DEP and prog<>0) union (select '<' & poziom5s ,0 as poziom,0 as prog , '<' & wydaj,cdbl(Ważny_OD) as OD,cdbl(Ważny_do) as DO  from progi where prog=1 and  dp_id=@DEP)) as ta where mont BETWEEN  ta.od and ta.do order by prog;"
  } else if (quer==1) {
      var que="Select ta.poziom,ta.prog,ta.wydaj from (SELECT top 1 b.descr ,cdbl(dateserial(mid(a.month,1,4),mid(a.month,5,2),1)) as mont, round(a.wyn*100,2) AS Wynik FROM (SELECT b.month,mid(a.mpk,1,4) as mpk, avg(a.wynik) AS wyn FROM wynik5s AS a, (SELECT month, week FROM wydaj where month>format(now()-60,'yyyymm') GROUP BY month, week)  AS b WHERE a.tydzień=b.week GROUP BY b.month,mpk)  AS a ,(select * from dpt where dp_id='400S0' ) AS b WHERE  a.mpk=b.mpk order by a.month desc) as Act,(select * from (select poziom,prog,wydaj,cdbl(Ważny_OD) as OD,cdbl(Ważny_DO) as DO from progi where dp_id=@DEP and prog<>0) union (select 0 as poziom,0 as prog , '<' & wydaj,cdbl(Ważny_OD) as OD,cdbl(Ważny_DO) as DO  from progi where prog=1 and  dp_id=@DEP) order by prog) as ta where mont BETWEEN  ta.od and ta.do order by prog;"
  } else if (quer==3 || quer==6) {
      var que= "Select ta.brak as braki,ta.poziom5s,ta.poziom,ta.prog,ta.wydaj from (SELECT top 1 b.descr ,cdbl(dateserial(mid(a.month,1,4),mid(a.month,5,2),1)) as mont, round(a.wyn*100,2) AS Wynik FROM (SELECT b.month,mid(a.mpk,1,4) as mpk, avg(a.wynik) AS wyn FROM wynik5s AS a, (SELECT month, week FROM wydaj where month>format(now()-60,'yyyymm') GROUP BY month, week)  AS b WHERE a.tydzień=b.week GROUP BY b.month,mpk)  AS a ,(select * from dpt where dp_id=@DEP) AS b WHERE  a.mpk=b.mpk order by a.month desc) as Act,(select * from (select format(braki,'###0.000') as brak,poziom5s,poziom,prog,wydaj,cdbl(Ważny_OD) as OD,cdbl(Ważny_DO) as DO from progi where dp_id=@DEP and prog<>0) union (select '>' & format(braki,'###0.000'), '<' & poziom5s ,0 as poziom,0 as prog , '<' & wydaj,cdbl(Ważny_OD) as OD,cdbl(Ważny_DO) as DO  from progi where prog=1 and  dp_id=@DEP) order by prog) as ta where mont BETWEEN  ta.od and ta.do order by prog;"
  } else if (quer==4) {
      var que= "Select ta.brak as braki,ta.poziom5s,ta.poziom,ta.prog,ta.wydaj,ta.wydajn_next_lev from (SELECT top 1 b.descr ,cdbl(dateserial(mid(a.month,1,4),mid(a.month,5,2),1)) as mont,round(a.wyn*100,2) AS Wynik FROM (SELECT b.month,mid(a.mpk,1,4) as mpk, avg(a.wynik) AS wyn FROM wynik5s AS a,(SELECT month, week FROM wydaj where month>format(now()-60,'yyyymm') GROUP BY month, week)  AS b WHERE a.tydzień=b.week GROUP BY b.month,mpk)  AS a ,(select * from dpt where dp_id=@DEP) AS b WHERE a.mpk=b.mpk order by a.month desc) as Act,(select * from (select format(braki,'###0.000') as brak,poziom5s,poziom,prog,wydaj,wydajn_next_lev,cdbl(Ważny_OD) as OD,cdbl(Ważny_DO) as DO from progi where dp_id=@DEP and prog<>0) union (select '>' & format(braki,'###0.000'), '<' & poziom5s ,0 as poziom,0 as prog , '<' & wydaj,'<' & wydajn_next_lev,cdbl(Ważny_OD) as OD,cdbl(Ważny_DO) as DO  from progi where prog=1 and  dp_id=@DEP) order by prog) as ta where mont BETWEEN  ta.od and ta.do order by prog; "
  };
  const  options   = {
    dsn :  Adsn,
    query : que,
    prepare: "true",
    Values : {
       Val_name1: '@DEP',
       Value1:DEP,
       Type1:'VarWChar',
       Len1:5,
      }
  }
  const  result= await db.Getoledb(options);
  if (!result.valid) {
    //      console.log ("Błąd /" + result.error + "/");
    var data={
      error : "yes"
    };
    chk=false;
  } else {
    //    console.log ("Dane pobrane z ACCESS")
    if (quer==2) {
      var data = {
        error : "no",
        title : "TABELA CELÓW DLA OBSZARU",
        table :[
          result.records.map(function(it){ return it.prog}),
          result.records.map(function(it){ return it.wydaj}),
          result.records.map(function(it){ return it.poziom5s}),
          result.records.map(function(it){ return it.poziom}),
        ],
        columns : ["Próg","Wydajność","5-S","Poziom"],
        waga : ["Waga","75%","25%"],
      };
    } else if (quer==1) {
      var data = {
        error : "no",
        title : "TABELA CELÓW DLA OBSZARU",
        table :[
          result.records.map(function(it){ return it.prog}),
          result.records.map(function(it){ return it.wydaj}),
          result.records.map(function(it){ return it.poziom}),
        ],
        columns : ["Próg","Wydajność","Poziom"],
        waga : ["Waga","100%"],
      };
    } else if (quer==3) {
      var data = {
        error : "no",
        title : "TABELA CELÓW DLA OBSZARU",
        table :[
          result.records.map(function(it){ return it.prog}),
          result.records.map(function(it){ return it.wydaj}),
          result.records.map(function(it){ return it.poziom5s}),
          result.records.map(function(it){ return it.braki}),
          result.records.map(function(it){ return it.poziom}),
        ],
        columns : ["Próg","Wydajność","5-S","Braki","Poziom"],
        waga : ["Waga","60%","25%","15%"],
      };
    } else if (quer==4) {
      var data = {
        error : "no",
        title : "TABELA CELÓW DLA OBSZARU",
        table :[
          result.records.map(function(it){ return it.prog}),
          result.records.map(function(it){ return it.wydaj}),
          result.records.map(function(it){ return it.poziom5s}),
          result.records.map(function(it){ return it.braki}),
          result.records.map(function(it){ return it.wydajn_next_lev}),
          result.records.map(function(it){ return it.poziom}),
        ],
        columns : ["Próg","Wyd.Szwal.","5-S","Odpad.","Wyd.kroj.","Poziom"],
        waga : ["Waga","10%","25%","50%","15%"],
        };
    } else if (quer==5) {
      var data = {
        error : "no",
        title : "TABELA CELÓW DLA OBSZARU",
        table :[
          result.records.map(function(it){ return it.prog}),
          result.records.map(function(it){ return it.wydaj}),
          result.records.map(function(it){ return it.poziom5s}),
          result.records.map(function(it){ return it.poziom}),
        ],
        columns : ["Próg","Term.Tapic.","5-S","Poziom"],
        waga : ["Waga","70%","30%"],
      };
    } else if (quer==6) {
      var data = {
        error : "no",
        title : "TABELA CELÓW DLA OBSZARU",
        table :[
          result.records.map(function(it){ return it.prog}),
          result.records.map(function(it){ return it.wydaj}),
          result.records.map(function(it){ return it.poziom5s}),
          result.records.map(function(it){ return it.braki}),
          result.records.map(function(it){ return it.poziom}),
        ],
        columns : ["Próg","Wyd.Szwal.","5-S","Rozp.Wałki","Poziom"],
        waga : ["Waga","25%","25%","50%"],
      };
    }
  };
  return data
}
async function score_barLeft(DEP) {
    const quer=await IS_one(DEP);
    if (quer==3) {
        var que= "SELECT top 1 b.descr,a.Month,Format(dateserial(left(a.month,4),mid(a.month,5,2),1),'mmmm yyyyr') as Moh, round(a.wsp_brk,4) AS wyda FROM mgwis as a,dpt AS b WHERE a.month>format(now()-40,'yyyymm') and  ((b.dp_id)=@DEP) and b.mpk=a.mpk order by a.month desc; "
    } else if (quer==4) {
        var que= "SELECT top 1 b.descr,a.Month,Format(dateserial(left(a.month,4),mid(a.month,5,2),1),'mmmm yyyyr') as Moh, round(a.wsp_brk*100,2) AS wyda FROM mgwis as a,dpt AS b WHERE a.month>format(now()-40,'yyyymm') and  ((b.dp_id)=@DEP) and b.ext1=a.mpk order by a.month desc; "
    } else if (quer==5) {
        var que= "SELECT top 1 b.descr,a.Month,Format(dateserial(left(a.month,4),mid(a.month,5,2),1),'mmmm yyyyr') as Moh, round(a.wynik_term*100,2) AS wyda FROM TERMINOWOŚ_month as a,dpt AS b WHERE a.month>format(now()-40,'yyyymm') and  ((b.dp_id)=@DEP) and b.ext=a.wydz order by a.month desc;"
    } else if (quer==6) {
        var que= "SELECT top 1 b.descr,a.Month,Format(dateserial(left(a.month,4),mid(a.month,5,2),1),'mmmm yyyyr') as Moh, round(a.wsp_brk,2) AS wyda FROM mgwis as a,dpt AS b WHERE a.month>format(now()-40,'yyyymm') and  ((b.dp_id)=@DEP) and b.ext1=a.mpk order by a.month desc; "
    } else {
        var que= "SELECT top 1 b.descr,a.Month,Format(a.valid_date,'mmmm yyyyr') as Moh, round((sum(a.SUM_H_MADE_CORR)/sum(a.WRK_H_SUM*7.5/8))*100,2) AS wyda FROM wydaj as a, dpt AS b WHERE a.valid_date>now()-40 and [a.END]=true and  ((b.dp_id)=@DEP) AND iif(b.department,a.department in (select department from dept where dp_id=@DEP group by department),a.department is not null) and iif(b.workcnt,a.work_center in (select workcnt from dept where dp_id=@DEP group by workcnt),a.work_center is not null) and iif(b.notworkcnt,a.work_center not in (select notworkcnt from dept where dp_id=@DEP group by notworkcnt),a.work_center is not null) GROUP BY b.descr,a.month ,Format(a.valid_date,'mmmm yyyyr') order by a.month desc;"
    }
    const  options   = {
      dsn :  Adsn,
      query : que,
      prepare: "true",
      Values : {
         Val_name1: '@DEP',
         Value1:DEP,
         Type1:'VarWChar',
         Len1:5,
        }
    }
    //    console.log ("Ustawienia ACCESS")
    const  result= await db.Getoledb(options);
    if (!result.valid) {
      //    console.log ("Błąd /" + result.error + "/");
      var data = {
        error : "yes"
      };
      chk=false;
    } else {
      try {
      //      console.log ("Dane pobrane z ACCESS")
      if (quer==3) {
        var data = {
          error : "no",
          title: "BRAKI / REKLAMACJE",
          subtitle: result.records[0].descr + " " + result.records[0].Moh.toUpperCase(),
          value : result.records[0].wyda,
          Mnt : result.records[0].Moh.toUpperCase(),
          color: ["rgba(194,152,44,0.5)"],
          jedn: '\u2030',
        };
      } else if (quer==4) {
        var data = {
          error : "no",
          title: "ODPADOWOŚĆ",
          subtitle: result.records[0].descr + " " + result.records[0].Moh.toUpperCase(),
          value : result.records[0].wyda,
          Mnt : result.records[0].Moh.toUpperCase(),
          color: ["rgba(45,212,195,0.4)"],
          jedn: '%',
        };
      } else if (quer==5) {
        var data = {
          error : "no",
          title: "PROD. NA CZAS",
          subtitle: "TAPICERNIA " + result.records[0].Moh.toUpperCase(),
          value : result.records[0].wyda,
          Mnt : result.records[0].Moh.toUpperCase(),
          color: ["rgba(180,152,197,0.5)"],
          jedn: '%',
        };
      } else if (quer==6) {
        var data = {
          error : "no",
          title: "WSKAŹNIK ROZP. WAŁKÓW",
          subtitle: result.records[0].descr + " " + result.records[0].Moh.toUpperCase(),
          value : result.records[0].wyda,
          Mnt : result.records[0].Moh.toUpperCase(),
          color: ["rgba(45,212,195,0.4)"],
          jedn: ' ',
        };
      } else {
        var data ={
          error : "no",
          title: "WYDAJNOŚĆ",
          subtitle: result.records[0].descr + " " + result.records[0].Moh.toUpperCase(),
          value : result.records[0].wyda,
          Mnt : result.records[0].Moh.toUpperCase(),
          color: ["rgba(255, 79, 48, 0.3)"],
          jedn: '%',
        };
      }
    }
    catch (err){
      var data = {
        error : "yes"
      };
    }
  }
  return data;
}
async function score_barRight(DEP) {
  const quer=await IS_one(DEP);
    if (quer==4) {
        var que= "SELECT top 1 b.descr,a.Month,Format(a.valid_date,'mmmm yyyyr') as mont, round(sum(a.SUM_H_MADE_CORR)/sum(a.WRK_H_SUM),2) AS Wynik FROM wydaj as a, dpt AS b WHERE a.valid_date>now()-40 and [a.END]=true and  ((b.dp_id)=@DEP) AND a.work_center=b.ext GROUP BY b.descr,a.month ,Format(a.valid_date,'mmmm yyyyr') order by a.month desc;  "
    } else {
        var que= "SELECT top 1 b.descr, Format(mid(a.month,5,2) & '-01-' & mid(a.month,1,4),'mmmm yyyyr') as mont, round(a.wyn*100,2) AS Wynik FROM (SELECT b.month,mid(a.mpk,1,4) as mpk, avg(a.wynik) AS wyn FROM wynik5s AS a, (SELECT month, week FROM wydaj where month>format(now()-60,'yyyymm') GROUP BY month, week)  AS b WHERE a.tydzień=b.week GROUP BY b.month,mpk)  AS a ,(select * from dpt where dp_id=@DEP) AS b WHERE  a.mpk=b.mpk order by a.month desc;"
    }
    const options = {
      dsn :  Adsn,
      query : que,
      prepare: "true",
      Values : {
         Val_name1: '@DEP',
         Value1:DEP,
         Type1:'VarWChar',
         Len1:5,
        }
    }
    //   console.log ("Ustawienia ACCESS")
    const  result= await db.Getoledb(options);
    if (!result.valid || quer==1) {
      //     console.log ("Błąd /" + result.error + "/");
     var data={
       error : "yes"
     };
     chk=false;
    } else {
    //     console.log ("Dane pobrane z ACCESS")
     if (quer==4) {
       var data={
          error : "no",
          title: "WYDAJNOŚĆ",
          subtitle: result.records[0].descr + " " + result.records[0].mont.toUpperCase(),
          value : result.records[0].Wynik,
          Mnt : result.records[0].mont.toUpperCase(),
          color: ["rgba(194,152,44,0.5)"],
          jedn: 'mb',
       }
     } else {
      var data = {
        error : "no",
        title: "WYNIK 5-S",
        subtitle: result.records[0].descr + " " + result.records[0].mont.toUpperCase(),
        value : result.records[0].Wynik,
        Mnt : result.records[0].mont.toUpperCase(),
        color: ["rgba(151,187,205,0.5)"],
        jedn: '%',
      }
    }
   };
   return data
}
async function score_barCentral(DEP) {
  const quer=await IS_one(DEP);
  if (quer==2) {
    var que= "select a.descr,a.mont,(max(b.poziom)*0.25+max(d.poziom)*0.75)*iif(max(b.poziom)=0 or max(d.poziom)=0,0.5,1) as Wynik,max(b.poziom) as five_S,max(d.poziom) as wyd,c.wyda,a.wynik as wys,b.jednostka,a.month from (SELECT top 2 b.descr,a.month,Format(mid(a.month,5,2) & '-01-' & mid(a.month,1,4),'mmmm yyyyr') as mont, round(a.wyn*100,2) AS Wynik FROM (SELECT b.month,mid(a.mpk,1,4) as mpk, avg(a.wynik) AS wyn FROM wynik5s AS a, (SELECT month, week FROM wydaj where month>format(now()-60,'yyyymm') GROUP BY month, week)  AS b WHERE a.tydzień=b.week GROUP BY b.month,mpk)  AS a ,(select * from dpt where dp_id=@DEP) AS b  WHERE  a.mpk=b.mpk order by a.month desc) as a,(SELECT top 2 b.descr,a.Month,Format(a.valid_date,'mmmm yyyyr') as Moh, round((sum(a.SUM_H_MADE_CORR)/sum(a.WRK_H_SUM*7.5/8))*100,2) AS wyda FROM wydaj as a, dpt AS b WHERE a.valid_date>now()-60 and [a.END]=true and ((b.dp_id)=@DEP) AND iif(b.department,a.department in (select department from dept where dp_id=@DEP group by department),a.department is not null) and iif(b.workcnt,a.work_center in (select workcnt from dept where dp_id=@DEP group by workcnt),a.work_center is not null) and iif(b.notworkcnt,a.work_center not in (select notworkcnt from dept where dp_id=@DEP group by notworkcnt),a.work_center is not null) GROUP BY b.descr,a.month ,Format(a.valid_date,'mmmm yyyyr') order by a.month desc) as c,(select Poziom5s,Poziom,jednostka,ważny_od,ważny_do from progi where dp_id=@DEP) as b,(select wydaj,Poziom,ważny_od,ważny_do from progi where dp_id=@DEP) as d  where a.month=c.month and (a.wynik>=b.poziom5S and (dateserial(left(a.month,4),mid(a.month,5,2),1) between  b.ważny_od and b.ważny_do) )  and  (c.wyda>=d.wydaj  and (dateserial(left(a.month,4),mid(a.month,5,2),1) between  d.ważny_od and d.ważny_do)) group by a.descr,a.mont,c.wyda,a.wynik,b.jednostka,a.month order by a.month desc; "
   } else if (quer==1) {
       var que= "select c.descr,c.moh,max(d.poziom) as Wynik,c.wyda,d.jednostka,c.month from (SELECT top 2 b.descr,a.Month,Format(a.valid_date,'mmmm yyyyr') as Moh, round((sum(a.SUM_H_MADE_CORR)/sum(a.WRK_H_SUM*7.5/8))*100,2) AS wyda FROM wydaj as a,dpt AS b WHERE a.valid_date>now()-60 and [a.END]=true and  ((b.dp_id)=@DEP) AND iif(b.department,a.department in (select department from dept where dp_id=@DEP group by department),a.department is not null) and iif(b.workcnt,a.work_center in (select workcnt from dept where dp_id=@DEP group by workcnt),a.work_center is not null) and iif(b.notworkcnt,a.work_center not in (select notworkcnt from dept where dp_id=@DEP group by notworkcnt),a.work_center is not null) GROUP BY b.descr,a.month ,Format(a.valid_date,'mmmm yyyyr') order by a.month desc) as c,(select wydaj,Poziom,jednostka,ważny_od,ważny_do from progi where dp_id=@DEP) as d where  c.wyda>=d.wydaj and dateserial(left(c.month,4),mid(c.month,5,2),1) between d.ważny_od and d.ważny_do group by  c.descr,c.moh,c.wyda,d.jednostka,c.month order by c.month desc;"
   } else if (quer==3) {
       var que= "select a.descr,a.mont,(max(b.poziom)*0.25+max(d.poziom)*0.6+max(f.poziom)*0.15)*iif(max(b.poziom)=0 or max(d.poziom)=0 or max(f.poziom)=0,0.5,1) as Wynik,max(b.poziom) as five_S,max(d.poziom) as wyd,max(f.poziom) as p_brak,c.wyda,a.wynik as wys,e.braki,b.jednostka,a.month from (SELECT top 2 b.descr,a.month,Format(mid(a.month,5,2) & '-01-' & mid(a.month,1,4),'mmmm yyyyr') as mont, round(a.wyn*100,2) AS Wynik FROM (SELECT b.month,mid(a.mpk,1,4) as mpk, avg(a.wynik) AS wyn FROM wynik5s AS a,(SELECT month, week FROM wydaj where month>format(now()-60,'yyyymm') GROUP BY month, week)  AS b WHERE a.tydzień=b.week GROUP BY b.month,mpk)  AS a ,(select * from dpt where dp_id=@DEP ) AS b  WHERE  a.mpk=b.mpk order by a.month desc) as a,(SELECT top 2 b.descr,a.Month,Format(a.valid_date,'mmmm yyyyr') as Moh, round((sum(a.SUM_H_MADE_CORR)/sum(a.WRK_H_SUM*7.5/8))*100,2) AS wyda FROM wydaj as a,dpt AS b WHERE a.valid_date>now()-60 and [a.END]=true and  ((b.dp_id)=@DEP) AND iif(b.department,a.department in (select department from dept where dp_id=@DEP group by department),a.department is not null) and iif(b.workcnt,a.work_center in (select workcnt from dept where dp_id=@DEP group by workcnt),a.work_center is not null) and iif(b.notworkcnt,a.work_center not in (select notworkcnt from dept where dp_id=@DEP group by notworkcnt),a.work_center is not null) GROUP BY b.descr,a.month ,Format(a.valid_date,'mmmm yyyyr') order by a.month desc) as c,(select Poziom5s,Poziom,jednostka from progi where dp_id=@DEP) as b,(select wydaj,Poziom from progi where dp_id=@DEP) as d,(SELECT top 2 b.descr,a.month,Format(mid(a.month,5,2) & '-01-' & mid(a.month,1,4),'mmmm yyyyr') as mont, round(a.WSp_BRK,4) AS Braki FROM (SELECT month,mpk,WSp_BRK FROM MGWiS AS a),(select * from dpt where dp_id=@DEP ) AS b  WHERE  a.mpk=b.mpk order by a.month desc) as e,(select braki,Poziom from progi where dp_id=@DEP) as f  where a.month=c.month and e.month=c.month and a.wynik>=b.poziom5S and  c.wyda>=d.wydaj and e.braki<=f.braki group by a.descr,a.mont,c.wyda,a.wynik,e.braki,b.jednostka,a.month order by a.month desc;"
   } else if (quer==4) {
       var que= "Select a.descr,a.mont,(max(b.poziom)*0.15+max(d.poziom)*0.1+max(f.poziom)*0.5+max(h.poziom)*0.25)*iif(max(b.poziom)=0 or max(d.poziom)=0 or max(f.poziom)=0 or max(h.poziom)=0,0.5,1) as Wynik,max(b.poziom) as five_S,max(d.poziom) as wyd,max(f.poziom) as p_brak,max(h.poziom) as p_wyd_MB,c.wyda,a.wynik as wys,e.braki,g.wyda_kroj as wyd_MB,b.jednostka,a.month from (SELECT top 2 b.descr,a.month,Format(mid(a.month,5,2) & '-01-' & mid(a.month,1,4),'mmmm yyyyr') as mont, round(a.wyn*100,2) AS Wynik FROM (SELECT b.month,mid(a.mpk,1,4) as mpk, avg(a.wynik) AS wyn FROM wynik5s AS a,(SELECT month, week FROM wydaj where month>format(now()-60,'yyyymm') GROUP BY month, week)  AS b WHERE a.tydzień=b.week GROUP BY b.month,mpk)  AS a ,(select * from dpt where dp_id=@DEP ) AS b  WHERE  a.mpk=b.mpk order by a.month desc) as a,(SELECT top 2 b.descr,a.Month,Format(a.valid_date,'mmmm yyyyr') as Moh, round((sum(a.SUM_H_MADE_CORR)/sum(a.WRK_H_SUM*7.5/8))*100,2) AS wyda FROM wydaj as a,dpt AS b WHERE a.valid_date>now()-60 and [a.END]=true and  ((b.dp_id)=@DEP) AND iif(b.department,a.department in (select department from dept where dp_id=@DEP group by department),a.department is not null) and iif(b.workcnt,a.work_center in (select workcnt from dept where dp_id=@DEP group by workcnt),a.work_center is not null) and iif(b.notworkcnt,a.work_center not in (select notworkcnt from dept where dp_id=@DEP group by notworkcnt),a.work_center is not null) GROUP BY b.descr,a.month ,Format(a.valid_date,'mmmm yyyyr') order by a.month desc) as c,(select Poziom5s,Poziom,jednostka from progi where dp_id=@DEP) as b,(select wydaj,Poziom from progi where dp_id=@DEP) as d,(SELECT top 2 b.descr,a.month,Format(mid(a.month,5,2) & '-01-' & mid(a.month,1,4),'mmmm yyyyr') as mont, round(a.WSp_BRK*100,4) AS Braki FROM (SELECT month,mpk,WSp_BRK FROM MGWiS AS a),(select * from dpt where dp_id=@DEP ) AS b  WHERE  a.mpk=b.ext1 order by a.month desc) as e,(select braki,Poziom from progi where dp_id=@DEP) as f,(SELECT top 2 b.descr,a.Month,Format(a.valid_date,'mmmm yyyyr') as Moh, round((sum(a.SUM_H_MADE_CORR)/sum(a.WRK_H_SUM)),2) AS wyda_kroj FROM wydaj as a,dpt AS b WHERE a.valid_date>now()-60 and [a.END]=true and  ((b.dp_id)=@DEP) AND a.work_center=b.ext GROUP BY b.descr,a.month ,Format(a.valid_date,'mmmm yyyyr') order by a.month desc) as g,(select Wydajn_next_lev as nxt_lev,Poziom from progi where dp_id=@DEP) as h where a.month=c.month and g.Month=c.month and e.month=c.month and a.wynik>=b.poziom5S and  c.wyda>=d.wydaj and e.braki<=f.braki and g.wyda_kroj>=h.nxt_lev group by a.descr,a.mont,c.wyda,a.wynik,e.braki,g.wyda_kroj,b.jednostka,a.month order by a.month desc;"
   } else if (quer==5) {
       var que = "select a.descr,a.mont,(max(b.poziom)*0.3+max(d.poziom)*0.7)*iif(max(b.poziom)=0 or max(d.poziom)=0,0.5,1) as Wynik,max(b.poziom) as five_S,max(d.poziom) as wyd,c.wyda,a.wynik as wys,b.jednostka,a.month from (SELECT top 2 b.descr,a.month,Format(mid(a.month,5,2) & '-01-' & mid(a.month,1,4),'mmmm yyyyr') as mont, round(a.wyn*100,2) AS Wynik FROM (SELECT b.month,mid(a.mpk,1,4) as mpk, avg(a.wynik) AS wyn FROM wynik5s AS a,(SELECT month, week FROM wydaj where month>format(now()-60,'yyyymm') GROUP BY month, week)  AS b WHERE a.tydzień=b.week GROUP BY b.month,mpk)  AS a ,(select * from dpt where dp_id=@DEP ) AS b  WHERE  a.mpk=b.mpk order by a.month desc) as a,(SELECT top 2 b.descr,a.Month,Format(DateValue(right(a.Month,2) & '/' & '/01/' & Left(a.Month,4)),'mmmm yyyyr') as Moh,round(wynik_term*100,2) AS wyda FROM TERMINOWOŚ_month as a,dpt AS b WHERE a.Month>format(now()-60,'yyyymm') and  ((b.dp_id)=@DEP) AND a.wydz=b.ext order by a.month desc) as c,(select Poziom5s,Poziom,jednostka,ważny_od,ważny_do from progi where dp_id=@DEP) as b,(select wydaj,Poziom,ważny_od,ważny_do from progi where dp_id=@DEP) as d where a.month=c.month and (a.wynik>=b.poziom5S and (dateserial(left(a.month,4),mid(a.month,5,2),1) between  b.ważny_od and b.ważny_do) ) and  (c.wyda>=d.wydaj  and (dateserial(left(a.month,4),mid(a.month,5,2),1) between  d.ważny_od and d.ważny_do)) group by a.descr,a.mont,c.wyda,a.wynik,b.jednostka,a.month order by a.month desc;"
   } else if (quer==6) {
       var que= "select a.descr,a.mont,(max(b.poziom)*0.25+max(d.poziom)*0.25+max(f.poziom)*0.5)*iif(max(b.poziom)=0 or max(d.poziom)=0 or max(f.poziom)=0,0.5,1) as Wynik,max(b.poziom) as five_S,max(d.poziom) as wyd,max(f.poziom) as p_brak,c.wyda,a.wynik as wys,e.braki,b.jednostka,a.month from (SELECT top 2 b.descr,a.month,Format(mid(a.month,5,2) & '-01-' & mid(a.month,1,4),'mmmm yyyyr') as mont, round(a.wyn*100,2) AS Wynik FROM (SELECT b.month,mid(a.mpk,1,4) as mpk, avg(a.wynik) AS wyn FROM wynik5s AS a,(SELECT month, week FROM wydaj where month>format(now()-60,'yyyymm') GROUP BY month, week)  AS b WHERE a.tydzień=b.week GROUP BY b.month,mpk)  AS a ,(select * from dpt where dp_id=@DEP ) AS b  WHERE  a.mpk=b.mpk order by a.month desc) as a,(SELECT top 2 b.descr,a.Month,Format(a.valid_date,'mmmm yyyyr') as Moh, round((sum(a.SUM_H_MADE_CORR)/sum(a.WRK_H_SUM*7.5/8))*100,2) AS wyda FROM wydaj as a,dpt AS b WHERE a.valid_date>now()-60 and [a.END]=true and  ((b.dp_id)=@DEP) AND  iif(b.department,a.department in (select department from dept where dp_id=@DEP group by department),a.department is not null) and iif(b.workcnt,a.work_center in (select workcnt from dept where dp_id=@DEP group by workcnt),a.work_center is not null) and iif(b.notworkcnt,a.work_center not in (select notworkcnt from dept where dp_id=@DEP group by notworkcnt),a.work_center is not null) GROUP BY b.descr,a.month ,Format(a.valid_date,'mmmm yyyyr') order by a.month desc) as c,(select Poziom5s,Poziom,jednostka from progi where dp_id=@DEP) as b,(select wydaj,Poziom from progi where dp_id=@DEP) as d,(SELECT top 2 b.descr,a.month,Format(mid(a.month,5,2) & '-01-' & mid(a.month,1,4),'mmmm yyyyr') as mont, round(a.WSp_BRK,4) AS Braki FROM (SELECT month,mpk,WSp_BRK FROM MGWiS AS a),(select * from dpt where dp_id=@DEP ) AS b  WHERE  a.mpk=b.ext1 order by a.month desc) as e,(select braki,Poziom from progi where dp_id=@DEP) as f  where a.month=c.month and e.month=c.month and a.wynik>=b.poziom5S and  c.wyda>=d.wydaj and e.braki<=f.braki group by a.descr,a.mont,c.wyda,a.wynik,e.braki,b.jednostka,a.month order by a.month desc;"
   }
   const options = {
     dsn :  Adsn,
     query : que,
     prepare: "true",
     Values : {
        Val_name1: '@DEP',
        Value1:DEP,
        Type1:'VarWChar',
        Len1:5,
       }
   }
   //  console.log ("Ustawienia ACCESS")
   const  result= await db.Getoledb(options);
   if (!result.valid) {
   //    console.log ("Błąd /" + result.error + "/");
     var data = {
       error : "yes"
     };
   } else {
     //        console.log ("Dane pobrane z ACCESS")
     var colr=["rgba(215, 84, 44, 0.9)"]
     if (result.records[0].jednostka=='%') {
     if (result.records[0].Wynik>100) {
       var colr=["rgba(148, 176, 76, 0.9)"]};
     } else {
       if (result.records[0].Wynik>430) {
         var colr=["rgba(148, 176, 76, 0.9)"]}};
     if (quer==2 || quer==5)  {
       var data={
         error : "no",
         title: "WYNIK PREMII DLA DZIAŁU",
         subtitle: result.records[0].descr + " " + result.records[0].mont.toUpperCase(),
         value : result.records[0].Wynik,
         Mnt : result.records[0].mont.toUpperCase(),
         five_S : result.records[0].five_S,
         Wyd : result.records[0].wyd,
         wydajnosc : result.records[0].wyda,
         s : result.records[0].wys,
         color: colr,
         jedn: result.records[0].jednostka,
       }
     } else if (quer==1) {
       var data = {
         error : "no",
         title: "WYNIK PREMII DLA DZIAŁU",
         subtitle: result.records[0].descr + " " + result.records[0].moh.toUpperCase(),
         value : result.records[0].Wynik,
         Mnt : result.records[0].moh.toUpperCase(),
         Wyd : result.records[0].Wynik,
         wydajnosc : result.records[0].wyda,
         color: colr,
         jedn: result.records[0].jednostka,
       }
     } else if (quer==3 || quer==6) {
       var data={
         error : "no",
         title: "WYNIK PREMII DLA DZIAŁU",
         subtitle: result.records[0].descr + " " + result.records[0].mont.toUpperCase(),
         value : result.records[0].Wynik,
         Mnt : result.records[0].mont.toUpperCase(),
         five_S : result.records[0].five_S,
         Wyd : result.records[0].wyd,
         wydajnosc : result.records[0].wyda,
         s : result.records[0].wys,
         p_BRAK : result.records[0].p_brak,
         braki: result.records[0].braki,
         color: colr,
         jedn: result.records[0].jednostka,
       }
     } else if (quer==4) {
       var data = {
         error : "no",
         title: "WYNIK PREMII DLA DZIAŁU",
         subtitle: result.records[0].descr + " " + result.records[0].mont.toUpperCase(),
         value : result.records[0].Wynik,
         Mnt : result.records[0].mont.toUpperCase(),
         five_S : result.records[0].five_S,
         Wyd : result.records[0].wyd,
         wydajnosc : result.records[0].wyda,
         s : result.records[0].wys,
         p_BRAK : result.records[0].p_brak,
         braki: result.records[0].braki,
         p_wyd_MB:result.records[0].p_wyd_MB,
         wyd_MB:result.records[0].wyd_MB,
         color: colr,
         jedn: result.records[0].jednostka,
       }
     };
   };
   return data
}
async function bar_OnTime(DEP) {
  const quer=await IS_one(DEP);
  const options = {
     dsn :  Adsn,
     query : "SELECT top 1 dpt.dp_ID,dpt.dester,Termin_dzienna.MOnth as mont,Format(Termin_dzienna.day,'mmmm yyyyr') as Moh, Termin_dzienna.WYDZ, Sum(Termin_dzienna.IL_PLAN) AS SumaOfIL_PLAN, Sum(Termin_dzienna.MADE_ON_TIME) AS SumaOfMADE_ON_TIME, Sum(Termin_dzienna.MADE_TOO_LATE) AS SumaOfMADE_TOO_LATE, round(Sum([MADE_ON_TIME])/Sum([IL_PLAN])*100,2) AS WYNIK_TERm FROM Termin_dzienna,dpt where Termin_dzienna.MOnth>=Format(now()-365,'yyyymm') and dpt.dpter= Termin_dzienna.WYDZ and dpt.dp_id=@DEP GROUP BY dpt.dp_ID,dpt.dester,Termin_dzienna.MOnth, Format(Termin_dzienna.day,'mmmm yyyyr'),Termin_dzienna.WYDZ order by month desc",
     prepare: "true",
     Values : {
        Val_name1: '@DEP',
        Value1:DEP,
        Type1:'VarWChar',
        Len1:6,
       }
    }
    //   console.log ("Ustawienia ACCESS")
   const  result= await db.Getoledb(options);
   if (!result.valid || quer==1) {
    //     console.log ("Błąd /" + result.error + "/");
      var data={
        error : "yes"
      };
      chk=false;
    } else {
      //       console.log ("Dane pobrane z ACCESS")
      var data = {
       error : "no",
       title: "PRODUKCJA NA CZAS",
       subtitle: result.records[0].dester + " " + result.records[0].Moh.toUpperCase(),
       value : result.records[0].WYNIK_TERm,
       Mnt : result.records[0].Moh.toUpperCase(),
       color: ["rgba(180,152,197,0.5)"],
       jedn: '%',
     }};
     return data
}
async function historical_score(DEP,days=365) {
  const quer=await IS_one(DEP);
    if (quer==2) {
       var que= "SELECT a.descr, a.mont, (max(b.poziom)*0.25+ max(d.poziom)*0.75)*iif(max(b.poziom)=0 or max(d.poziom)=0,0.5,1) AS Wynik, max(b.poziom) AS five_S, max(d.poziom) AS wyd, c.wyda, a.wynik AS wys, b.jednostka, a.month FROM (SELECT top " + (Math.round(days/30+1)) + " b.descr,a.month,Format(mid(a.month,5,2) & '-01-' & mid(a.month,1,4),'mmmm yyyyr') as mont, round(a.wyn*100,2) AS Wynik FROM (SELECT b.month,mid(a.mpk,1,4) as mpk, avg(a.wynik) AS wyn FROM wynik5s AS a, (SELECT month, week FROM wydaj where month>format(now()-" + days + ",'yyyymm') GROUP BY month, week)  AS b WHERE a.tydzień=b.week GROUP BY b.month,mpk)  AS a ,(select * from dpt where dp_id=@DEP ) AS b  WHERE  a.mpk=b.mpk order by a.month desc)  AS a, (SELECT top " + (Math.round(days/30+1)) + " b.descr,a.Month,Format(a.valid_date,'mmmm yyyyr') as Moh, round((sum(a.SUM_H_MADE_CORR)/sum(a.WRK_H_SUM*7.5/8))*100,2) AS wyda FROM wydaj as a, dpt AS b WHERE a.valid_date>now()-" + days + " and [a.END]=true and  ((b.dp_id)=@DEP) AND iif(b.department,a.department in (select department from dept where dp_id=@DEP group by department),a.department is not null) and iif(b.workcnt,a.work_center in (select workcnt from dept where dp_id=@DEP group by workcnt),a.work_center is not null) and iif(b.notworkcnt,a.work_center not in (select notworkcnt from dept where dp_id=@DEP group by notworkcnt),a.work_center is not null) GROUP BY b.descr,a.month ,Format(a.valid_date,'mmmm yyyyr') order by a.month desc)  AS c, (select Poziom5s,Poziom,jednostka,ważny_od,ważny_do from progi where dp_id=@DEP)  AS b, (select wydaj,Poziom,ważny_od,ważny_do from progi where dp_id=@DEP)  AS d WHERE a.month=c.month and (a.wynik>=b.poziom5S and (dateserial(left(a.month,4),mid(a.month,5,2),1) between  b.ważny_od and b.ważny_do) )  and  (c.wyda>=d.wydaj  and (dateserial(left(a.month,4),mid(a.month,5,2),1) between  d.ważny_od and d.ważny_do)) GROUP BY a.descr, a.mont, c.wyda, a.wynik, b.jednostka, a.month ORDER BY a.month asc;"
    } else if (quer==1) {
        var que= "select c.descr,c.moh,max(d.poziom) as Wynik,c.wyda,d.jednostka,c.month from (SELECT top " + (Math.round(days/30+1)) + " b.descr,a.Month,Format(a.valid_date,'mmmm yyyyr') as Moh, round((sum(a.SUM_H_MADE_CORR)/sum(a.WRK_H_SUM*7.5/8))*100,2) AS wyda FROM wydaj as a,dpt AS b WHERE a.valid_date>now()-" + days + " and [a.END]=true and  ((b.dp_id)=@DEP) AND iif(b.department,a.department in (select department from dept where dp_id=@DEP group by department),a.department is not null) and iif(b.workcnt,a.work_center in (select workcnt from dept where dp_id=@DEP group by workcnt),a.work_center is not null) and iif(b.notworkcnt,a.work_center not in (select notworkcnt from dept where dp_id=@DEP group by notworkcnt),a.work_center is not null) GROUP BY b.descr,a.month ,Format(a.valid_date,'mmmm yyyyr') order by a.month desc) as c,(select wydaj,Poziom,jednostka,ważny_od,ważny_do from progi where dp_id=@DEP) as d where  c.wyda>=d.wydaj and dateserial(left(c.month,4),mid(c.month,5,2),1) between d.ważny_od and d.ważny_do group by  c.descr,c.moh,c.wyda,d.jednostka,c.month order by c.month;"
    } else if (quer==3) {
        var que="select a.descr,a.mont,(max(b.poziom)*0.25+max(d.poziom)*0.6+max(f.poziom)*0.15)*iif(max(b.poziom)=0 or max(d.poziom)=0 or max(f.poziom)=0,0.5,1) as Wynik,max(b.poziom) as five_S,max(d.poziom) as wyd,max(f.poziom) as p_brak,c.wyda,a.wynik as wys,e.braki,b.jednostka,a.month from (SELECT top " + (Math.round(days/30+1)) + " b.descr,a.month,Format(mid(a.month,5,2) & '-01-' & mid(a.month,1,4),'mmmm yyyyr') as mont, round(a.wyn*100,2) AS Wynik FROM (SELECT b.month,mid(a.mpk,1,4) as mpk, avg(a.wynik) AS wyn FROM wynik5s AS a,(SELECT month, week FROM wydaj where month>format(now()-" + days + ",'yyyymm') GROUP BY month, week)  AS b WHERE a.tydzień=b.week GROUP BY b.month,mpk)  AS a ,(select * from dpt where dp_id=@DEP ) AS b  WHERE  a.mpk=b.mpk order by a.month desc) as a,(SELECT top " + (Math.round(days/30+1)) + " b.descr,a.Month,Format(a.valid_date,'mmmm yyyyr') as Moh, round((sum(a.SUM_H_MADE_CORR)/sum(a.WRK_H_SUM*7.5/8))*100,2) AS wyda FROM wydaj as a,dpt AS b WHERE a.valid_date>now()-" + days + " and [a.END]=true and  ((b.dp_id)=@DEP) AND iif(b.department,a.department in (select department from dept where dp_id=@DEP group by department),a.department is not null) and iif(b.workcnt,a.work_center in (select workcnt from dept where dp_id=@DEP group by workcnt),a.work_center is not null) and iif(b.notworkcnt,a.work_center not in (select notworkcnt from dept where dp_id=@DEP group by notworkcnt),a.work_center is not null) GROUP BY b.descr,a.month ,Format(a.valid_date,'mmmm yyyyr') order by a.month desc) as c,(select Poziom5s,Poziom,jednostka from progi where dp_id=@DEP) as b,(select wydaj,Poziom from progi where dp_id=@DEP) as d,(SELECT top " + (Math.round(days/30+1)) + " b.descr,a.month,Format(mid(a.month,5,2) & '-01-' & mid(a.month,1,4),'mmmm yyyyr') as mont, round(a.WSp_BRK,4) AS Braki FROM (SELECT month,mpk,WSp_BRK FROM MGWiS AS a),(select * from dpt where dp_id=@DEP ) AS b  WHERE  a.mpk=b.mpk order by a.month desc) as e,(select braki,Poziom from progi where dp_id=@DEP) as f  where a.month=c.month and e.month=c.month and a.wynik>=b.poziom5S and  c.wyda>=d.wydaj and e.braki<=f.braki group by a.descr,a.mont,c.wyda,a.wynik,e.braki,b.jednostka,a.month order by a.month asc;"
    } else if (quer==4) {
        var que= "select a.descr,a.mont,(max(b.poziom)*0.15+max(d.poziom)*0.1+max(f.poziom)*0.5+max(h.poziom)*0.25)*iif(max(b.poziom)=0 or max(d.poziom)=0 or max(f.poziom)=0 or max(h.poziom)=0,0.5,1) as Wynik,max(b.poziom) as five_S,max(d.poziom) as wyd,max(f.poziom) as p_brak,max(h.poziom) as p_wyd_MB,c.wyda,a.wynik as wys,e.braki,g.wyda_kroj as wyd_MB,b.jednostka,a.month from (SELECT top " + (Math.round(Number(days)/30+1)) + " b.descr,a.month,Format(mid(a.month,5,2) & '-01-' & mid(a.month,1,4),'mmmm yyyyr') as mont, round(a.wyn*100,2) AS Wynik FROM (SELECT b.month,mid(a.mpk,1,4) as mpk, avg(a.wynik) AS wyn FROM wynik5s AS a,(SELECT month, week FROM wydaj where month>format(now()-" + Number(days) + ",'yyyymm') GROUP BY month, week)  AS b WHERE a.tydzień=b.week GROUP BY b.month,mpk)  AS a ,(select * from dpt where dp_id=@DEP ) AS b  WHERE  a.mpk=b.mpk order by a.month desc) as a,(SELECT top " + (Math.round(Number(days)/30+1)) + " b.descr,a.Month,Format(a.valid_date,'mmmm yyyyr') as Moh, round((sum(a.SUM_H_MADE_CORR)/sum(a.WRK_H_SUM*7.5/8))*100,2) AS wyda FROM wydaj as a,dpt AS b WHERE a.valid_date>now()-" + Number(days) + " and [a.END]=true and  ((b.dp_id)=@DEP) AND iif(b.department,a.department in (select department from dept where dp_id=@DEP group by department),a.department is not null) and iif(b.workcnt,a.work_center in (select workcnt from dept where dp_id=@DEP group by workcnt),a.work_center is not null) and iif(b.notworkcnt,a.work_center not in (select notworkcnt from dept where dp_id=@DEP group by notworkcnt),a.work_center is not null) GROUP BY b.descr,a.month ,Format(a.valid_date,'mmmm yyyyr') order by a.month desc) as c,(select Poziom5s,Poziom,jednostka from progi where dp_id=@DEP) as b,(select wydaj,Poziom from progi where dp_id=@DEP) as d,(SELECT top " + (Math.round(days/30+1)) + " b.descr,a.month,Format(mid(a.month,5,2) & '-01-' & mid(a.month,1,4),'mmmm yyyyr') as mont, round(a.WSp_BRK*100,4) AS Braki FROM (SELECT month,mpk,WSp_BRK FROM MGWiS AS a),(select * from dpt where dp_id=@DEP ) AS b  WHERE  a.mpk=b.ext1 order by a.month desc) as e,(select braki,Poziom from progi where dp_id=@DEP) as f,(SELECT top " + (Math.round(days/30+1)) + " b.descr,a.Month,Format(a.valid_date,'mmmm yyyyr') as Moh, round((sum(a.SUM_H_MADE_CORR)/sum(a.WRK_H_SUM)),2) AS wyda_kroj FROM wydaj as a,dpt AS b WHERE a.valid_date>now()-" + days + " and [a.END]=true and  ((b.dp_id)=@DEP) AND a.work_center=b.ext GROUP BY b.descr,a.month ,Format(a.valid_date,'mmmm yyyyr') order by a.month desc) as g,(select Wydajn_next_lev as nxt_lev,Poziom from progi where dp_id=@DEP) as h where a.month=c.month and g.Month=c.month and e.month=c.month and a.wynik>=b.poziom5S and  c.wyda>=d.wydaj and e.braki<=f.braki and g.wyda_kroj>=h.nxt_lev group by a.descr,a.mont,c.wyda,a.wynik,e.braki,g.wyda_kroj,b.jednostka,a.month order by a.month asc;"
    } else if (quer==5) {
        var que = "select a.descr,a.mont,(max(b.poziom)*0.3+max(d.poziom)*0.7)*iif(max(b.poziom)=0 or max(d.poziom)=0,0.5,1) as Wynik,max(b.poziom) as five_S,max(d.poziom) as wyd,c.wyda,a.wynik as wys,b.jednostka,a.month from (SELECT top " + (Math.round(Number(days)/30+1)) + " b.descr,a.month,Format(mid(a.month,5,2) & '-01-' & mid(a.month,1,4),'mmmm yyyyr') as mont, round(a.wyn*100,2) AS Wynik FROM (SELECT b.month,mid(a.mpk,1,4) as mpk, avg(a.wynik) AS wyn FROM wynik5s AS a,(SELECT month, week FROM wydaj where month>format(now()-" + Number(days) + ",'yyyymm') GROUP BY month, week)  AS b WHERE a.tydzień=b.week GROUP BY b.month,mpk)  AS a ,(select * from dpt where dp_id=@DEP ) AS b  WHERE  a.mpk=b.mpk order by a.month desc) as a,(SELECT top " + (Math.round(Number(days)/30+1)) + " b.descr,a.Month,Format(DateValue(right(a.Month,2) & '/' & '/01/' & Left(a.Month,4)),'mmmm yyyyr') as Moh,round(wynik_term*100,2) AS wyda FROM TERMINOWOŚ_month as a,dpt AS b WHERE a.Month>format(now()-" + Number(days) + ",'yyyymm') and  ((b.dp_id)=@DEP) AND a.wydz=b.ext order by a.month desc) as c,(select Poziom5s,Poziom,jednostka,ważny_od,ważny_do from progi where dp_id=@DEP) as b,(select wydaj,Poziom,ważny_od,ważny_do from progi where dp_id=@DEP) as d where a.month=c.month and (a.wynik>=b.poziom5S and (dateserial(left(a.month,4),mid(a.month,5,2),1) between  b.ważny_od and b.ważny_do) ) and  (c.wyda>=d.wydaj  and (dateserial(left(a.month,4),mid(a.month,5,2),1) between  d.ważny_od and d.ważny_do)) group by a.descr,a.mont,c.wyda,a.wynik,b.jednostka,a.month order by a.month asc;"
    } else if (quer==6) {
        var que= "select a.descr,a.mont,(max(b.poziom)*0.25+max(d.poziom)*0.25+max(f.poziom)*0.5)*iif(max(b.poziom)=0 or max(d.poziom)=0 or max(f.poziom)=0,0.5,1) as Wynik,max(b.poziom) as five_S,max(d.poziom) as wyd,max(f.poziom) as p_brak,c.wyda,a.wynik as wys,e.braki,b.jednostka,a.month from (SELECT top " + (Math.round(Number(days)/30+1)) + " b.descr,a.month,Format(mid(a.month,5,2) & '-01-' & mid(a.month,1,4),'mmmm yyyyr') as mont, round(a.wyn*100,2) AS Wynik FROM (SELECT b.month,mid(a.mpk,1,4) as mpk, avg(a.wynik) AS wyn FROM wynik5s AS a,(SELECT month, week FROM wydaj where month>format(now()-" + Number(days) + ",'yyyymm') GROUP BY month, week)  AS b WHERE a.tydzień=b.week GROUP BY b.month,mpk)  AS a ,(select * from dpt where dp_id=@DEP ) AS b  WHERE  a.mpk=b.mpk order by a.month desc) as a,(SELECT top " + (Math.round(Number(days)/30+1)) + " b.descr,a.Month,Format(a.valid_date,'mmmm yyyyr') as Moh, round((sum(a.SUM_H_MADE_CORR)/sum(a.WRK_H_SUM*7.5/8))*100,2) AS wyda FROM wydaj as a,dpt AS b WHERE a.valid_date>now()-" + days + " and [a.END]=true and  ((b.dp_id)=@DEP) AND iif(b.department,a.department in (select department from dept where dp_id=@DEP group by department),a.department is not null) and iif(b.workcnt,a.work_center in (select workcnt from dept where dp_id=@DEP group by workcnt),a.work_center is not null) and iif(b.notworkcnt,a.work_center not in (select notworkcnt from dept where dp_id=@DEP group by notworkcnt),a.work_center is not null) GROUP BY b.descr,a.month ,Format(a.valid_date,'mmmm yyyyr') order by a.month desc) as c,(select Poziom5s,Poziom,jednostka from progi where dp_id=@DEP) as b,(select wydaj,Poziom from progi where dp_id=@DEP) as d,(SELECT top " + (Math.round(days/30+1)) + " b.descr,a.month,Format(mid(a.month,5,2) & '-01-' & mid(a.month,1,4),'mmmm yyyyr') as mont, round(a.WSp_BRK,4) AS Braki FROM (SELECT month,mpk,WSp_BRK FROM MGWiS AS a),(select * from dpt where dp_id=@DEP ) AS b  WHERE  a.mpk=b.ext1 order by a.month desc) as e,(select braki,Poziom from progi where dp_id=@DEP) as f  where a.month=c.month and e.month=c.month and a.wynik>=b.poziom5S and  c.wyda>=d.wydaj and e.braki<=f.braki group by a.descr,a.mont,c.wyda,a.wynik,e.braki,b.jednostka,a.month order by a.month asc;"
    };
    const  options   = {
      dsn :  Adsn,
      query : que,
      prepare: "true",
      Values : {
         Val_name1: '@DEP',
         Value1:DEP,
         Type1:'VarWChar',
         Len1:5,
        }
    }
    //    console.log ("Ustawienia ACCESS")
    const  result= await db.Getoledb(options);
    if (!result.valid) {
      //      console.log ("Błąd /" + result.error + "/");
      var data ={
        error : "yes"
      };
    } else {
      //    console.log ("Dane pobrane z ACCESS")
      if (quer==2) {
        var data = {
          error : "no",
          default_LT:365,
          LT_unit:"dni",
          labels : result.records.map(function(it){ return it.mont}),
          jedn: result.records[0].jednostka,
          title: "WYNIKI PREMII DLA DZIAŁU "  + result.records[0].descr,
          title1: "WYDAJNOŚĆ " + result.records[0].descr,
          title2: "WYNIKI 5-S "  + result.records[0].descr,
          subtitle:"Miesięcznie",
          datasets : [{
            fillColor : "rgba(148,176,76,0.5)",
            strokeColor : "rgba(151,187,205,1)",
            pointColor : "rgba(220,220,220,1)",
            pointStrokeColor : "#fff",
            data : result.records.map(function(it){ return it.Wynik})
          }],
          datasets1 : [{
            fillColor : "rgba(255, 79, 48, 0.3)",
            strokeColor : "rgba(151,187,205,1)",
            pointColor : "rgba(220,220,220,1)",
            pointStrokeColor : "#fff",
            data : result.records.map(function(it){ return it.wyda})
            }],
          datasets2 : [{
            fillColor : "rgba(151,187,205,0.5)",
            strokeColor : "rgba(151,187,205,1)",
            pointColor : "rgba(220,220,220,1)",
            pointStrokeColor : "#fff",
            data : result.records.map(function(it){ return it.wys})
          }]
        }
        //const json2csvParser = new Json2csvParser();
        //const csv = json2csvParser.parse(result.records);
        //fs.writeFile('d:/' + DEP +'.csv', csv.replace('.', ','), (err) => {
        //  if (err) throw err;
        //  console.log('The file has been saved!');
        //});
      } else if (quer==1) {
        var data = {
          error : "no",
          default_LT:365,
          LT_unit:"dni",
          labels : result.records.map(function(it){ return it.moh}),
          jedn: result.records[0].jednostka,
          title: "WYNIKI PREMII DLA DZIAŁU "  + result.records[0].descr,
          title1: "WYDAJNOŚĆ " + result.records[0].descr,
          subtitle:"Miesięcznie",
          datasets : [{
           fillColor : "rgba(148,176,76,0.5)",
           strokeColor : "rgba(151,187,205,1)",
           pointColor : "rgba(220,220,220,1)",
           pointStrokeColor : "#fff",
           data : result.records.map(function(it){ return it.Wynik})
          }],
          datasets1 : [{
           fillColor : "rgba(255, 79, 48, 0.3)",
           strokeColor : "rgba(151,187,205,1)",
           pointColor : "rgba(220,220,220,1)",
           pointStrokeColor : "#fff",
           data : result.records.map(function(it){ return it.wyda})
          }]
        }
        //const json2csvParser = new Json2csvParser();
        //const csv = json2csvParser.parse(result.records);
        //fs.writeFile('d:/' + DEP +'.csv', csv.replace('.', ','), (err) => {
        //  if (err) throw err;
        //  console.log('The file has been saved!');
        //});
      } else if (quer==3) {
        var data = {
          error : "no",
          default_LT:365,
          LT_unit:"dni",
          labels : result.records.map(function(it){ return it.mont}),
          jedn: result.records[0].jednostka,
          title: "WYNIKI PREMII DLA DZIAŁU "  + result.records[0].descr,
          title1: "WYDAJNOŚĆ " + result.records[0].descr,
          title2: "WYNIKI 5-S "  + result.records[0].descr,
          title3: "BRAKI /REKLAMACJE "  + result.records[0].descr,
          subtitle:"Miesięcznie",
          datasets : [{
           fillColor : "rgba(148,176,76,0.5)",
           strokeColor : "rgba(151,187,205,1)",
           pointColor : "rgba(220,220,220,1)",
           pointStrokeColor : "#fff",
           data : result.records.map(function(it){ return it.Wynik})
            }],
          datasets1 : [{
            fillColor : "rgba(255, 79, 48, 0.3)",
            strokeColor : "rgba(151,187,205,1)",
            pointColor : "rgba(220,220,220,1)",
            pointStrokeColor : "#fff",
            data : result.records.map(function(it){ return it.wyda})
            }],
          datasets2 : [{
           fillColor : "rgba(151,187,205,0.5)",
           strokeColor : "rgba(151,187,205,1)",
           pointColor : "rgba(220,220,220,1)",
           pointStrokeColor : "#fff",
           data : result.records.map(function(it){ return it.wys})
            }],
          datasets3 : [{
           fillColor : "rgba(194,152,44,0.5)",
           strokeColor : "rgba(151,187,205,1)",
           pointColor : "rgba(220,220,220,1)",
           pointStrokeColor : "#fff",
           data : result.records.map(function(it){ return it.braki})
          }]
        }
        //const json2csvParser = new Json2csvParser();
        //const csv = json2csvParser.parse(result.records);
        //fs.writeFile('d:/' + DEP +'.csv', csv.replace('.', ','), (err) => {
        //  if (err) throw err;
        //    console.log('The file has been saved!');
        //  });
      } else if (quer==4) {
        var data= {
          error : "no",
          default_LT:365,
          LT_unit:"dni",
          labels : result.records.map(function(it){ return it.mont}),
          jedn: result.records[0].jednostka,
          title: "WYNIKI PREMII DLA DZIAŁU "  + result.records[0].descr,
          title1: "WYDAJNOŚĆ DZIAŁ NADRZĘDNY WZGLĘDEM OBSZARU " + result.records[0].descr,
          title2: "WYNIKI 5-S "  + result.records[0].descr,
          title3: "ODPADOWOŚĆ "  + result.records[0].descr,
          title4: "WYDAJNOŚĆ "  + result.records[0].descr,
          subtitle:"Miesięcznie",
          datasets : [{
            fillColor : "rgba(148,176,76,0.5)",
            strokeColor : "rgba(151,187,205,1)",
            pointColor : "rgba(220,220,220,1)",
            pointStrokeColor : "#fff",
            data : result.records.map(function(it){ return it.Wynik})
          }],
          datasets1 : [{
            fillColor : "rgba(255, 79, 48, 0.3)",
            strokeColor : "rgba(151,187,205,1)",
            pointColor : "rgba(220,220,220,1)",
            pointStrokeColor : "#fff",
            data : result.records.map(function(it){ return it.wyda})
          }],
          datasets2 : [{
            fillColor : "rgba(151,187,205,0.5)",
            strokeColor : "rgba(151,187,205,1)",
            pointColor : "rgba(220,220,220,1)",
            pointStrokeColor : "#fff",
            data : result.records.map(function(it){ return it.wys})
          }],
          datasets3 : [{
            fillColor : "rgba(45,212,195,0.4)",
            strokeColor : "rgba(151,187,205,1)",
            pointColor : "rgba(220,220,220,1)",
            pointStrokeColor : "#fff",
            data : result.records.map(function(it){ return it.braki})
          }],
          datasets4 : [{
            fillColor : "rgba(194,152,44,0.5)",
            strokeColor : "rgba(151,187,205,1)",
            pointColor : "rgba(220,220,220,1)",
            pointStrokeColor : "#fff",
            data : result.records.map(function(it){ return it.wyd_MB})
          }]
        }
        //const json2csvParser = new Json2csvParser();
        //const csv = json2csvParser.parse(result.records);
        //fs.writeFile('d:/' + DEP +'.csv',csv.replace('.', ','), (err) => {
        //  if (err) throw err;
        //  console.log('The file has been saved!');
        //});
      } else if (quer==5) {
          var data = {
            error : "no",
            default_LT:365,
            LT_unit:"dni",
            labels : result.records.map(function(it){ return it.mont}),
            jedn: result.records[0].jednostka,
            title: "WYNIKI PREMII DLA DZIAŁU "  + result.records[0].descr,
            title1: "PRODUKCJA NA CZAS TAPICERNIA",
            title2: "WYNIKI 5-S "  + result.records[0].descr,
            subtitle:"Miesięcznie",
            datasets : [{
                fillColor : "rgba(148,176,76,0.5)",
                strokeColor : "rgba(151,187,205,1)",
                pointColor : "rgba(220,220,220,1)",
                pointStrokeColor : "#fff",
                data : result.records.map(function(it){ return it.Wynik})
            }],
            datasets1 : [{
                fillColor : "rgba(180,152,197,0.5)",
                strokeColor : "rgba(151,187,205,1)",
                pointColor : "rgba(220,220,220,1)",
                pointStrokeColor : "#fff",
                data : result.records.map(function(it){ return it.wyda})
            }],
            datasets2 : [{
                fillColor : "rgba(151,187,205,0.5)",
                strokeColor : "rgba(151,187,205,1)",
                pointColor : "rgba(220,220,220,1)",
                pointStrokeColor : "#fff",
                data : result.records.map(function(it){ return it.wys})
            }]
          }
          //const json2csvParser = new Json2csvParser();
          //const csv = json2csvParser.parse(result.records);
          //fs.writeFile('d:/' + DEP +'.csv', csv.replace('.', ','), (err) => {
          //  if (err) throw err;
          //  console.log('The file has been saved!');
          //});
        } else if (quer==6) {
          var data = {
            error : "no",
            default_LT:365,
            LT_unit:"dni",
            labels : result.records.map(function(it){ return it.mont}),
            jedn: result.records[0].jednostka,
            title: "WYNIKI PREMII DLA DZIAŁU "  + result.records[0].descr,
            title1: "WYDAJNOŚĆ DZIAŁ NADRZĘDNY WZGLĘDEM OBSZARU " + result.records[0].descr,
            title2: "WYNIKI 5-S "  + result.records[0].descr,
            title3: "WSKAŹNIK ROZPOCZĘTYCH WAŁKÓW "  + result.records[0].descr,
            subtitle:"Miesięcznie",
            datasets : [{
               fillColor : "rgba(148,176,76,0.5)",
               strokeColor : "rgba(151,187,205,1)",
               pointColor : "rgba(220,220,220,1)",
               pointStrokeColor : "#fff",
               data : result.records.map(function(it){ return it.Wynik})
            }],
            datasets1 : [{
              fillColor : "rgba(255, 79, 48, 0.3)",
              strokeColor : "rgba(151,187,205,1)",
              pointColor : "rgba(220,220,220,1)",
              pointStrokeColor : "#fff",
              data : result.records.map(function(it){ return it.wyda})
            }],
            datasets2 : [{
              fillColor : "rgba(151,187,205,0.5)",
              strokeColor : "rgba(151,187,205,1)",
              pointColor : "rgba(220,220,220,1)",
              pointStrokeColor : "#fff",
              data : result.records.map(function(it){ return it.wys})
            }],
            datasets3 : [{
              fillColor : "rgba(194,152,44,0.5)",
              strokeColor : "rgba(151,187,205,1)",
              pointColor : "rgba(220,220,220,1)",
              pointStrokeColor : "#fff",
              data : result.records.map(function(it){ return it.braki})
            }]
          };
          //const json2csvParser = new Json2csvParser();
          //const csv = json2csvParser.parse(result.records);
          //fs.writeFile('d:/' + DEP +'.csv', csv.replace('.', ','), (err) => {
          //  if (err) throw err;
          //  console.log('The file has been saved!');
          //});
        };
    };

    return data;
}
module.exports = {
  Present_score: (DEP) => present_score(DEP),
  Score_barLeft: (DEP) => score_barLeft(DEP),
  Score_barRight:(DEP) => score_barRight(DEP),
  Score_barCentral:(DEP) => score_barCentral(DEP),
  Bar_OnTime:(DEP)=>bar_OnTime(DEP),
  Historical_score:(DEP,days=365) => historical_score(DEP,days)
};
