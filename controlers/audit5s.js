const db = require('../connect/access.js')
const container = require('../connect/Transp_obj_5s.js')
const Adsn="Provider=Microsoft.ace.OLEDB.12.0;Data Source=WYDAJ.accdb;Persist Security Info=False;Mode=Read;";
async function day(DEP,dayFromNow=33) {
  const que = "SELECT b.descr, Format(a.valid_date,'yyyy-mm-dd') as dzien, round(a.wyn*100,1) AS Wynik FROM (SELECT b.valid_date,mid(a.mpk,1,4) as mpk, avg(a.wynik) AS wyn FROM wynik5s AS a,(SELECT valid_date, week FROM wydaj where valid_date>=(now()-" + Number(dayFromNow) + ") GROUP BY valid_date, week) AS b WHERE a.tydzień=b.week GROUP BY b.valid_date,mpk)  AS a ,(select * from dpt where dp_id=@DEP ) AS b WHERE  a.mpk=b.mpk order by Format(a.valid_date,'yyyy-mm-dd');"
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
  const  result  = await db.Getoledb(options);
  if (!result.valid || result.records.length==0) {
   //      console.log ("Błąd /" + result.error + "/");
    var data  ={
      error : "yes"
      };
  } else {
    const title="WYNIKI 5-S " + result.records[0].descr;
    const DataSet =
    [
      result.records.map(function(it){ return it.Wynik})
    ]
    var data = await container.Create(result.records.map(function(it){ return it.dzien}),title,"Dzienna",DataSet);
    data["default_LT"]=33;
    data["LT_unit"]="dni";
  }
  return data
};
async function wk(DEP,WeekFromNow=13) {
  const que = "select * from (SELECT top " + Number(WeekFromNow) + " b.descr, 'Tydz: ' & a.tydzień as week, round(a.wyn*100,2) AS Wynik FROM (SELECT tydzień, mid(mpk,1,4) as mpk1, avg(wynik) AS wyn FROM wynik5s where tydzień>Format(Now()-" + (Number(WeekFromNow)+3)*7 + ",'yyyy') & IIf(Len(Format(Now()-" + (Number(WeekFromNow)+3)*7 + ",'ww'))=1,'0' & Format(Now()-180,'ww'),Format(Now()-" + (Number(WeekFromNow)+3)*7 + ",'ww')) GROUP BY tydzień, mpk)  AS a ,(select * from dpt where dp_id=@DEP ) AS b WHERE a.mpk1=b.mpk order by a.tydzień desc) order by week;"
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
  const  result  = await db.Getoledb(options);
  if (!result.valid || result.records.length==0) {
   //      console.log ("Błąd /" + result.error + "/");
    var data  ={
      error : "yes"
      };
  } else {
    const title="WYNIKI 5-S " + result.records[0].descr;
    const DataSet =
    [
      result.records.map(function(it){ return it.Wynik})
    ]
    var data = await container.Create(result.records.map(function(it){ return it.week}),title,"Tygodniowa",DataSet);
    data["default_LT"]=13;
    data["LT_unit"]="tygodnie";
  }
  return data
};
async function mnth(DEP,DaysFromNow=365) {
  const que = "SELECT b.descr, Format(mid(a.month,5,2) & '-01-' & mid(a.month,1,4),'mmmm yyyyr') as mont, round(a.wyn*100,2) AS Wynik FROM (SELECT b.month,mid(a.mpk,1,4) as mpk, avg(a.wynik) AS wyn FROM wynik5s AS a, (SELECT month, week FROM wydaj where month>=Format(now()-" + Number(DaysFromNow) + ",'yyyymm') GROUP BY month, week)  AS b WHERE a.tydzień=b.week GROUP BY b.month,mpk)  AS a ,(select * from dpt where dp_id=@DEP ) AS b WHERE  a.mpk=b.mpk order by a.month;"
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
  const  result  = await db.Getoledb(options);
  if (!result.valid || result.records.length==0) {
   //      console.log ("Błąd /" + result.error + "/");
    var data  ={
      error : "yes"
      };
  } else {
    const title="WYNIKI 5-S " + result.records[0].descr;
    const DataSet =
    [
      result.records.map(function(it){ return it.Wynik})
    ]
    var data = await container.Create(result.records.map(function(it){ return it.mont}),title,"Miesięczna",DataSet);
    data["default_LT"]=365;
    data["LT_unit"]="dni";
  }
  return data
};
module.exports = {
  Day: (DEP,dayFromNow) => day(DEP,dayFromNow),
  Wk: (DEP,WeekFromNow) => wk(DEP,WeekFromNow),
  Mnth:(DEP,DaysFromNow) => mnth(DEP,DaysFromNow)
};
