const db = require('../connect/access.js')
const container = require('../connect/Transp_obj_wydnett.js')
const Adsn="Provider=Microsoft.ace.OLEDB.12.0;Data Source=WYDAJ.accdb;Persist Security Info=False;Mode=Read;";
async function day(DEP,dayFromNow=33) {
    const que = "SELECT b.dester as descr,Format(a.valid_date,'yyyy-mm-dd') as dzien,round((sum(a.wyk)/((sum(a.norma)+sum(c.sred))*7.5/8))*100) AS wyda,round(sum(a.wyk),2) as wykon,round(sum(a.norma)+sum(c.sred),2) as godz_brutt,round((sum(a.norma)+sum(c.sred))*7.5/8,2) as godz_netto FROM dpt AS b,(select valid_date,department,sum(SUM_H_MADE_NORM) as wyk,sum(WRK_norm) as norma from wydaj where [END]=true group by valid_date,department ) as a,(Select ID,wydz,mont,week,dat,sum(gr_sred) as sred from (select * from gr_sred union select Format(valid_date,'dd-mm-yyyy') & '_'  & department as ID, department as wydz,month as mont,week,valid_date as dat,0 as gr_sred from wydaj where valid_date>now()-" + Number(dayFromNow) + ") group by  ID,wydz,mont,week,dat) as c WHERE a.valid_date>now()-" + Number(dayFromNow) + " and ((b.dp_id)=@DEP) AND iif((b.dpter='ALL' or b.dpter='MAG_WE'),a.department in (select department from dept where dp_id=@DEP group by department),a.department=b.dpter) and (c.ID=Format(a.valid_date,'dd-mm-yyyy') & '_'  & a.department) GROUP BY b.dester ,Format(a.valid_date,'yyyy-mm-dd') order by Format(a.valid_date,'yyyy-mm-dd');"
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
    const title="WYDAJNOŚĆ NETTO " + result.records[0].descr;
    var DataSet =
    [
      result.records.map(function(it){ return it.wyda}),
      result.records.map(function(it){ return it.wykon}),
      result.records.map(function(it){ return it.godz_brutt}),
      result.records.map(function(it){ return it.godz_netto})]
    var data = await container.Create(result.records.map(function(it){ return it.dzien}),title,"Dzienna",DataSet);
    data["default_LT"]=33;
    data["LT_unit"]="dni";
  }
  return data
};
async function wk(DEP,WeekFromNow=13) {
  const que = "select * from (SELECT top " + Number(WeekFromNow) + " b.dester as descr,'Tydz: ' & a.week AS week,round((sum(a.wyk)/((sum(a.norma)+sum(c.sred))*7.5/8))*100,2) AS wyda,round(sum(a.wyk),2) as wykon,round(sum(a.norma)+sum(c.sred),2) as godz_brutt,round((sum(a.norma)+sum(c.sred))*7.5/8,2) as godz_netto FROM dpt AS b,(select valid_date,week,department,sum(SUM_H_MADE_NORM) as wyk,sum(WRK_norm) as norma from wydaj where [END]=true group by valid_date,week,department ) as a,(Select ID,wydz,mont,week,dat,sum(gr_sred) as sred from (select * from gr_sred union select Format(valid_date,'dd-mm-yyyy') & '_'  & department as ID, department as wydz,month as mont,week,valid_date as dat,0 as gr_sred from wydaj where valid_date>now()-" + (Number(WeekFromNow)+3)*7 + ") group by  ID,wydz,mont,week,dat) as c WHERE a.valid_date>now()-" + (Number(WeekFromNow)+3)*7 + " and ((b.dp_id)=@DEP) AND iif((b.dpter='ALL' or b.dpter='MAG_WE'),a.department in (select department from dept where dp_id=@DEP group by department),a.department=b.dpter) and  (c.ID=Format(a.valid_date,'dd-mm-yyyy') & '_'  & a.department) GROUP BY b.dester ,'Tydz: ' & a.week order by 'Tydz: ' & a.week desc) order by week asc"
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
    const title="WYDAJNOŚĆ NETTO " + result.records[0].descr;
    const DataSet =
    [
      result.records.map(function(it){ return it.wyda}),
      result.records.map(function(it){ return it.wykon}),
      result.records.map(function(it){ return it.godz_brutt}),
      result.records.map(function(it){ return it.godz_netto})]
    var data = await container.Create(result.records.map(function(it){ return it.week}),title,"Tygodniowa",DataSet);
    data["default_LT"]=13;
    data["LT_unit"]="tygodnie";
  }
  return data
};
async function mnth(DEP,DaysFromNow=365) {
  const que = "SELECT b.dester as descr,Format(a.valid_date,'yyyymm') as Mont,Format(a.valid_date,'mmmm yyyyr') as Moh, round((sum(a.wyk)/((sum(a.norma)+sum(c.sred))*7.5/8))*100,2) AS wyda,round(sum(a.wyk),2) as wykon,round(sum(a.norma)+sum(c.sred),2) as godz_brutt,round((sum(a.norma)+sum(c.sred))*7.5/8,2) as godz_netto FROM dpt AS b,(select valid_date,department,sum(SUM_H_MADE_NORM) as wyk,sum(WRK_norm) as norma from wydaj where [END]=true group by valid_date,department ) as a,(Select ID,wydz,mont,week,dat,sum(gr_sred) as sred from (select * from gr_sred union select Format(valid_date,'dd-mm-yyyy') & '_'  & department as ID, department as wydz,month as mont,week,valid_date as dat,0 as gr_sred from wydaj where valid_date>now()-" + Number(DaysFromNow) + ") group by  ID,wydz,mont,week,dat) as c WHERE a.valid_date>now()-" + Number(DaysFromNow) + " and ((b.dp_id)=@DEP) AND iif((b.dpter='ALL' or b.dpter='MAG_WE'),a.department in (select department from dept where dp_id=@DEP  group by department),a.department=b.dpter) and (c.ID=Format(a.valid_date,'dd-mm-yyyy') & '_'  & a.department) GROUP BY b.dester ,Format(a.valid_date,'yyyymm'),Format(a.valid_date,'mmmm yyyyr') order by Format(a.valid_date,'yyyymm');"
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
    const title="WYDAJNOŚĆ NETTO " + result.records[0].descr;
    const DataSet =
    [
      result.records.map(function(it){ return it.wyda}),
      result.records.map(function(it){ return it.wykon}),
      result.records.map(function(it){ return it.godz_brutt}),
      result.records.map(function(it){ return it.godz_netto})]
    var data = await container.Create(result.records.map(function(it){ return it.Moh}),title,"Miesięczna",DataSet);
    data["default_LT"]=365;
    data["LT_unit"]="dni";
  }
  return data
};
module.exports = {
  Day: (DEP,dayFromNow=33) => day(DEP,dayFromNow),
  Wk: (DEP,WeekFromNow=13) => wk(DEP,WeekFromNow),
  Mnth:(DEP,DaysFromNow=365) => mnth(DEP,DaysFromNow)
};
