const db = require('../connect/access.js')
const container = require('../connect/Transp_obj_termin.js')
const Adsn="Provider=Microsoft.ace.OLEDB.12.0;Data Source=WYDAJ.accdb;Persist Security Info=False;Mode=Read;";
async function day(DEP,dayFromNow=33) {
  const que = "SELECT dpt.dp_ID,dpt.dester,r.WYDZ & r.DAY AS ID, r.MOnth, r.WEEk, format(r.DAY,'yyyy-mm-dd') as dzie, r.WYDZ, r.REFRDAT, r.IL_PLAN,r.MADE_ON_TIME, r.MADE_TOO_LATE, round(r.MADE_ON_TIME/r.IL_PLAN*100) AS TERMIN FROM Termin_dzienna as r,dpt WHERE r.day>now()-" + Number(dayFromNow) + " and dpt.dpter=r.WYDZ and dpt.dp_id=@DEP order by r.DAY asc ;"
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
  const result = await db.Getoledb(options);
  if (!result.valid || result.records.length==0) {
    var data  ={
      error : "yes"
      };
  } else {
    const DataSet =
    [
      result.records.map(function(it){ return it.TERMIN}),
      result.records.map(function(it){ return it.IL_PLAN}),
      result.records.map(function(it){ return it.MADE_ON_TIME}),
      result.records.map(function(it){ return it.MADE_TOO_LATE})]
    var data= await container.Create(result.records.map(function(it){ return it.dzie}),"PROD. NA CZAS "  + result.records[0].dester,"Dzienna",DataSet);
    data["default_LT"]=33;
    data["LT_unit"]="dni";
  }
  return data;
}
async function wk(DEP,WeekFromNow=13) {
  const que = "select * from (SELECT top " + Number(WeekFromNow) + " dpt.dp_ID,dpt.dester, 'Tydz: ' & Termin_dzienna.WEEk AS WEEk, Termin_dzienna.WYDZ, Sum(Termin_dzienna.IL_PLAN) AS SumaOfIL_PLAN, Sum(Termin_dzienna.MADE_ON_TIME) AS SumaOfMADE_ON_TIME, Sum(Termin_dzienna.MADE_TOO_LATE) AS SumaOfMADE_TOO_LATE, round(Sum([MADE_ON_TIME])/Sum([IL_PLAN])*100,2) AS WYNIK_TERm FROM Termin_dzienna,dpt where dpt.dpter=Termin_dzienna.WYDZ and dpt.dp_id=@DEP GROUP BY dpt.dp_ID,dpt.dester,Termin_dzienna.WEEk, Termin_dzienna.WYDZ order by Termin_dzienna.WEEk desc )  order by week asc"
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
  const result = await db.Getoledb(options);
  if (!result.valid || result.records.length==0) {
    var data  ={
      error : "yes"
      };
  } else {
    const DataSet =
    [
      result.records.map(function(it){ return it.WYNIK_TERm}),
      result.records.map(function(it){ return it.SumaOfIL_PLAN}),
      result.records.map(function(it){ return it.SumaOfMADE_ON_TIME}),
      result.records.map(function(it){ return it.SumaOfMADE_TOO_LATE})]
    var data= await container.Create(result.records.map(function(it){ return it.WEEk}),"PROD. NA CZAS "  + result.records[0].dester,"Tygodniowa",DataSet);
    data["default_LT"]=13;
    data["LT_unit"]="tygodne";
  }
  return data;
};
async function mnth(DEP,DaysFromNow=365) {
  const que = "SELECT dpt.dp_ID,dpt.dester,Termin_dzienna.MOnth,Format(Termin_dzienna.day,'mmmm yyyyr') as Moh, Termin_dzienna.WYDZ, Sum(Termin_dzienna.IL_PLAN) AS SumaOfIL_PLAN, Sum(Termin_dzienna.MADE_ON_TIME) AS SumaOfMADE_ON_TIME, Sum(Termin_dzienna.MADE_TOO_LATE) AS SumaOfMADE_TOO_LATE, round(Sum([MADE_ON_TIME])/Sum([IL_PLAN])*100,2) AS WYNIK_TERm FROM Termin_dzienna,dpt where Termin_dzienna.MOnth>=Format(now()-" + Number(DaysFromNow) + ",'yyyymm') and dpt.dpter= Termin_dzienna.WYDZ and dpt.dp_id=@DEP GROUP BY dpt.dp_ID,dpt.dester,Termin_dzienna.MOnth, Format(Termin_dzienna.day,'mmmm yyyyr'),Termin_dzienna.WYDZ order by month asc;"
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
  const result = await db.Getoledb(options);
  if (!result.valid || result.records.length==0) {
    var data  ={
      error : "yes"
      };
  } else {
    const DataSet =
    [
      result.records.map(function(it){ return it.WYNIK_TERm}),
      result.records.map(function(it){ return it.SumaOfIL_PLAN}),
      result.records.map(function(it){ return it.SumaOfMADE_ON_TIME}),
      result.records.map(function(it){ return it.SumaOfMADE_TOO_LATE})]
    var data= await container.Create(result.records.map(function(it){ return it.Moh}),"PROD. NA CZAS "  + result.records[0].dester,"Miesięczna",DataSet);
    data["default_LT"]=365;
    data["LT_unit"]="dni";
  }
  return data;
};
module.exports = {
  Day: (DEP,dayFromNow=33) => day(DEP,dayFromNow),
  Wk: (DEP,WeekFromNow=13) => wk(DEP,WeekFromNow),
  Mnth: (DEP,DaysFromNow=365) => mnth(DEP,DaysFromNow)
};
