const db = require('../connect/access.js')
const container = require('../connect/Transp_obj_wyd.js')
const Adsn="Provider=Microsoft.ace.OLEDB.12.0;Data Source=WYDAJ.accdb;Persist Security Info=False;Mode=Read;";
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
async function day(DEP,dayFromNow=33) {
  const quer=await IS_one(DEP);
  var que = "SELECT b.descr,Format(a.valid_date,'yyyy-mm-dd') AS dzien, round((sum(a.SUM_H_MADE_CORR)/sum(a.WRK_H_SUM*7.5/8))*100) AS wyda,round(sum(SUM_H_MADE_CORR),2) as wykon,round(sum(WRK_H_SUM),2) as godz_brutt,round(sum(a.WRK_H_SUM)*7.5/8,2) as godz_netto FROM wydaj as a,dpt AS b WHERE [a.END]=true and a.valid_date>dateadd('d',-" + Number(dayFromNow) +",date()) and ((b.dp_id)= @DEP ) AND iif(b.department,a.department in (select department from dept where dp_id= @DEP group by department),a.department is not null) and iif(b.workcnt,a.work_center in (select workcnt from dept where dp_id = @DEP group by workcnt),a.work_center is not null) and iif(b.notworkcnt,a.work_center not in (select notworkcnt from dept where dp_id= @DEP group by notworkcnt),a.work_center is not null)  GROUP BY b.descr,Format(valid_date,'yyyy-mm-dd');"
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
    let title="WYDAJNOŚĆ " + result.records[0].descr;
    if (quer==4 || quer==6) {
      title="WYDAJNOŚĆ DZIAŁ NADRZĘDNY DLA OBSZARU " + result.records[0].descr;
    }
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
  const quer=await IS_one(DEP);
  var que = "select * from (SELECT top " + Number(WeekFromNow) + " b.descr, 'Tydz: ' & a.week AS week, Round((Sum(a.SUM_H_MADE_CORR)/Sum(a.WRK_H_SUM*7.5/8))*100,2) AS wyda,round(sum(SUM_H_MADE_CORR),2) as wykon,round(sum(WRK_H_SUM),2) as godz_brutt,round(sum(a.WRK_H_SUM)*7.5/8,2) as godz_netto FROM wydaj AS a, dpt AS b WHERE [a.END]=True AND a.week>Format(Now()-" + (Number(WeekFromNow)+3)*7 + ",'yyyy') & IIf(Len(Format(Now()-"+ (Number(WeekFromNow)+3)*7 + ",'ww'))=1,'0' & Format(Now()-"+ (Number(WeekFromNow)+3)*7 + ",'ww'),Format(Now()-" + (Number(WeekFromNow)+3)*7 + ",'ww')) AND ((b.dp_id)=@DEP) AND iif(b.department,a.department in (select department from dept where dp_id=@DEP group by department),a.department is not null) and iif(b.workcnt,a.work_center in (select workcnt from dept where dp_id=@DEP group by workcnt),a.work_center is not null) and iif(b.notworkcnt,a.work_center not in (select notworkcnt from dept where dp_id=@DEP group by notworkcnt),a.work_center is not null)  GROUP BY b.descr, a.week order by a.week desc ) order by week;"
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
    let title="WYDAJNOŚĆ " + result.records[0].descr;
    if (quer==4 || quer==6) {
      title="WYDAJNOŚĆ DZIAŁ NADRZĘDNY DLA OBSZARU " + result.records[0].descr;
    }
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
  const quer=await IS_one(DEP);
  var que = "SELECT b.descr,a.Month,Format(a.valid_date,'mmmm yyyyr') as Moh, round((sum(a.SUM_H_MADE_CORR)/sum(a.WRK_H_SUM*7.5/8))*100,2) AS wyda,round(sum(a.SUM_H_MADE_CORR),2) as wykon,round(sum(a.WRK_H_SUM),2) as godz_brutt,round(sum(a.WRK_H_SUM)*7.5/8,2) as godz_netto FROM wydaj as a, dpt AS b WHERE [a.END]=true and a.Month>=Format(now()-" + Number(DaysFromNow) + ",'yyyymm') and  ((b.dp_id)=@DEP) AND iif(b.department,a.department in (select department from dept where dp_id=@DEP group by department),a.department is not null) and iif(b.workcnt,a.work_center in (select workcnt from dept where dp_id=@DEP  group by workcnt),a.work_center is not null) and iif(b.notworkcnt,a.work_center not in (select notworkcnt from dept where dp_id=@DEP group by notworkcnt),a.work_center is not null) GROUP BY b.descr,a.month ,Format(a.valid_date,'mmmm yyyyr') order by a.month asc;"
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
    let title="WYDAJNOŚĆ " + result.records[0].descr;
    if (quer==4 || quer==6) {
      title="WYDAJNOŚĆ DZIAŁ NADRZĘDNY DLA OBSZARU " + result.records[0].descr;
    }
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
