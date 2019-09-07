Connect1="Provider=OraOLEDB.Oracle.1;Password=pass;Persist Security Info=True;User ID=radkos;Data Source=prod8;"
Set objConnection1 = CreateObject("ADODB.Connection")
Set rs2 = CreateObject("ADODB.Recordset")
objConnection1.Open (connect1)
rs2.CursorLocation = 3
rs2.LockType=3
Set rs2.ActiveConnection = objConnection1
rs2.open "select value from nls_database_parameters where parameter='NLS_CHARACTERSET'"
Wscript.Echo rs2(0)
rs2.close
rs2.open "SELECT to_char(VALID_DATE,'DD-MM-IYYY')||'_'||WRK_Cent iD_WRK,To_Char(VALID_DATE,'IYYYMM') Month,To_Char(VALID_DATE,'IYYYIW') WEEK,VALID_DATE,ifsapp.work_center_api.Get_Department_No('ST',WRK_Cent) Department,WRK_Cent Work_center,sum(SUM_H_MADE_corr+decode(WRK_Cent,'220SE',decode(sign(to_date(20160401,'yyyymmdd')-VALID_DATE),'1',((WRK_H_SUM-WRK_comm)*7.5/8)*0.065,0),'410MO',decode(sign(to_date(20160101,'yyyymmdd')-VALID_DATE),'1',((WRK_H_SUM-WRK_comm)*7.5/8)*0.01,0),'420TW',decode(sign(to_date(20160101,'yyyymmdd')-VALID_DATE),'1',((WRK_H_SUM-WRK_comm)*7.5/8)*0.01,0),'430TT',decode(sign(to_date(20160101,'yyyymmdd')-VALID_DATE),'1',((WRK_H_SUM-WRK_comm)*7.5/8)*0.01,0),0))   SUM_H_MADE_corr,sum(SUM_H_MADE_norm) SUM_H_MADE_norm,sum(SUM_H_Repair) SUM_H_Repair,sum(Sum_H_demurrage) Sum_H_demurrage,sum(WRK_H_SUM-WRK_comm)  WRK_H_SUM,sum(WRK_H_SUM) WRK_norm,sum(SUM_H_MADE_norm+SUM_H_Repair+Sum_H_demurrage+WRK_H_SUM-WRK_comm+WRK_H_SUM) CHK  FROM (SELECT TRANSACTION_DATE||'_'||C_WORK_GROUP_ID ID_WRK,Nvl(Sum(Decode(SubStr(part_no,1,1),'1',MANUFACTURING_TIME,'4',MANUFACTURING_TIME,'W',MANUFACTURING_TIME)),0) SUM_H_MADE_corr,Nvl(Sum(Decode(SubStr(part_no,1,1),'1',MANUFACTURING_TIME,'4',MANUFACTURING_TIME)),0) SUM_H_MADE_norm,nvl(Sum(Decode(SubStr(part_no,1,1),'P',MANUFACTURING_TIME,0)),0) WRK_comm,nvl(Sum(Decode(SubStr(part_no,1,1),'N',MANUFACTURING_TIME,0)),0) SUM_H_Repair,nvl(Sum(Decode(SubStr(part_no,1,1),'T',MANUFACTURING_TIME,0)),0) Sum_H_demurrage,Decode(SubStr(C_WORK_GROUP_ID,1,5),'510L8','550LP','300OI','310OI',SubStr(C_WORK_GROUP_ID,1,5)) WRK_Cent FROM ifsapp.OPERATION_HISTORY WHERE to_char(TRANSACTION_DATE,'YYYYMM')>=(SELECT decode(sign(to_char(sysdate-730,'DD')-5),'-1',To_Char(sysdate-730,'YYYYMM'),To_Char(sysdate-730,'YYYYMM')) FROM dual) AND To_Date(TRANSACTION_DATE)!=(SELECT To_Date(SYSDATE) FROM dual) AND REVERSED_FLAG_DB='N'AND TRANSACTION_CODE='OPFEED' GROUP BY  C_WORK_GROUP_ID,TRANSACTION_DATE ORDER BY ID_wrk) MADE_g,(SELECT VALID_DATE||'_'||WORK_GROUP_ID ID_WRK,WORK_GROUP_ID,VALID_DATE,Sum(HOURS_QTY) WRK_H_SUM FROM ifsapp.C_WORK_GROUP_DET WHERE MEMBER_WORK_STATE_DB IN ('W','O') GROUP BY VALID_DATE,WORK_GROUP_ID ORDER BY ID_wrk) WRK_g WHERE WRK_g.ID_wrk=MADE_g.ID_wrk group by VALID_DATE,WRK_Cent ORDER BY  VALID_DATE,Department,WORK_CENTER"
rs2.movefirst
do until rs2.eof
for i=0 to 12
	wscript.Echo rs2(i)
next
wscript.Echo "Next record"
loop
