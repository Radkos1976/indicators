Dim Source
Dim Connect
Dim Rs1
Dim da
dim str
dim mag
dim sup
dim dmd
dim bzam
dim bmag
dim fso
DIM flds()
Wscript.Echo "Utworzono tablicê indeksów"
kl=now() 
Source1="Select * from (SELECT to_char(VALID_DATE,'DD-MM-IYYY')||'_'||WRK_Cent iD_WRK,To_Char(VALID_DATE,'IYYYMM') Month,To_Char(VALID_DATE,'IYYYIW') WEEK,VALID_DATE,ifsapp.work_center_api.Get_Department_No('ST',WRK_Cent) Department,WRK_Cent Work_center,sum(SUM_H_MADE_corr+decode(WRK_Cent,'220SE',decode(sign(to_date(20160401,'yyyymmdd')-VALID_DATE),'1',((WRK_H_SUM-WRK_comm)*7.5/8)*0.065,0),'410MO',decode(sign(to_date(20160101,'yyyymmdd')-VALID_DATE),'1',((WRK_H_SUM-WRK_comm)*7.5/8)*0.01,0),'420TW',decode(sign(to_date(20160101,'yyyymmdd')-VALID_DATE),'1',((WRK_H_SUM-WRK_comm)*7.5/8)*0.01,0),'430TT',decode(sign(to_date(20160101,'yyyymmdd')-VALID_DATE),'1',((WRK_H_SUM-WRK_comm)*7.5/8)*0.01,0),0))   SUM_H_MADE_corr,sum(SUM_H_MADE_norm) SUM_H_MADE_norm,sum(SUM_H_Repair) SUM_H_Repair,sum(Sum_H_demurrage) Sum_H_demurrage,sum(WRK_H_SUM-WRK_comm)  WRK_H_SUM,sum(WRK_H_SUM) WRK_norm,sum(SUM_H_MADE_norm+SUM_H_Repair+Sum_H_demurrage+WRK_H_SUM-WRK_comm+WRK_H_SUM) CHK  FROM (SELECT TRANSACTION_DATE||'_'||C_WORK_GROUP_ID ID_WRK,Nvl(Sum(Decode(SubStr(part_no,1,1),'1',MANUFACTURING_TIME,'4',MANUFACTURING_TIME,'W',MANUFACTURING_TIME)),0) SUM_H_MADE_corr,Nvl(Sum(Decode(SubStr(part_no,1,1),'1',MANUFACTURING_TIME,'4',MANUFACTURING_TIME)),0) SUM_H_MADE_norm,nvl(Sum(Decode(SubStr(part_no,1,1),'P',MANUFACTURING_TIME,0)),0) WRK_comm,nvl(Sum(Decode(SubStr(part_no,1,1),'N',MANUFACTURING_TIME,0)),0) SUM_H_Repair,nvl(Sum(Decode(SubStr(part_no,1,1),'T',MANUFACTURING_TIME,0)),0) Sum_H_demurrage,Decode(SubStr(C_WORK_GROUP_ID,1,5),'510L8','550LP','300OI','310OI',SubStr(C_WORK_GROUP_ID,1,5)) WRK_Cent FROM ifsapp.OPERATION_HISTORY WHERE to_char(TRANSACTION_DATE,'YYYYMM')>=(SELECT decode(sign(to_char(sysdate-730,'DD')-5),'-1',To_Char(sysdate-730,'YYYYMM'),To_Char(sysdate-730,'YYYYMM')) FROM dual) AND To_Date(to_char(TRANSACTION_DATE,'YYYY-MM-DD'))!=(SELECT To_Date(to_char(SYSDATE,'YYYY-MM-DD')) FROM dual) AND REVERSED_FLAG_DB='N'AND TRANSACTION_CODE='OPFEED' GROUP BY  C_WORK_GROUP_ID,TRANSACTION_DATE ORDER BY ID_wrk) MADE_g,(SELECT VALID_DATE||'_'||WORK_GROUP_ID ID_WRK,WORK_GROUP_ID,VALID_DATE,Sum(HOURS_QTY) WRK_H_SUM FROM ifsapp.C_WORK_GROUP_DET WHERE MEMBER_WORK_STATE_DB IN ('W','O') GROUP BY VALID_DATE,WORK_GROUP_ID ORDER BY ID_wrk) WRK_g WHERE WRK_g.ID_wrk=MADE_g.ID_wrk group by VALID_DATE,WRK_Cent union all SELECT to_char(TRANSACTION_DATE,'DD-MM-IYYY')||'_'||obsz||'K' iD_WRK,month,week,TRANSACTION_DATE valid_date,'100KR' DEpartment ,obsz||'K' work_center,Sum(mb) SUM_H_MADE_corr,Sum(mb) SUM_H_MADE_norm,0 SUM_H_Repair,0 Sum_H_demurrage,Sum(hour) WRK_H_SUM,Sum(hour) WRK_norm,owa_opt_lock.checksum(Sum(mb)||Sum(hour)) chk FROM (SELECT To_Char (TRANSACTION_DATE,'YYYYMM') month,To_Char (TRANSACTION_DATE,'IYYYIW') week,TRANSACTION_DATE,b.ID_WRK,b.obsz,Nvl(Sum(IN_DAY),0) mb,WRK_H_SUM hour FROM (SELECT order_no||'_'||To_Date(TRANSACTION_DATE) id,To_Date(tRANSACTION_DATE)||'_'||C_WORK_GROUP_ID id_W,SubStr(C_WORK_GROUP_ID,1,5) obsz,To_Date(TRANSACTION_DATE) TRANSACTION_DATE,C_WORK_GROUP_ID ID_WRK,ORDER_NO FROM ifsapp.OPERATION_HISTORY WHERE SubStr(C_WORK_GROUP_ID,1,1)='1' AND SubStr(part_no,1,1) IN ('4','1') AND REVERSED_FLAG_DB='N'AND TRANSACTION_CODE='OPFEED' GROUP BY To_Date(TRANSACTION_DATE),C_WORK_GROUP_ID,ORDER_NO) b left JOIN (SELECT order_no||'_'||To_Date(DATE_APPLIED) id ,ORDER_NO,DATE_APPLIED,Decode(SubStr(part_no,1,2),'58','110KS','100KT')typ,Sum(to_number(DIRECTION||QUANTITY)*-1) in_day FROM  ifsapp.INVENTORY_TRANSACTION_HIST2 WHERE (SubStr(PART_NO,1,1)='5' or SubStr(PART_NO,1,2)='61') AND ORDER_TYPE= 'Zlec.produkcyjne' GROUP BY order_no,DATE_APPLIED,decode(SubStr(part_no,1,2),'58','110KS','100KT')) a ON  a.ID=b.id AND a.typ=b.obsz left JOIN (SELECT To_Date(VALID_DATE)||'_'||WORK_GROUP_ID ID_WRK,WORK_GROUP_ID,VALID_DATE,Sum(HOURS_QTY) WRK_H_SUM FROM ifsapp.C_WORK_GROUP_DET WHERE MEMBER_WORK_STATE_DB IN ('W','O') GROUP BY VALID_DATE,WORK_GROUP_ID ORDER BY ID_wrk) WRK_g on WRK_g.ID_wrk=b.ID_W WHERE  WRK_H_SUM is not NULL GROUP BY  TRANSACTION_DATE,b.ID_WRK,b.obsz,WRK_H_SUM) GROUP BY month,week,TRANSACTION_DATE,obsz) ORDER BY  VALID_DATE,Department,WORK_CENTER"
Connect="Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D:\boardinfo\WYDAJ.accdb;UID=Admin;PWD= ;"
Connect1="Provider=OraOLEDB.Oracle.1;Password=pass;Persist Security Info=True;User ID=radkos;Data Source=prod8;"
Set objConnection1 = CreateObject("ADODB.Connection")
Set rs2 = CreateObject("ADODB.Recordset")
Wscript.Echo "Pobieram dane z ORACLE"
objConnection1.Open (connect1)
rs2.CursorLocation = 3
rs2.Open Source1, objConnection1,0,4
rs2.movefirst
Wscript.Echo "Od³¹czam zestaw rekordów od Ÿród³a ORACLE"
Set rs2.ActiveConnection = Nothing
'Od³¹cz rekordy ORACLE od po³¹czenia
objConnection1.close
set objConnection1=Nothing
'rs2.movefirst
Set objConnection = CreateObject("ADODB.Connection")
Set rs1 = CreateObject("ADODB.Recordset")
objConnection.Open (connect)
rs1.CursorLocation = 3
rs1.LockType=4
Set rs1.ActiveConnection = objConnection
'rs1.open "Delete * from TMP_DMD where WORK_DAY<now-1"
rs1.open "Select * FROM WYDAJ ORDER BY  VALID_DATE,Department,WORK_CENTER"
Wscript.Echo "Pobieram dane z ACCESS"
Wscript.Echo "Od³¹czam zestaw rekordów od Ÿród³a ACCESS"
'Od³¹cz rekordy ACCESS od po³¹czenia
Set rs1.ActiveConnection = Nothing
objConnection.close
Set objConnection = Nothing
chk=true
actrec=1
j=0
k=0
active_ind=""
Wscript.Echo "No to do roboty..."
Wscript.Echo "Usuwam niepotrzebne indeksy..."
'Kasujemy w rs1 rekordy nie wystêpuj¹ce w rs2
if not rs1.eof then rs1.movefirst
rs2.movefirst
if not rs1.eof then
	active_ind=rs1("iD_WRK")
	rs2.filter="iD_WRK='" & active_ind & "'"
end if
Do While not rs1.EOF
	if active_ind<>rs1("iD_WRK") then
		active_ind=rs1("iD_WRK")
			'Wscript.Echo "Sprawdzam :" & active_ind
		rs2.filter="iD_WRK='" & active_ind & "'"
		if rs2.eof or rs2.bof then
			Wscript.Echo "KASUJE :" & active_ind
			rs1.delete
		end if
	end if
	rs1.movenext
loop

Wscript.Echo "AKtualizujê zapisy ... Usuwanie bilansów z nieistniej¹c¹ dat¹"
rs2.filter=0
'rs1.close
'Kasujemy rekordy z indeksami nie w dacie
R_FLDS=0
redim flds(12,R_FLDS)
'rs1.open "Select * FROM TMP_DMD order by part_no,work_day ",objConnection,0,3
if not rs1.eof then rs1.movefirst
rs2.movefirst
active_ind=rs2("iD_WRK")
maxrec=rs2.RecordCount
rs1.filter="iD_WRK='" & active_ind & "'"
Do While not rs2.EOF and rs2.AbsolutePosition<maxrec
	if active_ind<>rs2("iD_WRK") then
	'zmieni³ siê indeks

		active_ind=rs2("iD_WRK")
		rs1.filter="iD_WRK='" & active_ind & "'"

		'Wscript.Echo "Sprawdzam indeks " & active_ind & " " & R_FLDS
	end if
If not rs1.eof then
	if rs1("SUM_H_MADE_CORR").value<>cdbl(rs2("SUM_H_MADE_CORR").value) or rs1("CHK").value<>cdbl(rs2("CHK").value) then
				Wscript.Echo "Modyfikujê rekord " & active_ind & " " & rs1("iD_WRK")
				rs1("SUM_H_MADE_CORR").value=cdbl(rs2("SUM_H_MADE_CORR").value)
				rs1("SUM_H_MADE_NORM").value=cdbl(rs2("SUM_H_MADE_NORM").value)
				rs1("SUM_H_REPAIR").value=cdbl(rs2("SUM_H_REPAIR").value)
				rs1("SUM_H_DEMURRAGE").value=cdbl(rs2("SUM_H_DEMURRAGE").value)
				rs1("WRK_H_SUM").value=cdbl(rs2("WRK_H_SUM").value)
				rs1("WRK_norm").value=cdbl(rs2("WRK_norm").value)
				rs1("CHK").value=cdbl(rs2("CHK").value)
	end if
	
else
	if rs2.AbsolutePosition<maxrec+1 then
		Wscript.Echo "Dodajê rekord " & active_ind & " " & R_FLDS  & " " & rs2("iD_WRK")
		for i=0 to 12
			flds(i,R_FLDS)=rs2(i)
		next
		R_FLDS=R_FLDS+1
		redim Preserve flds(12,R_FLDS)
	end if
end if
if not rs2.eof then rs2.movenext
if rs2("iD_WRK") is nothing or rs2.eof then exit do
loop
'rs1.close
rs1.filter=0
rs2.close
'Dodajê rekordy z indeksami
'rs1.open "Select * FROM TMP_DMD order by part_no,work_day ",objConnection,0,3
for i=0 to R_FLDS-1
	Wscript.Echo "Dodajê rekord do bazy wiersz :" & i
	rs1.addnew
	for j=0 to 12
		rs1(j)=flds(j,i)
	next
	rs1.update
next
rs1.movefirst
Wscript.Echo "Otwieram Po³¹czenie " & Rs1.RecordCount & ":" & maxrec
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open (connect)
rs1.ActiveConnection = objConnection
Wscript.Echo "UPDATE "
rs1.updatebatch
Wscript.Echo rs1.RecordCount
rs1.close
objConnection.execute "UPDATE WYDAJ set end=true where end=false"
Set Rs2 = Nothing
set source=nothing
objConnection.close
Wscript.Echo "Zamykam Po³¹czenie "
Set objConnection = Nothing
Set Rs1 = Nothing
Source1="SELECT * FROM (SELECT TO_CHAR(VALID_DATE,'DD-MM-IYYY')||'_'||Decode(wydz,'5','500LI','4','400ST','2','200SW',wydz) ID,Decode(wydz,'5','500LI','4','400ST','2','200SW',wydz) Wydz,To_Char(VALID_DATE,'IYYYMM') MONTH,To_Char(VALID_DATE,'IYYYIW') WEEK,To_Date(VALID_DATE) DAT,Sum(HOURS_QTY) Gr_Sred FROM ((SELECT a.EMP_no,a.VALID_DATE,a.wydz,a.HOURS_QTY,b.mi FROM (SELECT EMP_NO,Decode(SubStr(WORK_GROUP_ID,3,1),'9',5,SubStr(WORK_GROUP_ID,3,1)) Wydz,VALID_DATE,HOURS_QTY FROM ifsapp.C_WORK_GROUP_DET WHERE SubStr(WORK_GROUP_ID,1,1)='9' AND SubStr(WORK_GROUP_ID,5,1)='W' AND MEMBER_WORK_STATE_DB IN ('W','O')) a,(SELECT EMP_NO,Decode(SubStr(WORK_GROUP_ID,3,1),'9',5,SubStr(WORK_GROUP_ID,3,1)) Wydz,Min(VALID_DATE) mi FROM ifsapp.C_WORK_GROUP_DET WHERE SubStr(WORK_GROUP_ID,1,1)='9' AND SubStr(WORK_GROUP_ID,5,1)='W' AND MEMBER_WORK_STATE_DB IN ('W','O') GROUP BY EMP_NO,Decode(SubStr(WORK_GROUP_ID,3,1),'9',5,SubStr(WORK_GROUP_ID,3,1)) ) b WHERE b.Emp_NO||'_'||b.wydz=a.Emp_NO||'_'||a.wydz AND a.VALID_DATE-Decode(Sign(a.VALID_DATE-To_Date('20151001','YYYYMMDD')),'-1',21,28)>b.mi) UNION (SELECT EMP_NO,VALID_DATE,Decode(SubStr(WORK_GROUP_ID,3,1),'9',5,SubStr(WORK_GROUP_ID,3,1)) Wydz,HOURS_QTY,SYSDATE mi FROM ifsapp.C_WORK_GROUP_DET WHERE SubStr(WORK_GROUP_ID,1,1)='9' AND SubStr(WORK_GROUP_ID,5,1)!='W' AND MEMBER_WORK_STATE_DB IN ('W','O'))) GROUP BY Decode(wydz,'5','500LI','4','400ST','2','200SW',wydz),VALID_DATE) ORDER BY dat,wydz"
Set objConnection1 = CreateObject("ADODB.Connection")
Set rs2 = CreateObject("ADODB.Recordset")
Wscript.Echo "Pobieram dane z ORACLE"
objConnection1.Open (connect1)
rs2.CursorLocation = 3
rs2.Open Source1, objConnection1,0,4
rs2.movefirst
Wscript.Echo "Od³¹czam zestaw rekordów od Ÿród³a ORACLE"
Set rs2.ActiveConnection = Nothing
'Od³¹cz rekordy ORACLE od po³¹czenia
objConnection1.close
set objConnection1=Nothing
Set objConnection = CreateObject("ADODB.Connection")
Set rs1 = CreateObject("ADODB.Recordset")
objConnection.Open (connect)
rs1.CursorLocation = 3
rs1.LockType=4
Set rs1.ActiveConnection = objConnection
rs1.open "Select * FROM GR_Sred ORDER BY  DAT,WYDZ"
Wscript.Echo "Pobieram dane z ACCESS"
Wscript.Echo "Od³¹czam zestaw rekordów od Ÿród³a ACCESS"
'Od³¹cz rekordy ACCESS od po³¹czenia
Set rs1.ActiveConnection = Nothing
objConnection.close
Set objConnection = Nothing
chk=true
actrec=1
j=0
k=0
active_ind=""
Wscript.Echo "No to do roboty..."
Wscript.Echo "Usuwam niepotrzebne indeksy..."
'Kasujemy w rs1 rekordy nie wystêpuj¹ce w rs2
if not rs1.eof then rs1.movefirst
rs2.movefirst
if not rs1.eof then
	active_ind=rs1("iD")
	rs2.filter="iD='" & active_ind & "'"
end if
Do While not rs1.EOF
	if active_ind<>rs1("iD") then
		active_ind=rs1("iD")
			'Wscript.Echo "Sprawdzam :" & active_ind
		rs2.filter="iD='" & active_ind & "'"
		if rs2.eof or rs2.bof then
			Wscript.Echo "KASUJE :" & active_ind
			rs1.delete
		end if
	end if
	rs1.movenext
loop

Wscript.Echo "AKtualizujê zapisy ... Usuwanie bilansów z nieistniej¹c¹ dat¹"
rs2.filter=0

R_FLDS=0
redim flds(5,R_FLDS)

if not rs1.eof then rs1.movefirst
rs2.movefirst
active_ind=rs2("iD")
maxrec=rs2.RecordCount
rs1.filter="iD='" & active_ind & "'"
Do While not rs2.EOF and rs2.AbsolutePosition<maxrec
	if active_ind<>rs2("iD") then
	'zmieni³ siê indeks

		active_ind=rs2("iD")
		rs1.filter="iD='" & active_ind & "'"

		'Wscript.Echo "Sprawdzam indeks " & active_ind & " " & R_FLDS
	end if
If not rs1.eof then
	if rs1("Gr_sred").value<>cdbl(rs2("Gr_sred").value) then
				Wscript.Echo "Modyfikujê rekord " & active_ind & " " & rs1("iD")
				rs1("Gr_sred").value=cdbl(rs2("Gr_sred").value)
	end if
else
	if rs2.AbsolutePosition<maxrec+1 then
		Wscript.Echo "Dodajê rekord " & active_ind & " " & R_FLDS  & " " & rs2("iD")
		for i=0 to 5
			flds(i,R_FLDS)=rs2(i)
		next
		R_FLDS=R_FLDS+1
		redim Preserve flds(5,R_FLDS)
	end if
end if
if not rs2.eof then rs2.movenext
if rs2("iD") is nothing or rs2.eof then exit do
loop

rs1.filter=0
rs2.close
'Dodajê rekordy z indeksami

for i=0 to R_FLDS-1
	Wscript.Echo "Dodajê rekord do bazy wiersz :" & i
	rs1.addnew
	for j=0 to 5
		rs1(j)=flds(j,i)
	next
	rs1.update
next
rs1.movefirst
Wscript.Echo "Otwieram Po³¹czenie " & Rs1.RecordCount & ":" & maxrec
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open (connect)
rs1.ActiveConnection = objConnection
Wscript.Echo "UPDATE "
rs1.updatebatch
Wscript.Echo rs1.RecordCount
rs1.close	


set source=nothing
objConnection.close
Wscript.Echo "Zamykam Po³¹czenie "
Set objConnection = Nothing
Set Rs1 = Nothing
