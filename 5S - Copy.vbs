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
kl=now()
Dim objIEDebugWindow
on error resume next
Debug "SKRypt 5s"
Source1="SELECT * FROM `\\10.0.1.22\sys\KatPryw\RADKOS\Podsumowanie tygodni SITS\ZestawienieDlaKJ.accdb`.`5simp`"
Connect="Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D:\boardinfo\WYDAJ.accdb;UID=Admin;PWD= ;"
Connect1="DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=\\10.0.1.22\sys\KatPryw\RADKOS\Podsumowanie tygodni SITS\ZestawienieDlaKJ.accdb;UID=Admin;PWD= ;"
Debug "£¹cze z pierwsz¹ baz¹"
Set objConnection1 = CreateObject("ADODB.Connection")
Set rs2 = CreateObject("ADODB.Recordset")
Wscript.Echo "Pobieram dane z KJ"
Debug "Pobieram dane z KJ"
objConnection1.Open (connect1)
Debug "£¹cze zestaw " & Err.Description
rs2.CursorLocation = 3
rs2.LockType=1
Set rs2.ActiveConnection = objConnection1
rs2.Open Source1
Debug "Otwieram " & Err.Description
rs2.movefirst
Debug "Od³¹czam zestaw rekordów od Ÿród³a ORACLE " & Err.Description
on error goto 0
Wscript.Echo "Od³¹czam zestaw rekordów od Ÿród³a ORACLE"
Set rs2.ActiveConnection = Nothing
'Od³¹cz rekordy ORACLE od po³¹czenia
objConnection1.close
set objConnection1=Nothing
'rs2.movefirst
Debug "Druga baza"
Set objConnection = CreateObject("ADODB.Connection")
Set rs1 = CreateObject("ADODB.Recordset")
objConnection.Open (connect)
rs1.CursorLocation = 3
rs1.LockType=4
Set rs1.ActiveConnection = objConnection
'rs1.open "Delete * from TMP_DMD where WORK_DAY<now-1"
rs1.open "SELECT * FROM `D:\boardinfo\WYDAJ.accdb`.`Wynik5S`"
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
	active_ind=rs1("Tydzieñ")
	active_ind1=rs1("Gniazgo")
	active_ind2=rs1("MPK")
	rs2.filter="Tydzieñ='" & active_ind & "' AND Gniazgo='" & active_ind1 & "' AND PK='" & active_ind2 & "'"
end if
Do While not rs1.EOF
	if active_ind&active_ind1&active_ind2<>rs1("Tydzieñ") & rs1("Gniazgo")& rs1("MPK")then
		active_ind=rs1("Tydzieñ")
		active_ind1=rs1("Gniazgo")
		active_ind2=rs1("MPK")
			'Wscript.Echo "Sprawdzam :" & active_ind
		rs2.filter="Tydzieñ='" & active_ind & "' AND Gniazgo='" & active_ind1 & "' AND PK='" & active_ind2 & "'"
		if rs2.eof or rs2.bof then
			Wscript.Echo "KASUJE :" & active_ind & "_" & active_ind1
			rs1.delete
		end if
	end if
	rs1.movenext
loop

Wscript.Echo "AKtualizujê zapisy ... Usuwanie bilansów z nieistniej¹c¹ dat¹"
rs2.filter=0
R_FLDS=0
redim flds(4,R_FLDS)
if not rs1.eof then rs1.movefirst
rs2.movefirst
active_ind=rs2("Tydzieñ")
active_ind1=rs2("Gniazgo")
active_ind2=rs2("PK")
maxrec=rs2.RecordCount
rs1.filter="Tydzieñ='" & active_ind & "' AND Gniazgo='" & active_ind1 & "' AND MPK='" & active_ind2 & "'"
Do While not rs2.EOF
	if active_ind&active_ind1&active_ind2<>rs2("Tydzieñ") & rs2("Gniazgo") &rs2("PK") then
	'zmieni³ siê indeks
		active_ind=rs2("Tydzieñ")
		active_ind1=rs2("Gniazgo")
		active_ind2=rs2("PK")
		rs1.filter="Tydzieñ='" & active_ind & "' AND Gniazgo='" & active_ind1 & "' AND MPK='" & active_ind2 & "'"

		'Wscript.Echo "Sprawdzam indeks " & active_ind & "_" & active_ind1 & " " & active_ind2 & " " & R_FLDS
	end if
If not rs1.eof then
	a=rs2("Wynik").value
	b=rs1("Wynik").value
	'Wscript.Echo a & "=" & b
	If a=b then
	else
				Wscript.Echo "Modyfikujê rekord " & active_ind & "_" & active_ind1 & "_" & active_ind2 &"-"&rs1.recordcount
				rs1("WYNIK")=rs2("WYNIK")
	end if
else
	if rs2.AbsolutePosition<maxrec+1 then
		Wscript.Echo "Dodajê rekord " & active_ind & "_" & active_ind1 & " " & R_FLDS
		for i=0 to 4
			flds(i,R_FLDS)=rs2(i)
		next
		R_FLDS=R_FLDS+1
		redim Preserve flds(4,R_FLDS)
	end if
end if
if not rs2.eof then rs2.movenext
if rs2("Tydzieñ") is nothing or rs2.eof then exit do
loop
rs2.close
Set rs2= Nothing
rs1.filter=0
'Dodajê rekordy z indeksami
'rs1.open "Select * FROM TMP_DMD order by part_no,work_day ",objConnection,0,3
for i=0 to R_FLDS-1
	Wscript.Echo "Dodajê rekord do bazy wiersz :" & i
	rs1.addnew
	for j=0 to 3
		rs1(j)=flds(j,i)
	next
	rs1.update
next
rs1.movefirst
Wscript.Echo "L_rekord " & rs1.RecordCount
Wscript.Echo "Otwieram Po³¹czenie " & Rs1.RecordCount & ":" & maxrec
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open (connect)
rs1.ActiveConnection = objConnection
Wscript.Echo "UPDATE "
rs1.updatebatch
rs1.close
objConnection.close
Wscript.Echo "Zamykam Po³¹czenie "
Set objConnection = Nothing
Set rs1 = Nothing

Source1="SELECT * FROM (SELECT To_Char(DATE_APPLIED,'YYYYMM') MONTH,'SUR' TYP,sum(Decode(LOCATION_NO,'9500-9510',Decode(DIRECTION,'+',2,'-',-1,0),0))/sum(Decode(LOCATION_NO,'0100',1,0))*1000 wsp_brk,'4520' MPK FROM IFSAPP.INVENTORY_TRANSACTION_HIST2 where To_Date(DATE_APPLIED)!=(SELECT To_Date(SYSDATE) FROM dual) and ((LOCATION_NO='9500-9510' AND TRANSACTION_CODE!='COUNT-OUT') OR (LOCATION_NO='0100' AND TRANSACTION_CODE IN ('INVM-ISS' ,'INVM-TRISS'))) AND SubStr(PART_NO,1,1)='6' GROUP BY To_Char(DATE_APPLIED,'YYYYMM') ORDER BY MONTH) UNION ALL (SELECT zest.MONTH,zest.typ,sum(dtasor.L_rekl)/Sum(dtasor.IL)*1000 wsp_brk,Decode(zest.typ,'IKEA','502i','SITS','502s') AS MPK FROM (SELECT c.MONTH,St_d,end_d,b.typ FROM (SELECT MONTH,Decode(MONTH,'201607','201607',To_Char(Add_Months(To_Date(MONTH,'yyyymm'),-1),'yyyymm')) st_d,MONTH end_d from (SELECT To_Char(WORK_DAY,'YYYYMM') MONTH,To_Date(To_Char(WORK_DAY,'YYYYMM')||'01','YYYYMMdd') da FROM ifsapp.work_time_calendar_pub where CALENDAR_ID='SITS' AND WORK_DAY BETWEEN To_Date(20150101,'YYYYMMDD') AND SYSDATE  GROUP BY To_Char(WORK_DAY,'YYYYMM')) ORDER BY month) c,(SELECT 'IKEA' typ FROM dual UNION ALL SELECT 'SITS' typ FROM dual) b ORDER BY c.MONTH ) zest,(SELECT a.MONTH,a.TYP,a.L_REKL,meb.IL FROM (SELECT TYP,MONTH,Sum(L_REKL) L_REKL FROM ( SELECT * from (SELECT Decode(SubStr(ifsapp.return_material_api.Get_Customer_No(RMA_NO),1,4),'IKEA','IKEA','164','IKEA','3721','IKEA','002-','IKEA','4000','IKEA','1752','IKEA','3445','IKEA','3711','IKEA','3613','IKEA','5100','IKEA','885','IKEA','622','IKEA','494','IKEA','458','IKEA','SITS') typ,To_Char(Nvl(DATE_RETURNED,ifsapp.return_material_api.Get_Registration_Date(rma_no)),'YYYYMM') month,Sum(QTY_RECEIVED) L_rekl FROM ifsapp.RETURN_MATERIAL_LINE WHERE SubStr(RETURN_REASON_CODE,1,1)=4 AND subStr(RETURN_REASON_CODE,1,3)!='4.8' and (DATE_RETURNED IS NOT NULL OR (ifsapp.return_material_api.Get_Status_Id(RMA_NO) IN (7,11) and QTY_RECEIVED IS NOT NULL)) GROUP BY  Decode(SubStr(ifsapp.return_material_api.Get_Customer_No(RMA_NO),1,4),'IKEA','IKEA','164','IKEA','3721','IKEA','002-','IKEA','4000','IKEA','1752','IKEA','3445','IKEA','3711','IKEA','3613','IKEA','5100','IKEA','885','IKEA','622','IKEA','494','IKEA','458','IKEA','SITS'),To_Char(Nvl(DATE_RETURNED,ifsapp.return_material_api.Get_Registration_Date(rma_no)),'YYYYMM') ORDER BY  To_Char(Nvl(DATE_RETURNED,ifsapp.return_material_api.Get_Registration_Date(rma_no)),'YYYYMM')) UNION all (SELECT typ,month,L_REKL FROM (SELECT To_Char(WORK_DAY,'YYYYMM') month,0 L_REKL FROM ifsapp.work_time_calendar_pub where CALENDAR_ID='SITS' AND WORK_DAY BETWEEN To_Date(20150101,'YYYYMMDD') AND SYSDATE GROUP BY To_Char(WORK_DAY,'YYYYMM')) c,(SELECT 'IKEA' typ FROM dual UNION ALL SELECT 'SITS' typ FROM dual) b)) GROUP BY  TYP,MONTH ORDER BY MONTH,TYP) a,(select To_Char(INVOICE_DATE,'YYYYMM') Month,Decode(SubStr(ifsapp.customer_order_api.Get_Customer_No(ORDER_NO),1,4),'IKEA','IKEA','164','IKEA','3721','IKEA','002-','IKEA','4000','IKEA','1752','IKEA','3445','IKEA','3711','IKEA','3613','IKEA','5100','IKEA','885','IKEA','622','IKEA','494','IKEA','458','IKEA','SITS') typ,Sum(INVOICED_QTy)*Decode(To_Char(INVOICE_DATE,'YYYYMM'),To_Char(SYSDATE,'YYYYMM'),wsp,1) IL,wsp FROM ifsapp.CUSTOMER_ORDER_INV_ITEM_JOIN,(SELECT decode(to_char(sysdate,'dd'),'01',1,ifsapp.work_time_calendar_api.Get_Work_Days_Between('SITS',To_Date(To_Char(sysdate,'yyyymm'),'yyyymm'),add_months(To_Date(To_Char(sysdate,'yyyymm'),'yyyymm'),1))/ifsapp.work_time_calendar_api.Get_Work_Days_Between('SITS',To_Date(To_Char(SYSDATE,'yyyymm'),'yyyymm'),SYSDATE)) wsp FROM dual ) WHERE  To_Date(INVOICE_DATE)!=(SELECT To_Date(SYSDATE) FROM dual) AND CATALOG_NO IS NOT NULL AND SubStr(CATALOG_NO,2,1)!='U' AND SubStr(CATALOG_NO,1,1) IN ('2','1') GROUP BY To_Char(INVOICE_DATE,'YYYYMM'),decode(SubStr(ifsapp.customer_order_api.Get_Customer_No(ORDER_NO),1,4),'IKEA','IKEA','164','IKEA','3721','IKEA','002-','IKEA','4000','IKEA','1752','IKEA','3445','IKEA','3711','IKEA','3613','IKEA','5100','IKEA','885','IKEA','622','IKEA','494','IKEA','458','IKEA','SITS'),wsp ORDER BY MONTH,typ) meb WHERE meb.MONTH=a.MONTH AND meb.typ=a.typ) dtasor WHERE dtasor.typ=zest.typ AND (dtasor.MONTH=zest.st_d OR dtasor.MONTH=zest.end_d)  GROUP BY zest.MONTH,zest.typ)"
Connect="Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D:\boardinfo\WYDAJ.accdb;UID=Admin;PWD= ;"
Connect1="DRIVER={Microsoft ODBC for Oracle};UID=RADKOS;SERVER=prod8;PWD=pass"
Set objConnection1 = CreateObject("ADODB.Connection")
Set rs2 = CreateObject("ADODB.Recordset")
Wscript.Echo "Pobieram dane z ORACLE"
objConnection1.Open (connect1)
rs2.CursorLocation = 3
rs2.LockType=4
Set rs2.ActiveConnection = objConnection1
rs2.Open Source1
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
rs1.open "SELECT * FROM MGWiS"
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
	active_ind=rs1("MONTH")
	active_ind1=rs1("TYP")
	active_ind2=rs1("MPK")
	rs2.filter="MONTH='" & active_ind & "' AND TYP='" & active_ind1 & "' AND MPK='" & active_ind2 & "'"
end if
Do While not rs1.EOF
	if active_ind&active_ind1&active_ind2<>rs1("MONTH") & rs1("TYP")& rs1("MPK")then
		active_ind=rs1("MONTH")
		active_ind1=rs1("TYP")
		active_ind2=rs1("MPK")
			'Wscript.Echo "Sprawdzam :" & active_ind
		rs2.filter="MONTH='" & active_ind & "' AND TYP='" & active_ind1 & "' AND MPK='" & active_ind2 & "'"
		if rs2.eof or rs2.bof then
			Wscript.Echo "KASUJE :" & active_ind & "_" & active_ind1
			rs1.delete
		end if
	end if
	rs1.movenext
loop

Wscript.Echo "AKtualizujê zapisy ... Usuwanie bilansów z nieistniej¹c¹ dat¹"
rs2.filter=0
R_FLDS=0
redim flds(3,R_FLDS)
if not rs1.eof then rs1.movefirst
rs2.movefirst
active_ind=rs2("MONTH")
active_ind1=rs2("TYP")
active_ind2=rs2("MPK")
maxrec=rs2.RecordCount
rs1.filter="MONTH='" & active_ind & "' AND TYP='" & active_ind1 & "' AND MPK='" & active_ind2 & "'"
Do While not rs2.EOF
	if active_ind&active_ind1&active_ind2<>rs2("MONTH") & rs2("TYP") &rs2("MPK") then
	'zmieni³ siê indeks
		active_ind=rs2("MONTH")
		active_ind1=rs2("TYP")
		active_ind2=rs2("MPK")
		rs1.filter="MONTH='" & active_ind & "' AND TYP='" & active_ind1 & "' AND MPK='" & active_ind2 & "'"

		'Wscript.Echo "Sprawdzam indeks " & active_ind & "_" & active_ind1 & " " & R_FLDS
	end if
If not rs1.eof then
	if cstr(rs1("WSP_BRK"))<>cstr(rs2("WSP_BRK")) then
				Wscript.Echo "Modyfikujê rekord " & active_ind & "_" & active_ind1 & "_" & active_ind2 &"-"&rs1.recordcount
				rs1("WSP_BRK")=rs2("WSP_BRK")
	end if
else
	if rs2.AbsolutePosition<maxrec+1 then
		Wscript.Echo "Dodajê rekord " & active_ind & "_" & active_ind1 & " " & R_FLDS
		for i=0 to 3
			flds(i,R_FLDS)=rs2(i)
		next
		R_FLDS=R_FLDS+1
		redim Preserve flds(3,R_FLDS)
	end if
end if
if not rs2.eof then rs2.movenext
if rs2("MONTH") is nothing or rs2.eof then exit do
loop
rs2.close
Set rs2= Nothing
rs1.filter=0
'Dodajê rekordy z indeksami
'rs1.open "Select * FROM TMP_DMD order by part_no,work_day ",objConnection,0,3
for i=0 to R_FLDS-1
	Wscript.Echo "Dodajê rekord do bazy wiersz :" & i
	rs1.addnew
	for j=0 to 3
		rs1(j)=flds(j,i)
	next
	rs1.update
next
rs1.movefirst
Wscript.Echo "L_rekord " & rs1.RecordCount
Wscript.Echo "Otwieram Po³¹czenie " & Rs1.RecordCount & ":" & maxrec
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open (connect)
rs1.ActiveConnection = objConnection
Wscript.Echo "UPDATE "
rs1.updatebatch
rs1.close
objConnection.close
Wscript.Echo "Zamykam Po³¹czenie "
Set objConnection = Nothing
Set rs1 = Nothing
Sub Debug( myText )
	'Dim fso, f
    	'Set fso = CreateObject("Scripting.FileSystemObject")
    	'Set f = fso.OpenTextFile("D:\boardinfo\5S.txt", 8)
    	'f.Writeline  DATE  & " " & time & " " & myText
 	'f.Close
	'set fso=nothing
	
End Sub