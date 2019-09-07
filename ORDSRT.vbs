
dim flds1()
dim TBL(6,2)
Wscript.Echo "START"
Connect="Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=\\10.0.1.22\sys\KatPryw\RADKOS\Podsumowanie tygodni SITS\M_B.accdb;UID=Admin;PWD= ;"
Set objConnection = CreateObject("ADODB.Connection")
Set rs1 = CreateObject("ADODB.Recordset")
objConnection.Open (connect)
rs1.CursorLocation = 3
rs1.LockType=3
Set rs1.ActiveConnection = objConnection
Connect1="Provider=OraOLEDB.Oracle.1;Password=pass;Persist Security Info=True;User ID=radkos;Data Source=prod8;"
Set objConnection1 = CreateObject("ADODB.Connection")
Set rs2 = CreateObject("ADODB.Recordset")
objConnection1.Open (connect1)
rs2.CursorLocation = 3
rs2.LockType=3
Set rs2.ActiveConnection = objConnection1
teraz=year(now())
if len(cstr(month(now())))=1 then teraz=teraz & "0" & cstr(month(now())) else teraz=teraz & cstr(month(now()))
if len(cstr(day(now())))=1 then teraz=teraz & "0" & cstr(day(now())) else  teraz=teraz & cstr(day(now()))
if len(cstr(Hour(now())))=1 then teraz=teraz & "0" & cstr(Hour(now())) else  teraz=teraz & cstr(Hour(now()))
Wscript.Echo "Pobieram MASKI DNI"
rs1.open "SELECT TOP 7 TMP_DMD.WORK_DAY FROM TMP_DMD GROUP BY TMP_DMD.WORK_DAY"
if not rs1.eof then
	rs1.movefirst
	for i=0 to 6
		teraz1=year(rs1("Work_day"))
		if len((month(rs1("Work_day"))))=1 then teraz1=teraz1 & "0" & cstr(month(rs1("Work_day"))) else teraz1=teraz1 & cstr(month(rs1("Work_day")))
		if len((day(rs1("Work_day"))))=1 then teraz1=teraz1 & "0" & cstr(day(rs1("Work_day"))) else  teraz1=teraz1 & cstr(day(rs1("Work_day")))
		tbl(i,0)=teraz1
		tbl(i,1)="("
		tbl(i,2)="("
		if not rs1.eof then rs1.movenext
	next
end if
rs1.close
Wscript.Echo "Pobieram INDEKSY"
rs1.open "SELECT TMP_DMD.PART_NO,TMP_DMD.RODZAJ, TMP_DMD.QTY_DEMAND, TMP_DMD.BIL_MAG, TMP_DMD.WORK_DAY FROM 6dni INNER JOIN TMP_DMD ON [6dni].WORK_DAY = TMP_DMD.WORK_DAY GROUP BY tMP_DMD.PART_NO, TMP_DMD.RODZAJ, TMP_DMD.QTY_DEMAND,TMP_DMD.BIL_MAG, TMP_DMD.WORK_DAY, TMP_DMD.KOOR HAVING (((TMP_DMD.BIL_MAG)>0) AND ((TMP_DMD.KOOR)<>'*' And (TMP_DMD.KOOR)<>'LUCPRZ'))"
if not rs1.eof then
	Wscript.Echo "... Tworzenie zapytania ..."
	rs1.movefirst
	do while not rs1.eof
		teraz1=year(rs1("Work_day"))
		if len(cstr(month(rs1("Work_day"))))=1 then teraz1=teraz1 & "0" & cstr(month(rs1("Work_day"))) else teraz1=teraz1 & cstr(month(rs1("Work_day")))
		if len(cstr(day(rs1("Work_day"))))=1 then teraz1=teraz1 & "0" & cstr(day(rs1("Work_day"))) else  teraz1=teraz1 & cstr(day(rs1("Work_day")))
		'Wscript.Echo "Zapytanie dla " & rs1("PART_NO") & " DATA " & teraz1
			i=0
			do until teraz1=tbl(i,0)
				'Wscript.Echo teraz1 & " ---> " & tbl(i,0)
				i=i+1
			loop
			'Wscript.Echo teraz1 & " ---> " & tbl(i,0)

			if rs1("QTY_DEMAND")=rs1("BIL_MAG") then
			Wscript.Echo "PE³NY BRAK     :" & rs1(0)  & " DATA " & teraz1
			 	tbl(i,1)=tbl(i,1) & "'" & rs1("PART_NO") & "',"
			else
			Wscript.Echo "Czêœciowy BRAK :" & rs1(0)  & " DATA " & teraz1
				tbl(i,2)=tbl(i,2) & "'" & rs1("PART_NO") & "',"
			end if
		if not rs1.eof then rs1.movenext
	loop
for i=0 to 6
	for j=1 to 2
		tbl(i,j)=left(tbl(i,j),len(tbl(i,j))-1) & ")"
	next
next
rs1.close
rs1.open "Delete * from ALLSUMM",objConnection,0,3
rs1.open "Delete * from PARTSUMM"

else
rs1.close
set rs2=nothing
set objConnection1=nothing
end if
Wscript.Echo "Uruchamiam Procesy zbierajace dane"
strComputer = "."
Const HIDDEN_WINDOW = 0
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")
Set objStartup = objWMIService.Get("Win32_ProcessStartup")
Set objConfig = objStartup.SpawnInstance_
objConfig.ShowWindow = HIDDEN_WINDOW
Set objProcess = objWMIService.Get("Win32_Process")
intReturn = objProcess.Create _
    ("C:\Windows\SysWOW64\cscript.exe //NoLogo D:\boardinfo\TBL_PP.vbs", Null, objConfig, intProcessID)
WScript.Sleep 1000
intReturn = objProcess.Create _
    ("C:\Windows\SysWOW64\cscript.exe //NoLogo D:\boardinfo\TBL_PC6.vbs", Null, objConfig, intProcessID)
WScript.Sleep 1000
'intReturn = objProcess.Create _
    '("C:\Windows\SysWOW64\cscript.exe //NoLogo D:\boardinfo\TBL_NP.vbs", Null, objConfig, intProcessID)
'intReturn = objProcess.Create _
    '("C:\Windows\SysWOW64\cscript.exe //NoLogo D:\boardinfo\TBL_NC.vbs", Null, objConfig, intProcessID)
for i=5 to -3 step -2
	if i<0 then b=-3 else b=0
	if len(tbl(i-b,1))>3 then
	Wscript.Echo tbl(i-b,0)
	if i=0 then label="Decode(Sign(DATE_REQUIRED-To_Date(to_char(SYSDATE,'YYYY-MM-DD'))),'-1',To_Date(to_char(SYSDATE,'YYYY-MM-DD')),DATE_REQUIRED)" else label="DATE_REQUIRED"
rs2.open "SELECT part_no,DATE_REQUIRED,ifsapp.inventory_part_api.Get_Planner_Buyer('ST',a.part_no) Part_Buyer,Decode(ORDER_SUPPLY_DEMAND_TYPE,'Rez. mat. ZP',ifsapp.shop_ord_api.Get_State(a.order_no,a.line_no,a.rel_no)||'//'||Nvl(ifsapp.shop_order_operation_api.Get_Work_Group_Id(a.order_no,a.line_no,a.rel_no,1),Nvl(ifsapp.dop_head_api.Get_C_Trolley_Id(SubStr(ifsapp.shop_ord_api.Get_Source(a.order_no,a.line_no,a.rel_no),10,InStr(ifsapp.shop_ord_api.Get_Source(a.order_no,a.line_no,a.rel_no),'^',10)-10)),' '))||'//', STATUS_DESC)||' //'||ifsapp.shop_order_operation_list_api.Get_Next_Op_Work_Center(a.order_no,a.line_no,a.rel_no,0) STAT ,Nvl(Decode(ORDER_SUPPLY_DEMAND_TYPE,'Rez. mat. ZP',SubStr(ifsapp.shop_ord_api.Get_Source(a.order_no,a.line_no,a.rel_no),10,InStr(ifsapp.shop_ord_api.Get_Source(a.order_no,a.line_no,a.rel_no),'^',10)-10),'Potrzeby DOP',a.ORDER_NO),a.ORDER_NO) DOP_ord,Nvl(Decode(ORDER_SUPPLY_DEMAND_TYPE,'Rez. mat. ZP',REPLACE(SubStr(ifsapp.shop_ord_api.Get_Source(a.order_no,a.line_no,a.rel_no),InStr(ifsapp.shop_ord_api.Get_Source(a.order_no,a.line_no,a.rel_no),'^',10)),'^',''),'Potrzeby DOP',a.LINE_NO),a.info)  DOP_lin,a.ORDER_NO,ORDER_SUPPLY_DEMAND_TYPE, QTY_DEMAND,ifsapp.inventory_part_in_stock_api. Get_Plannable_Qty_Onhand('ST',part_no,'*') mag,Nvl(ifsapp.shop_ord_api.Get_Revised_Qty_Due(a.order_no,a.line_no,a.rel_no),0) Prod_QTY,to_date('" & tbl(i-b,0) & "','yyyymmdd') RPT_DAT,ifsapp.inventory_part_api.Get_Description('ST',a.part_no) Short_nam,Decode(ORDER_SUPPLY_DEMAND_TYPE,'Rez. mat. ZP',Decode(Sign(ifsapp.shop_ord_api.Get_Revised_Due_Date(a.order_no,a.line_no,a.rel_no)-SYSDATE),'-1',To_Date(to_char(SYSDATE,'YYYY-MM-DD')),ifsapp.shop_ord_api.Get_Revised_Due_Date(a.order_no,a.line_no,a.rel_no)),to_date('" & tbl(i-b,0) & "','yyyymmdd')) realizacja FROM ifsapp.order_supply_demand_ext a where ORDER_SUPPLY_DEMAND_TYPE not in ('Zapotrzeb. zakupu','Potrzeby zaplan. w MRP','Zam. zakupu','Zadanie transportu') AND part_no in " & tbl(i-b,1) & " AND " & label & "=(select to_date('" & tbl(i-b,0) & "','yyyymmdd') from dual) ORDER BY DATE_REQUIRED,order_no"
		if not rs2.eof then
			rs2.movefirst
			rs1.open "SElect * from ALLSUMM",objConnection,0,3
			do while not rs2.eof
				Wscript.Echo "ADD NEW DAY ->> PART_NO:" & RS2(0)
				rs1.addnew
				for j=0 to 11
					rs1(j)=rs2(j)
				next
				if rs2(4)=rs2(6) then rs1(12)="MRP" else rs1(12)="DOP"
				rs1(13)=rs2(12)
				rs1(14)=rs2(13)
				rs1.update
			if not rs2.eof then rs2.movenext
			loop
			rs1.close
		end if
	rs2.close
	end if
next


chk="true"
do while not chk="false"
chk="false"
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colProcesses = objWMIService.ExecQuery _
("SELECT * FROM Win32_Process WHERE Name = 'cscript.exe'")
Wscript.Echo "Czekam - Liczba skryptow " & colProcesses.Count
If colProcesses.Count > 0 Then
	For Each objitem In colProcesses
    		if instr(1,objitem.CommandLine,"TBL_")>0 then chk="true"
	Next
end if
WScript.Sleep 2000
loop
WScript.Sleep 1000

Wscript.Echo "Rozpoczynam analizê Czêœciowych BRAKow"
Set rs3 = CreateObject("ADODB.Recordset")
rs3.CursorLocation = 3
rs3.LockType=3
Set rs3.ActiveConnection = objConnection
rs3.open "Select * from ALLSUMM",objConnection,0,3
rs1.open "SELECT PARTSUMM.*,[Komponenty na braku].QTY_DEMAND-[Komponenty na braku].BIL_MAG as BIL_MAG, STATUSY.ID FROM (PARTSUMM INNER JOIN [Komponenty na braku] ON (PARTSUMM.PART_NO = [Komponenty na braku].PART_NO) AND (PARTSUMM.RPT_DAT = [Komponenty na braku].WORK_DAY)) INNER JOIN STATUSY ON MID(PARTSUMM.STAT,1,3)=STATUSY.PRZEDR ORDER BY PARTSUMM.PART_NO,PARTSUMM.RPT_DAT,STATUSY.ID,PARTSUMM.QTY_DEMAND DESC,PARTSUMM.DOP_ORD"
if not rs1.eof then
	rs1.movefirst
	Wscript.Echo rs1.RecordCount
	ind=rs1("part_no") & rs1("RPT_DAT")
	bil=0
	bil_mag=rs1("bil_mag")
	dod=0
	do while not rs1.eof
		if ind<>rs1("Part_no") & rs1("RPT_DAT") then
			bil=0
			bil_mag=rs1("bil_mag")
			ind=rs1("part_no") & rs1("RPT_DAT")
			dod=0
		end if
		bil=bil+rs1("QTY_DEMAND")
		rs3.filter="DOP_ORD='" & rs1("DOP_ORD") & "'"
		if not rs3.eof then
			if bil_mag<bil then
				Wscript.Echo "Dodaje do Calkowitych " & rs1(0)
				rs3.addnew
				for i=0 to 14
					rs3(i)=rs1(i)
				next
				rs3.update
			end if
		else
			if bil_mag<bil then
				rs3.addnew
				for i=0 to 14
					rs3(i)=rs1(i)
				next
				rs3.update
			end if
		end if

	if not rs1.eof then rs1.movenext
	loop
end if
rs3.close
rs1.close
Wscript.Echo "Obliczam mianownik ilosci zlecen w IFS"
rs2.open "SELECT Decode(Sign(REVISED_DUE_DATE-SYSDATE),'-1',To_Date(to_char(SYSDATE,'YYYY-MM-DD')),REVISED_DUE_DATE) WORK_DAY,Decode(source,'','MRP','DOP') TYP,Sum(REVISED_QTY_DUE) QTY_ALL FROM ifsapp.shop_ord WHERE OBJSTATE <> (select ifsapp.SHOP_ORD_API.FINITE_STATE_ENCODE__('Zamkniête') from dual) AND OBJSTATE <> (select ifsapp.SHOP_ORD_API.FINITE_STATE_ENCODE__('Wstrzymany') from dual) AND REVISED_DUE_DATE<=(select ifsapp.work_time_calendar_api.Get_End_Date(ifsapp.site_api.Get_Manuf_Calendar_Id('ST'),To_Date(SYSDATE),6) from dual) GROUP BY Decode(Sign(REVISED_DUE_DATE-SYSDATE),'-1',To_Date(to_char(SYSDATE,'YYYY-MM-DD')),REVISED_DUE_DATE),Decode(source,'','MRP','DOP') ORDER BY work_day,typ"
rs2.movefirst
Wscript.Echo "Dodaje dzien"
Connect2="Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=\\10.0.1.22\sys\KatPryw\RADKOS\Podsumowanie tygodni SITS\M_B.mdb;UID=Admin;PWD= ;"
Set objConnection2 = CreateObject("ADODB.Connection")
objConnection2.Open (connect2)
set rs3=nothing
Set rs3 = CreateObject("ADODB.Recordset")
rs3.CursorLocation = 3
rs3.LockType=3
Set rs3.ActiveConnection = objConnection2
rs3.open "Delete * from DAY_QTY where WORK_DAY<now-1"
rs3.open "Select * from DAY_QTY"
rs1.open "Delete * from DAY_QTY"
rs1.open "SElect * from DAY_QTY",objConnection,0,3
do while not rs2.eof
	Wscript.Echo "Dodaje dzienÅ„ "  & rs2(0)
	rs1.addnew
	for i=0 to 2
		rs1(i)=rs2(i)
	next
	rs1.update
	rs3.filter="Work_day=#" & rs2(0) & "# and typ='" & rs2(1) & "'"
	if rs3.eof then
		rs3.addnew
		for i=0 to 2
			rs3(i)=rs2(i)
		next
		rs3.update
	else
		rs3(2)=rs2(2)
	end if
	if not rs2.eof then rs2.movenext
loop
rs3.filter=0
rs3.close
rs2.close
rs1.close
rs1.open "SELECT [6DNITYPY].WORK_DAY AS DAT, [6DNITYPY].TYP AS PROD_TYPE, Sum(iif(isnull([StAtystyki iloœci braków w prod.PROD_QTY]),0,[StAtystyki iloœci braków w prod.PROD_QTY])) AS PROD_QTY FROM 6DNITYPY LEFT JOIN [StAtystyki iloœci braków w prod] ON ([6DNITYPY].TYP=[StAtystyki iloœci braków w prod].PROD_TYPE) AND ([6DNITYPY].WORK_DAY = [StAtystyki iloœci braków w prod].DAT) GROUP BY [6DNITYPY].WORK_DAY, [6DNITYPY].TYP"
rs3.open "Select * from DAY_QTY"
do while not rs1.eof
	Wscript.Echo "AKtualizuje dzien "  & rs1(0)
	rs3.filter="Work_day=#" & rs1(0) & "# and typ='" & rs1(1) & "'"
		rs3(3)=rs1(2)
	if not rs1.eof then rs1.movenext
loop
rs3.filter=0
rs3.close
rs1.close
'Zaktualizuj AllSUM
set rs3=nothing
Set rs3 = CreateObject("ADODB.Recordset")
rs3.CursorLocation = 3
rs3.LockType=3
Set rs3.ActiveConnection = objConnection2
rs1.open "SElect * from ALLSUMM",objConnection,0,3
rs3.open "Delete * from ALLSUMM",objConnection2,0,3
rs3.open "Select * from ALLSUMM",objConnection2,0,3
Wscript.Echo "No to do roboty..."
Wscript.Echo "Usuwam niepotrzebne indeksy..."
Wscript.Echo "Nowe dane w Calkowitych brakach"
do while not rs1.eof
	rs3.addnew
	for i=0 to 14
		rs3(i)=rs1(i)
	next
	rs3.update
	if not rs1.eof then rs1.movenext
loop
rs1.close
rs3.close
rs1.open "SElect * from PARTSUMM",objConnection,0,3
rs3.open "Delete * from PARTSUMM",objConnection2,0,3
rs3.open "Select * from PARTSUMM",objConnection2,0,3
Wscript.Echo "Nowe dane w czesciowych brakach"
do while not rs1.eof
	rs3.addnew
	for i=0 to 14
		rs3(i)=rs1(i)
	next
	rs3.update
	if not rs1.eof then rs1.movenext
loop
rs1.close
rs3.close
Wscript.Echo "Uruchamiam kwerende kupcow"
rs1.open "SELECT [%$##@_Alias].RPT_DAT, [%$##@_Alias].PART_BUYER, [%$##@_Alias].PROD_TYPE, Sum([%$##@_Alias].PROD_QTY) AS BRAK FROM (SELECT ALLSUMM.RPT_DAT, [StAtystyki iloœci braków w prod].DOP_ord,ALLSUMM.PART_BUYER, [StAtystyki iloœci braków w prod].PROD_TYPE, [StAtystyki iloœci braków w prod].PROD_QTY FROM ALLSUMM INNER JOIN [StAtystyki iloœci braków w prod] ON ALLSUMM.DOP_ORD=[StAtystyki iloœci braków w prod].DOP_ORD GROUP BY ALLSUMM.RPT_DAT, [StAtystyki iloœci braków w prod].DOP_ord, ALLSUMM.PART_BUYER, [StAtystyki iloœci braków w prod].PROD_TYPE, [StAtystyki iloœci braków w prod].PROD_QTY)  AS [%$##@_Alias] GROUP BY [%$##@_Alias].RPT_DAT, [%$##@_Alias].PART_BUYER, [%$##@_Alias].PROD_TYPE"
rs3.open "Delete * from braki_KOOR where work_day<now-1"
rs3.open "select * from braki_KOOR"
do while not rs1.eof
	Wscript.Echo "AKtualizujê KUPCA"  & rs1(0)
	rs3.filter="Work_day=#" & rs1(0) & "# and PROD_type='" & rs1(2) & "' and PART_BUYER='" & rs1(1) & "'"
		if not rs3.eof then
			rs3(3)=cdbl(rs1(3))
		else
			rs3.addnew
			for i=0 to 3
				rs3(i)=rs1(i)
			next
			rs3.update
		end if
	if not rs1.eof then rs1.movenext
loop
rs3.filter=0
rs3.movefirst
do while not rs3.eof
	Wscript.Echo "Sprawdzam czy zapis istnieje "  & rs3(0)
	rs1.filter="RPT_DAT=#" & rs3(0) & "# and PROD_type='" & rs3(2) & "' and PART_BUYER='" & rs3(1) & "'"
		if rs1.eof then
			rs3.delete
		end if
	if not rs3.eof then rs3.movenext
loop
Wscript.Echo "Koniec"
rs1.filter=0
rs1.close
rs3.close
set rs2=nothing
objConnection1.close
objConnection.close
set rs3=nothing
set rs1=nothing
set objConnection=nothing
set objConnection1=nothing


Set objWMIService1 = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")
Set objStartup1 = objWMIService1.Get("Win32_ProcessStartup")
Set objConfig1 = objStartup1.SpawnInstance_
objConfig1.ShowWindow = HIDDEN_WINDOW
Set objProcess1 = objWMIService.Get("Win32_Process")
intReturn = objProcess1.Create _
    ("C:\Windows\SysWOW64\cscript.exe //NoLogo D:\boardinfo\ARCH.vbs", Null, objConfig, intProcessID)
