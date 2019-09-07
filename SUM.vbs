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
Source1="SELECT PART_NO,ifsapp.work_time_calendar_api.Get_Nearest_Work_Day(ifsapp.site_api.Get_Manuf_Calendar_Id('ST'),DATE_REQUIRED) WORK_DAY,ifsapp.inventory_part_api.Get_Part_Product_Code('ST',part_no) Rodzaj,ifsapp.inventory_part_api.Get_Planner_Buyer('ST',part_no) KOOR,Sum(QTY_SUPPLY) QTY_SUPPLY,Sum(QTY_DEMAND) QTY_DEMAND,ifsapp.inventory_part_in_stock_api.Get_Plannable_Qty_Onhand ('ST',part_no,'*') mag,ifsapp.inventory_part_api.Get_Volume_Net('ST',part_no) VOlume,ifsapp.inventory_part_def_loc_api.Get_Location_No('ST',PART_NO) Location FROM (SELECT PART_NO,Decode(Sign(DATE_REQUIRED-To_Date(To_Char(SYSDATE,'YYYY-MM-DD'))),'-1',To_Date(To_Char(SYSDATE,'YYYY-MM-DD')),Decode(ORDER_SUPPLY_DEMAND_TYPE,'Zam. zakupu',ifsapp.work_time_calendar_api.Get_Next_Work_Day(ifsapp.site_api.Get_Manuf_Calendar_Id('ST'),DATE_REQUIRED),DATE_REQUIRED)) DATE_REQUIRED,QTY_SUPPLY,Sum(QTY_DEMAND) QTY_DEMAND  FROM ifsapp.order_supply_demand WHERE  SubStr(part_no,1,1) IN ('5','6') AND ORDER_SUPPLY_DEMAND_TYPE NOT IN ('Zapotrzeb. zakupu','Zadanie transportu') GROUP BY PART_NO,DATE_REQUIRED,ORDER_SUPPLY_DEMAND_TYPE,QTY_SUPPLY UNION ALL  SELECT PART_NO,Decode(Sign(DATE_REQUIRED-To_date(To_Char(SYSDATE,'YYYY-MM-DD'))),'-1',To_Date(To_Char(SYSDATE,'YYYY-MM-DD')),DATE_REQUIRED),0 QTY_SUPPLY,Sum(QTY_DEMAND) QTY_DEMAND FROM ifsapp.dop_order_demand_ext WHERE SubStr(part_no,1,1) IN ('5','6')GROUP BY PART_NO,DATE_REQUIRED) WHERE  DATE_REQUIRED<=(select SYSDATE+30 from dual) GROUP BY PART_NO,DATE_REQUIRED order by PART_NO,DATE_REQUIRED"
Connect="Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=\\10.0.1.22\sys\KatPryw\RADKOS\History.mdb;UID=Admin;PWD= ;"
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
rs1.open "Delete * from TMP_DMD where WORK_DAY<now-1"
rs1.open "Select * FROM TMP_DMD where part_no is not null order by part_no,work_day"
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
rs1.movefirst
rs2.movefirst
if not rs1.eof then
	active_ind=rs1("PART_NO")
	rs2.filter="PART_no='" & active_ind & "'"
end if
Do While not rs1.EOF
	if active_ind<>rs1("PART_NO") then
		active_ind=rs1("PART_NO")
			'Wscript.Echo "Sprawdzam :" & active_ind
		rs2.filter="PART_no='" & active_ind & "'"
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
redim flds(7,R_FLDS)
'rs1.open "Select * FROM TMP_DMD order by part_no,work_day ",objConnection,0,3
if not rs1.eof then rs1.movefirst
rs2.movefirst
mag=cdbl(rs2("MAG").value)
active_ind=rs2("PART_NO")
maxrec=rs2.RecordCount
rs1.filter="PART_no='" & active_ind & "'"
sup=0
dmd=0
bzam=cdbl(mag)
bmag=cdbl(mag)
Do While not rs2.EOF and rs2.AbsolutePosition<maxrec
	if active_ind<>rs2("PART_NO") then
	'zmieni³ siê indeks
		Do While (not rs1.EOF)
			Wscript.Echo "Kasujê indeks w dalszej dacie" & active_ind & " " & rs1("WORK_DAY")
			if rs1("Work_day")> last_dta then
				rs1.delete
			end if
			if not rs1.eof then rs1.movenext
		loop
		active_ind=rs2("PART_NO")
		rs1.filter="PART_no='" & active_ind & "'"
		mag=cdbl(rs2("MAG").value)

		bzam=cdbl(mag)
		bmag=cdbl(mag)
		'Wscript.Echo "Sprawdzam indeks " & active_ind & " " & R_FLDS
	end if
	last_dta=rs2("Work_day")
If not rs1.eof then
	if rs1("WORK_DAY")<rs2("WORK_DAY") then
		Wscript.Echo "Kasujê indeks data wczeœniejsza :" & active_ind & " " & rs1("WORK_DAY")
		rs1.delete
		'rs1.update
		if not rs1.eof then rs1.movenext
	else
		if rs1("WORK_DAY")=rs2("WORK_DAY") then
			s=cdbl(rs2("QTY_SUPPLY").value)
			d=cdbl(rs2("QTY_DEMAND").value)

			bzam=cdbl(bzam)-cdbl(d)+cdbl(s)
			bmag=cdbl(bmag)-cdbl(d)
			if bzam<0 then
				if bzam*-1<d then brakzam=bzam*-1 else brakzam=d
			else
				brakzam=0
			end if
			if bmag<0 then
				if bmag*-1<d then brakmag=bmag*-1 else brakmag=d
			else
				brakmag=0
			end if
			if rs1("QTY_SUPPLY").value<>cdbl(rs2("QTY_SUPPLY").value) or rs1("QTY_DEMAND").value<>cdbl(rs2("QTY_DEMAND").value) or rs1("BIL_ZAM_MAG").value<>brakzam or rs1("BIL_MAG").value<>brakmag  then
				Wscript.Echo "Modyfikujê rekord " & active_ind & " " & rs1("WORK_DAY")
				if rs1("KOOR")<>rs2("KOOR") then rs1("KOOR")=rs2("KOOR")
				rs1("QTY_SUPPLY").value=cdbl(rs2("QTY_SUPPLY").value)
				rs1("QTY_DEMAND").value=cdbl(rs2("QTY_DEMAND").value)
				rs1("BIL_ZAM_MAG").value=cdbl(brakzam)
				rs1("BIL_MAG").value=cdbl(brakmag)
			else
				if rs1("KOOR")<>rs2("KOOR") then rs1("KOOR")=rs2("KOOR")
			end if
			if not rs1.eof then rs1.movenext
			if not rs2.eof then rs2.movenext
		else
			if rs1("WORK_DAY")>rs2("WORK_DAY") then
				Wscript.Echo "Dodajê " & active_ind & " " & R_FLDS & " " & rs2("WORK_DAY")
				s=cdbl(rs2("QTY_SUPPLY").value)
				d=cdbl(rs2("QTY_DEMAND").value)

				bzam=cdbl(bzam)-cdbl(d)+cdbl(s)
				bmag=cdbl(bmag)-cdbl(d)
			if bzam<0 then
				if bzam*-1<d then brakzam=bzam*-1 else brakzam=d
			else
				brakzam=0
			end if
			if bmag<0 then
				if bmag*-1<d then brakmag=bmag*-1 else brakmag=d
			else
				brakmag=0
			end if
				for i=0 to 5
					flds(i,R_FLDS)=rs2(i)
				next
				flds(6,R_FLDS)=cdbl(brakzam)
				flds(7,R_FLDS)=cdbl(brakmag)
				R_FLDS=R_FLDS+1
				last_dta=rs2("Work_day")
				redim Preserve flds(7,R_FLDS)
				if not rs2.eof then rs2.movenext
			end if
		end if
	end if
else
	if rs2.AbsolutePosition<maxrec+1 then
	Do While active_ind=rs2("PART_NO").value
		Wscript.Echo "Dodajê rekord " & active_ind & " " & R_FLDS  & " " & rs2("WORK_DAY")
				s=cdbl(rs2("QTY_SUPPLY").value)
				d=cdbl(rs2("QTY_DEMAND").value)

				bzam=cdbl(bzam)-cdbl(d)+cdbl(s)
				bmag=cdbl(bmag)-cdbl(d)
			if bzam<0 then
				if bzam*-1<d then brakzam=bzam*-1 else brakzam=d
			else
				brakzam=0
			end if
			if bmag<0 then
				if bmag*-1<d then brakmag=bmag*-1 else brakmag=d
			else
				brakmag=0
			end if
		for i=0 to 5
			flds(i,R_FLDS)=rs2(i)
		next
		flds(6,R_FLDS)=cdbl(brakzam)
		flds(7,R_FLDS)=cdbl(brakmag)
		R_FLDS=R_FLDS+1
		redim Preserve flds(7,R_FLDS)
		last_dta=rs2("Work_day")
		if not rs2.eof then rs2.movenext
		if rs2("part_no") is nothing or rs2.eof then exit do
	loop
	end if
end if

loop
rs1.filter=0
'rs1.close
rs2.close
'Dodajê rekordy z indeksami
'rs1.open "Select * FROM TMP_DMD order by part_no,work_day ",objConnection,0,3
for i=0 to R_FLDS-1
	'Wscript.Echo "Dodajê rekord do bazy wiersz :" & i
	rs1.addnew
	for j=0 to 7
		rs1(j)=flds(j,i)
	next
	rs1.update
next
Wscript.Echo "Otwieram Po³¹czenie " & Rs1.RecordCount & ":" & maxrec
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open (connect)
rs1.ActiveConnection = objConnection
Wscript.Echo "UPDATE "
rs1.movefirst
rs1.updatebatch 
Wscript.Echo rs1.RecordCount
rs1.close
objConnection.execute "DELETE TMP_DMD.PART_NO FROM TMP_DMD WHERE (((TMP_DMD.PART_NO) Is Null))"
objConnection.execute "INSERT INTO TMP_DMD (WORK_DAY,KOOR,RODZAJ,QTY_DEMAND,QTY_SUPPLY,BIL_ZAM_MAG,BIL_MAG ) SELECT [Zerowe potrzeby dzieñ].WORK_DAY,[Zerowe potrzeby dzieñ].KOOR,[Zerowe potrzeby dzieñ].RODZAJ,0.000001 AS Wyr1,0 AS Wyr2, 0 AS Wyr3, 0 AS Wyr4 FROM [Zerowe potrzeby dzieñ]"
set source=nothing
objConnection.close
Wscript.Echo "Zamykam Po³¹czenie "
Set objConnection = Nothing
Set Rs1 = Nothing

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
chk="false"
chk1="false"
Set colProcesses = objWMIService.ExecQuery _
("SELECT * FROM Win32_Process WHERE Name = 'cscript.exe'")
Wscript.Echo "Liczba skryptów " & colProcesses.Count
If colProcesses.Count > 0 Then
 	For Each objitem In colProcesses
    		if instr(1,objitem.CommandLine,"ORDSRT")>0 then chk="true"
	Next
end if

Const HIDDEN_WINDOW = 0
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")
Set objStartup = objWMIService.Get("Win32_ProcessStartup")
Set objConfig = objStartup.SpawnInstance_
objConfig.ShowWindow = HIDDEN_WINDOW
Set objProcess = objWMIService.Get("Win32_Process")

if chk<>"true" then
intReturn = objProcess.Create _
    ("C:\Windows\SysWOW64\cscript.exe //NoLogo D:\boardinfo\ORDSRT.vbs", Null, objConfig, intProcessID)
end if
Wscript.Echo "All process Execute in " & (now()-kl)*24*60
chk="true"
do while not chk="false"
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
chk="false"
chk1="false"
Set colProcesses = objWMIService.ExecQuery _
("SELECT * FROM Win32_Process WHERE Name = 'cscript.exe'")
If colProcesses.Count > 1 Then
 chk="true"
end if
WScript.Sleep 2000
loop
Set fs = CreateObject("Scripting.FileSystemObject")
Set ts = fs.OpenTextFile("\\10.0.1.22\sys\KatPryw\RADKOS\compactall.vbs")
body = ts.ReadAll
ts.Close
Execute body
'loop
