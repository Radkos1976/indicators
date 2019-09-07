DIM flds()
Dim connect
Connect1="DRIVER={Oracle in OraClient12Home2_32bit};Pwd=pass;Uid=radkos;DBQ=prod8;QTO=F;"
connect="Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D:\boardinfo\WYDAJ.accdb;UID=Admin;PWD= ;"
conIMP="Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=\\10.0.1.22\sys\KatPryw\RADKOS\HISTORY.mdb;UID=Admin;PWD= ;"
Set objConnection = CreateObject("ADODB.Connection")
Set rs1 = CreateObject("ADODB.Recordset")
objConnection.Open (connect)
rs1.CursorLocation = 3
rs1.LockType=3
Set rs1.ActiveConnection = objConnection
Set objConnection1 = CreateObject("ADODB.Connection")
Set rs2 = CreateObject("ADODB.Recordset")
objConnection1.Open (connect1)
rs2.CursorLocation = 3
rs2.LockType=3
Set rs2.ActiveConnection = objConnection1
offset=0.02083333333333333333333333333333
teraz=year(now())
if len(cstr(month(now()-offset)))=1 then teraz=teraz & "0" & cstr(month(now()-offset)) else teraz=teraz & cstr(month(now()-offset))
if len(cstr(day(now()-offset)))=1 then teraz=teraz & "0" & cstr(day(now()-offset)) else  teraz=teraz & cstr(day(now()-offset))
if len(cstr(int(Hour(now()-offset)/6)))=1 then teraz=teraz & "0" & cstr(int(Hour(now()-offset)/6)) else  teraz=teraz & cstr(int(Hour(now()-offset)/6))
Wscript.Echo teraz
typ="POPRZEDNI DZIEÑ"
Set fsx = CreateObject("Scripting.FileSystemObject")
Set tsx = fsx.OpenTextFile("D:\boardinfo\SQL.txt")
conq = tsx.ReadAll
tsx.Close
rs1.open "select * from TERMINOWOŒÆ where refr='" & teraz & "' and Typ_rap='POPRZEDNI DZIEÑ'"
Wscript.Echo "Po³¹czenie z ACCESS"
if rs1.eof then
	'Connect1="DRIVER={Oracle in instantclient_11_2};DBQ=PROD8;UID=RADKOS;PWD=pass;"
	Wscript.Echo "Po³¹czenie z ORACLE"
	rs2.open replace(replace(conq,"_teraz_",teraz),"_TYP_",typ)
	rs2.movefirst
	Wscript.Echo "Get DAT"
	R_FLDS=0
	redim flds(10,R_FLDS)
	do while not rs2.eof
		for i=0 to 9			
			flds(i,R_FLDS)=rs2(i)
		next
		flds(10,R_FLDS)=typ
		R_FLDS=R_FLDS+1
		redim Preserve flds(10,R_FLDS)
		if not rs2.eof then rs2.movenext
	loop
	rs2.close
	rs1.close
	rs1.open "SElect * from TERMINOWOŒÆ",objConnection,0,3
	for i=0 to R_FLDS-1
		Wscript.Echo "Dodajê rekord do bazy wiersz :" & i
		rs1.addnew
		'Wscript.Echo "check"
		for j=0 to 10
			'Wscript.Echo flds(j,i)
			rs1(j)=flds(j,i)
		next
		'Wscript.Echo "END check"
		rs1.update
		'Wscript.Echo "UPDT"
	next
end if
rs1.close
typ="BIE¯¥CY DZIEÑ"
rs1.open "select * from TERMINOWOŒÆ where refr='" & teraz & "' and Typ_rap='BIE¯¥CY DZIEÑ'"
if rs1.eof then	
	rs2.open replace(replace(replace(conq,"_teraz_",teraz),"_TYP_",typ),"(SELECT ifsapp.work_time_calendar_api.Get_Previous_Work_Day('SITS',SYSDATE) FROM dual)","(SELECT ifsapp.work_time_calendar_api.Get_Previous_Work_Day('SITS',SYSDATE+1) FROM dual)")
	rs2.movefirst
	R_FLDS=0
	redim flds(10,R_FLDS)
	do while not rs2.eof
		for i=0 to 10
			flds(i,R_FLDS)=rs2(i)
		next
		R_FLDS=R_FLDS+1
		redim Preserve flds(10,R_FLDS)
		if not rs2.eof then rs2.movenext
	loop
	rs2.close
	rs1.close
	rs1.open "SElect * from TERMINOWOŒÆ",objConnection,0,3
	for i=0 to R_FLDS-1
		Wscript.Echo "Dodajê rekord do bazy wiersz :" & i
		rs1.addnew
		'Wscript.Echo "check"
		for j=0 to 10
			'Wscript.Echo flds(j,i)
			rs1(j)=flds(j,i)
		next
		'Wscript.Echo "END check"
		rs1.update
		'Wscript.Echo "UPDT"
	next
end if
rs1.close
set objConnection1=nothing
Set objConnection1 = CreateObject("ADODB.Connection")
objConnection1.Open (conIMP)
rs2.CursorLocation = 3
rs2.LockType=4
Set rs2.ActiveConnection = objConnection1
objConnection.close
objConnection.Open (connect)
rs1.CursorLocation = 3
rs1.LockType=1
Set rs1.ActiveConnection = objConnection
Set fs = CreateObject("Scripting.FileSystemObject")
Set ts = fs.OpenTextFile("D:\boardinfo\SQLTXT.txt")
txt = ts.ReadAll
ts.Close
rs1.open txt
Set rs1.ActiveConnection = Nothing
objConnection.close
Set objConnection = Nothing
rs2.open "select * from TERMin"
Set rs2.ActiveConnection = Nothing
objConnection1.close
Wscript.Echo "No to do roboty..."
Wscript.Echo "Usuwam niepotrzebne indeksy..."
'Kasujemy w rs2 rekordy nie wystêpuj¹ce w rs1
if not rs2.eof then rs2.movefirst
rs1.movefirst
if not rs2.eof then
	active_ind=rs2("ID")
	rs2.filter="ID='" & active_ind & "'"
end if
Do While not rs2.EOF
	if active_ind<>rs2("iD") then
		active_ind=rs2("iD2")
			'Wscript.Echo "Sprawdzam :" & active_ind
		rs1.filter="iD='" & active_ind & "'"
		if rs1.eof or rs1.bof then
			Wscript.Echo "KASUJE :" & active_ind
			rs2.delete
		end if
	end if
	rs2.movenext
loop
Wscript.Echo "AKtualizujê zapisy ... Usuwanie bilansów z nieistniej¹c¹ dat¹"
rs1.filter=0
'rs1.close
'Kasujemy rekordy z indeksami nie w dacie
R_FLDS=0
redim flds(9,R_FLDS)
'rs1.open "Select * FROM TMP_DMD order by part_no,work_day ",objConnection,0,3
if not rs2.eof then rs2.movefirst
rs1.movefirst
active_ind=rs1("iD")
maxrec=rs1.RecordCount
rs2.filter="iD='" & active_ind & "'"
Do While not rs1.EOF and rs1.AbsolutePosition<maxrec+1
	if active_ind<>rs1("iD") then
	'zmieni³ siê indeks
		active_ind=rs1("iD")
		rs2.filter="iD='" & active_ind & "'"
		'Wscript.Echo "Sprawdzam indeks " & active_ind & " " & R_FLDS
	end if
If not rs2.eof then
	if rs1("IL_plan").value<>cdbl(rs2("IL_plan").value) then rs2("IL_plan").value=cdbl(rs1("IL_plan").value)
	if rs1("MADE_ON_TIME").value<>cdbl(rs2("MADE_ON_TIME").value) then rs2("MADE_ON_TIME").value=cdbl(rs1("MADE_ON_TIME").value)
	if rs1("MADE_TOO_LATE").value<>cdbl(rs2("MADE_TOO_LATE").value) then rs2("MADE_TOO_LATE").value=cdbl(rs1("MADE_TOO_LATE").value)
	if rs1("TERMIN").value<>cdbl(rs2("TERMIN").value) then rs2("TERMIN").value=cdbl(rs1("TERMIN").value)
else
	if rs1.AbsolutePosition<maxrec+1 then
		Wscript.Echo "Dodajê rekord " & active_ind & " " & R_FLDS  & " " & rs1("iD")
		for i=0 to 9
			flds(i,R_FLDS)=rs1(i)
		next
		R_FLDS=R_FLDS+1
		redim Preserve flds(9,R_FLDS)
	end if
end if
if not rs1.eof then rs1.movenext
if rs1("iD") is nothing or rs1.eof then exit do
loop
rs2.filter=0
Wscript.Echo rs2.RecordCount
rs1.close
set rs1=nothing
'Dodajê rekordy z indeksami
'rs1.open "Select * FROM TMP_DMD order by part_no,work_day ",objConnection,0,3
for i=0 to R_FLDS-1
	Wscript.Echo "Dodajê rekord do bazy wiersz :" & i
	rs2.addnew
	for j=0 to 9
		rs2(j)=flds(j,i)
	next
	rs2.update
next
rs2.movefirst
objConnection1.Open (conIMP)
Set rs2.ActiveConnection = objConnection1
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open (connect)
Wscript.Echo "UPDATE "
rs2.updatebatch
Wscript.Echo rs2.RecordCount
rs2.close
set rs2=nothing
objConnection1.close

set objConnection1=nothing
