DIM flds()
Dim connect
connect="Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=\\10.0.1.22\sys\KatPryw\RADKOS\History.mdb;UID=Admin;PWD= ;"
Set objConnection = CreateObject("ADODB.Connection")
Set rs1 = CreateObject("ADODB.Recordset")
objConnection.Open (connect)
rs1.CursorLocation = 3
rs1.LockType=3
Set rs1.ActiveConnection = objConnection
teraz=year(now())
if len(cstr(month(now())))=1 then teraz=teraz & "0" & cstr(month(now())) else teraz=teraz & cstr(month(now()))
if len(cstr(day(now())))=1 then teraz=teraz & "0" & cstr(day(now())) else  teraz=teraz & cstr(day(now()))
if len(cstr(Hour(now())))=1 then teraz=teraz & "0" & cstr(Hour(now())) else  teraz=teraz & cstr(Hour(now()))
Wscript.Echo teraz
rs1.open "select * from hist_brak where refr=" & teraz
if rs1.eof then
	rs1.close 
	rs1.open "SELECT TOP 4 TMP_DMD.WORK_DAY FROM TMP_DMD GROUP BY TMP_DMD.WORK_DAY"
	rs1.movelast
	a=rs1(0).value
	rs1.close
	rs1.open "SELECT TMP_DMD.KOOR, DateValue(TMP_DMD.WORK_DAY) AS WORK_DAY, TMP_DMD.RODZAJ, Sum(TMP_DMD.BIL_MAG) AS SUM_BRAK, Sum(TMP_DMD.BIL_ZAM_MAG) AS SUM_BRAK_PROG, Sum(TMP_DMD.QTY_DEMAND) AS SUM_ALL, '" & teraz & "' AS REFR, IIf(IsNull([BRAK]),0,[BRAK]) AS PROD_SHORT,DAY_QTY1.QTY_ALL AS PROD_ALL, DAY_QTY1.SHORTage AS ALL_SHORTAGE FROM (SELECT WORK_DAY, iif(TYP='MRP','IKEA','SITS') AS type, QTY_ALL, SHORTage FROM DAY_QTY1)  AS DAY_QTY1 INNER JOIN (TMP_DMD LEFT JOIN (SELECT Brak_KOOR1.dat AS WORK_DAY, Brak_KOOR1.PART_BUYER,iif(Brak_KOOR1.PROD_TYPE='MRP','IKEA','SITS') AS TYPE, Brak_KOOR1.BRAK FROM Brak_KOOR1)  AS Braki_KOOR ON (TMP_DMD.KOOR = Braki_KOOR.PART_BUYER) AND (TMP_DMD.WORK_DAY = Braki_KOOR.WORK_DAY) AND (TMP_DMD.RODZAJ = Braki_KOOR.TYPE)) ON (DAY_QTY1.TYPe = TMP_DMD.RODZAJ) AND (DAY_QTY1.WORK_DAY = TMP_DMD.WORK_DAY) GROUP BY TMP_DMD.KOOR, DateValue(TMP_DMD.WORK_DAY), TMP_DMD.RODZAJ, IIf(IsNull([BRAK]),0,[BRAK]), DAY_QTY1.QTY_ALL, DAY_QTY1.SHORTage HAVING (((TMP_DMD.KOOR)<>'*' And (TMP_DMD.KOOR)<>'LUCPRZ') AND ((DateValue([TMP_DMD].[WORK_DAY]))<=#" & a & "#))"
	rs1.movefirst
	R_FLDS=0
	redim flds(9,R_FLDS)
	do while not rs1.eof
		for i=0 to 9
			flds(i,R_FLDS)=rs1(i)
		next
		R_FLDS=R_FLDS+1
		redim Preserve flds(9,R_FLDS)
		if not rs1.eof then rs1.movenext
	loop
	rs1.close
	rs1.open "SElect * from hist_brak",objConnection,0,3
	for i=0 to R_FLDS-1
		Wscript.Echo "Dodajê rekord do bazy wiersz :" & i
		rs1.addnew
		'Wscript.Echo "check"
		for j=0 to 9
			'Wscript.Echo flds(j,i)
			rs1(j)=flds(j,i)
		next
		'Wscript.Echo "END check"
		rs1.update
		'Wscript.Echo "UPDT"
	next
end if
rs1.close
objConnection.close
set rs1=nothing
set objConnection=nothing