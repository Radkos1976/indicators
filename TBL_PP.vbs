DIM flds()
DIM flds1()
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
		if len(cstr(month(rs1("Work_day"))))=1 then teraz1=teraz1 & "0" & cstr(month(rs1("Work_day"))) else teraz1=teraz1 & cstr(month(rs1("Work_day")))
		if len(cstr(day(rs1("Work_day"))))=1 then teraz1=teraz1 & "0" & cstr(day(rs1("Work_day"))) else  teraz1=teraz1 & cstr(day(rs1("Work_day")))
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
			Wscript.Echo "PE£NY BRAK     :" & rs1(0)  & " DATA " & teraz1
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
Wscript.Echo "URUCHAMIAM KWERENDY W ORACLE - PE£NY BRAK"
for i=4 to 6 step 2
	if len(tbl(i,1))>3 then
	Wscript.Echo tbl(i,0)
	if i=0 then label="Decode(Sign(DATE_REQUIRED-To_Date(to_char(SYSDATE,'YYYY-MM-DD'))),'-1',To_Date(to_char(SYSDATE,'YYYY-MM-DD')),DATE_REQUIRED)" else label="DATE_REQUIRED"
rs2.open "SELECT part_no,DATE_REQUIRED,ifsapp.inventory_part_api.Get_Planner_Buyer('ST',a.part_no) Part_Buyer,Decode(ORDER_SUPPLY_DEMAND_TYPE,'Rez. mat. ZP',ifsapp.shop_ord_api.Get_State(a.order_no,a.line_no,a.rel_no)||'//'||Nvl(ifsapp.shop_order_operation_api.Get_Work_Group_Id(a.order_no,a.line_no,a.rel_no,1),Nvl(ifsapp.dop_head_api.Get_C_Trolley_Id(SubStr(ifsapp.shop_ord_api.Get_Source(a.order_no,a.line_no,a.rel_no),10,InStr(ifsapp.shop_ord_api.Get_Source(a.order_no,a.line_no,a.rel_no),'^',10)-10)),' '))||'//', STATUS_DESC)||' //'||ifsapp.shop_order_operation_list_api.Get_Next_Op_Work_Center(a.order_no,a.line_no,a.rel_no,0) STAT ,Nvl(Decode(ORDER_SUPPLY_DEMAND_TYPE,'Rez. mat. ZP',SubStr(ifsapp.shop_ord_api.Get_Source(a.order_no,a.line_no,a.rel_no),10,InStr(ifsapp.shop_ord_api.Get_Source(a.order_no,a.line_no,a.rel_no),'^',10)-10),'Potrzeby DOP',a.ORDER_NO),a.ORDER_NO) DOP_ord,Nvl(Decode(ORDER_SUPPLY_DEMAND_TYPE,'Rez. mat. ZP',REPLACE(SubStr(ifsapp.shop_ord_api.Get_Source(a.order_no,a.line_no,a.rel_no),InStr(ifsapp.shop_ord_api.Get_Source(a.order_no,a.line_no,a.rel_no),'^',10)),'^',''),'Potrzeby DOP',a.LINE_NO),a.info)  DOP_lin,a.ORDER_NO,ORDER_SUPPLY_DEMAND_TYPE, QTY_DEMAND, ifsapp.inventory_part_in_stock_api. Get_Plannable_Qty_Onhand('ST',part_no,'*') mag,Nvl(ifsapp.shop_ord_api.Get_Revised_Qty_Due(a.order_no,a.line_no,a.rel_no),0) Prod_QTY,to_date('" & tbl(i,0) & "','yyyymmdd') RPT_DAT,ifsapp.inventory_part_api.Get_Description('ST',a.part_no) Short_nam,Decode(ORDER_SUPPLY_DEMAND_TYPE,'Rez. mat. ZP',Decode(Sign(ifsapp.shop_ord_api.Get_Revised_Due_Date(a.order_no,a.line_no,a.rel_no)-SYSDATE),'-1',To_Date(to_char(SYSDATE,'YYYY-MM-DD')),ifsapp.shop_ord_api.Get_Revised_Due_Date(a.order_no,a.line_no,a.rel_no)),to_date('" & tbl(i,0) & "','yyyymmdd')) realizacja FROM ifsapp.order_supply_demand_ext a where ORDER_SUPPLY_DEMAND_TYPE not in ('Zapotrzeb. zakupu','Potrzeby zaplan. w MRP','Zam. zakupu','Zadanie transportu') AND part_no in " & tbl(i,1) & " AND " & label & "=(select to_date('" & tbl(i,0) & "','yyyymmdd') from dual) ORDER BY DATE_REQUIRED,order_no"
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
else
rs1.close
set rs2=nothing
set objConnection1=nothing
end if	