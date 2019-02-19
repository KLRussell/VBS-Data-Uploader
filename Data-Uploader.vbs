const My_SQL_Server = ""
const My_SQL_DB = ""
const VCD_TBL = ""
const VCD_DIR = ""
const VCD_ODS_COLs = "Initiated_By, Batch_ID, VTN, Claim_Line_Num, Claim_Status, Escalation, Claim_Type, Sub_Claim_ID, STC_Claim_Number, Received_Date, Resolution_Date, Completion_Date, Cust_Corp_ID, Company_Name, State, BAN, Billing_System, Bill_Date, Dispute_Amount, Approved_Amount, Denied, Total_Amount, WTN, USOC, PON, Dispute_Reason, Resolution"
const VCD_Excel_COLs = "INITIATED BY DSCR, BATCH_ID, CLAIM_TRACKING_NUMBER, CLAIM_LINE_NUMBER, CLAIM STATUS, ESCALATION, CLAIM TYPE, SUBCLAIM ID, CUSTOMER_CLAIM_NUMBER, RECEIVED_DATE, RESOLUTION_DATE, COMPLETION_DATE, CUSTCORP_ID, COMPANY_NAME, STATE, BAN, BILLING SYSTEM, BILL_DATE, LINE ITEM CLAIM AMOUNT, AMOUNT APPROVED, AMOUNT DENIED, TOTAL CLAIM AMOUNT, CIRCUIT ID OR TN, USOC, PON_ASR_LSR, CLAIM_DESCRIPTION, RESOLUTION"

Public WshShell, oFSO, Log_Path, networkInfo, VCD_File_Path, My_List, My_SQL, SQL_Filepath
dim Temp, myquery, oFolder, oFileCollection, oFile, myitems(1), Q, mydata, SQL_List(4), varresponse

Set WshShell = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set networkInfo = CreateObject("WScript.NetWork")


Set oFolder = oFSO.GetFolder(VCD_DIR)
Set oFileCollection = oFolder.Files

For each oFile in oFileCollection
	if oFSO.getextensionname(Cstr(oFile.Name)) = "xlsx" then
		if len(myitems(0)) > 0 then
			if myitems(0) < oFile.DateCreated then
				myitems(0) = oFile.DateCreated
				myitems(1) = oFile.Name
			end if
		else
			myitems(0) = oFile.DateCreated
			myitems(1) = oFile.Name
		end if
	end if
Next

varresponse = MsgBox("Warning! System assumes to append (" & myitems(1) & "). Would you like to proceed?", vbYesNo, "Append Data")
If varresponse <> vbYes Then
	myitems(1) = empty
	do while isempty(myitems(1))
		varresponse = inputbox("Please type in the filename of the file you would like to append? Make sure to include the file extension","Append Data Manually")
		if isempty(varresponse) then
			wscript.quit()
		else
			if oFSO.fileexists(VCD_DIR & varresponse) then
				myitems(1) = varresponse
			end if
		end if
	loop
end if

set oFolder = nothing
Set oFileCollection = Nothing
Set oFile = Nothing

Log_Path = replace(WScript.ScriptFullName,".vbs","") & "_Log.txt"

if len(myitems(1)) > 0 then
	VCD_File_Path = VCD_DIR & myitems(1)

	My_List = Excel_Records(VCD_File_Path)

	if IsArray(My_List) then
		SQL_Filepath = oFSO.GetParentFolderName(WScript.ScriptFullName) & "\SQL_Script\VTN_Upload_Cleanup.sql"
		Q = Ceil(ubound(My_List, 1) / 1000) - 1
		redim My_SQL(Q)

		for Q = lbound(My_SQL, 1) to ubound(My_SQL, 1)
			if ((Q + 1) * 1000) - 1 > (ubound(My_List, 1)) then
				Append_Records (Q * 1000), (ubound(My_List, 1)), mydata
			else
				Append_Records (Q * 1000), ((Q + 1) * 1000) - 1, mydata
			end if
			My_SQL(Q) = mydata
		Next
		set My_List = nothing

		for Q = lbound(My_SQL, 1) to ubound(My_SQL, 1)
			SQL_List(Q mod 5) = My_SQL(Q)
			if Q > 3 and Q mod 5 = 4 then
				Append_ODS join(SQL_List, chr(13))
				SQL_List(0) = ""
				SQL_List(1) = ""
				SQL_List(2) = ""
				SQL_List(3) = ""
				SQL_List(4) = ""
			end if
		next

		if not Q mod 5 = 0 then
			Append_ODS join(SQL_List, chr(13))
		end if

		set My_SQL = nothing

		Append_ODS "update " & VCD_TBL & " set EDIT_DATE=getdate() where EDIT_DATE is null"

		Execute_SQL

		msgbox("Upload is now Finished!")
	else
		write_log Now() & " * Error * " & networkInfo.UserName & " * No Data found or workbook has invalid column format (" & myitems(1) & ")"
	end if

else
	write_log Now() & " * Warning * " & networkInfo.UserName & " * User chose to not append data. Exiting"
end if

function Excel_Records(Filepath)
	if len(filepath) > 0 then
		if oFSO.fileexists(filepath) then
                	Dim myExcelWorker, oWorkBook, strSaveDefaultPath
			Dim My_Header, Header_Master, My_Data, My_Items, dim1, dim2, Q, R, Line_CHK
            
                	Set myExcelWorker = CreateObject("Excel.Application")
                	strSaveDefaultPath = myExcelWorker.DefaultFilePath
            
                	myExcelWorker.DisplayAlerts = False
                	myExcelWorker.AskToUpdateLinks = False
                	myExcelWorker.AlertBeforeOverwriting = False
                	myExcelWorker.FeatureInstall = msoFeatureInstallNone
            
                	myExcelWorker.DefaultFilePath = oFSO.GetParentFolderName(Filepath)
            
                	On Error Resume Next
            
                	Set oWorkBook = myExcelWorker.Workbooks.Open(Filepath)

                	If Err.Number <> 0 Then
                    		write_log Now() & " * Error * " & networkInfo.UserName & " * Open Workbook (" & Err.Description & ")"
            
                    		myExcelWorker.DefaultFilePath = strSaveDefaultPath
                    		myExcelWorker.Quit
                    		Set myExcelWorker = Nothing
                    		Set oWorkBook = Nothing
                    		Exit function
                	End If

			oworkbook.Application.ScreenUpdating = FALSE
			oworkbook.Application.EnableEvents = FALSE

			My_Header = oworkbook.sheets(1).cells(1, 1).Resize(1, oworkbook.sheets(1).UsedRange.Columns.Count)
			My_Data = oworkbook.sheets(1).cells(2, 1).Resize(oworkbook.sheets(1).UsedRange.Rows.Count - 1, oworkbook.sheets(1).UsedRange.Columns.Count)

			if isarray(My_Header) and isarray(My_Data) then
				Header_Master = split(VCD_Excel_COLs,", ")
				dim1 = ubound(My_Data,1) - lbound(My_Data,1)
				dim2 = ubound(My_Data,2) - lbound(My_Data,2)

					redim My_List(dim1, ubound(Header_Master)) 

					for Q = lbound(Header_Master) to ubound(Header_Master)
						Line_CHK = TRUE

						for R = lbound(My_Header, 2) to ubound(My_Header, 2)
							if lcase(replace(Header_Master(Q),"_"," ")) = lcase(replace(My_Header(1, R),"_"," ")) then
								Line_CHK = FALSE

								for S = lbound(My_Data, 1) to ubound(My_Data)
									my_list(S-1, Q) = My_Data(S, R)
								next

								exit for
							end if
						next

						if Line_CHK then
							write_log Now() & " * Error * " & networkInfo.UserName & " * " & lcase(replace(Header_Master(Q),"_"," ")) & " was not found in (" & oFSO.GetParentFolderName(Filepath) & ")"

							oworkbook.Application.ScreenUpdating = TRUE
							oworkbook.Application.EnableEvents = TRUE

                					oWorkBook.Close

							myExcelWorker.DefaultFilePath = strSaveDefaultPath
            
                					If Err.Number <> 0 Then
                    						err.clear
                    						myExcelWorker.Quit
                					End If

                					Set myExcelWorker = Nothing
                					Set oWorkBook = Nothing

							exit function
						end if
					next

					Excel_Records = My_List
                		
			end if

			oworkbook.Application.ScreenUpdating = TRUE
			oworkbook.Application.EnableEvents = TRUE

                	oWorkBook.Close

			myExcelWorker.DefaultFilePath = strSaveDefaultPath
            
                	If Err.Number <> 0 Then
                    		err.clear
                    		myExcelWorker.Quit
                	End If

                	Set myExcelWorker = Nothing
                	Set oWorkBook = Nothing
		end if
	else
		write_log Now() & " * Error * " & networkInfo.UserName & " * File does not exist (" & oFSO.GetFileName(filepath) & ")"
	end if
end function

sub Append_Records(mystart, myend, mydata)
	on error resume next
	dim My_Records, My_Temp, S, T, myline, Temp_List, myitem

	S = cint(ubound(My_List, 2))

	redim My_Temp(S)
	S = myend - mystart
	redim My_Records(S)

	for S = mystart to myend
		myline = S - mystart
		for T = lbound(My_List, 2) to ubound(My_List, 2)
			myitem = My_List(S, T)

			if isnull(myitem) or isempty(myitem) then
				My_Temp(T) = "NULL"
			else
				My_Temp(T) = replace(trim(myitem),"'","''")
			end if
		next
		if isarray(My_Temp) then
			My_Records(myline) = "(" & replace("'" & join(My_Temp, "', '") & "'", "'NULL'", "NULL") & ")"
		end if
	next
	mydata = "insert into " & VCD_TBL & " (" & VCD_ODS_COLS & ") values " & join(My_Records, ",") & ";"
end sub

Sub Append_ODS(myquery)
	On Error Resume Next
	Dim constr, conn

	constr = "Provider=SQLOLEDB;Data Source=" & My_SQL_Server & ";Initial Catalog=" & My_SQL_DB & ";Integrated Security=SSPI;"
	Set conn = CreateObject("ADODB.Connection")

	conn.Open constr

	If Err.Number <> 0 Then
		write_log Now() & " * Error * " & networkInfo.UserName & " * Open SQL Conn (" & Err.Description & ")"
		Set conn = Nothing
		exit sub
	end if

	conn.CommandTimeout = 0

	conn.Execute myquery

	If Err.Number <> 0 Then
		write_log Now() & " * Error * " & networkInfo.UserName & " * SQL Execute Query (" & Err.Description & ")"
		Set conn = Nothing
		exit sub
	end if
    
	conn.Close

	If Err.Number <> 0 Then
		write_log Now() & " * Error * " & networkInfo.UserName & " * SQL Close Con (" & Err.Description & ")"
	end if
    
	Set conn = Nothing
End Sub

Sub Execute_SQL()
	Dim objFile, strLine

	if oFSO.fileexists(SQL_Filepath) then
		Set objFile = oFSO.OpenTextFile(SQL_Filepath)
		Do Until objFile.AtEndOfStream
			if len(strLine) > 0 then
				strLine= strLine & vbcrlf & objFile.ReadLine
			else
    				strLine= objFile.ReadLine
			end if
		Loop
		objfile.close

		Append_ODS strLine
	else
		write_log Now() & " * Error * " & networkInfo.UserName & " * SQL Script (" & SQL_Filepath & ") does not exist"
	end if

	set objFile = Nothing
end sub

Sub Write_Log(ByVal text)
	Dim objFile, strLine

	if oFSO.fileexists(Log_Path) then
		Set objFile = oFSO.OpenTextFile(Log_Path)
		Do Until objFile.AtEndOfStream
			if len(strLine) > 0 then
				strLine= strLine & vbcrlf & objFile.ReadLine
			else
    				strLine= objFile.ReadLine
			end if
		Loop
		objfile.close
		Set objfile = oFSO.CreateTextFile(Log_Path,True)
		objfile.write strLine & vbcrlf & text
		objfile.close
	else

		Set objfile = oFSO.CreateTextFile(Log_Path,True)
		objfile.write text
		objfile.close
	end if

	msgbox(text)

	set objFile = Nothing
End Sub

Function Ceil(x)
    If Round(x) = x Then
        Ceil = x
    Else
        Ceil = Round(x + 0.5)
    End If
End Function

Function IsArray(anArray)
    Dim I
    On Error Resume Next
    I = UBound(anArray, 1)
    If Err.Number = 0 Then
        IsArray = True
    Else
        IsArray = False
    End If
End Function
