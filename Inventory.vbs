Option Explicit
On Error Resume Next 

Dim objWMIService, objFSO
Dim colPingResults, colSMBIOS, colLogicalDrives, colDiskDrives, colCompSys, colNetAdapters
Dim objPingResult, objSMBIOS, objLogicalDrive, objDiskDrive, objCompSys, objNetAdapter
Dim sqlServerIP, sqlInstance, sqlDatabase, sqlUser, sqlPassword
Dim sqlComputerTable, sqlNetworkTable, sqlLogicalDriveTable, sqlDiskDriveTable
Dim strSQL, conn, rs, strSQLConn, bolSQLInsert

Dim strComputerName(), strSerialNumber(), strBIOSAssetTag()
Dim strLogicalName(), strLogicalDesc(), strLogicalFS(), strlogicalSize(), strLogicalFree()
Dim strDiskName(), strDiskModel(), strDiskSize()
Dim strCompModel(), strCompManuf(), strTotalRam()
Dim strNetworkName(), strNetworkManuf(), strMACAddress(), strNetworkType()
Dim strComputer, i, j

If WScript.Arguments.Count > 0 Then
	strComputer = WScript.Arguments.Item(0)
	If Left(strComputer,1) = "\" or Left(strComputer,1) = "-" Then
		strComputer = Mid(strComputer,2)
	End If
	'wscript.echo(strComputer)
Else 
	strComputer = "."
End If

sqlServerIP = "IPADDRESS"
sqlInstance = "INSTANCE"
sqlDatabase = "DATABASE"
sqlUser = "SQLUSER"
sqlPassword = "SQLPASS"
sqlComputerTable = "ComputerInfo"
sqlNetworkTable = "NetworkInfo"
sqlLogicalDriveTable = "LogicalDriveInfo"
sqlDiskDriveTable = "DiskDriveInfo"

Set objWMIService = GetObject("WinMgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colPingResults = objWMIService.ExecQuery ("SELECT * FROM Win32_PingStatus WHERE Address = '" & SQLServerIP & "'")

For Each objPingResult In colPingResults
	Select Case True
		Case Not IsObject(objPingResult) : wscript.quit
		Case objPingResult.StatusCode = 0
		Case Else : wscript.quit
	End Select
Next

Set colPingResults = Nothing
Set Conn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.RecordSet")

strSQLConn = "Provider=SQLOLEDB.1;Password=" & SQLPassword & ";Persist Security Info=True;" &_
				"User ID=" & SQLUser & ";Initial Catalog=" & SQLDatabase & ";" &_
				"Data Source=" & SQLServerIP & "\" & SQLInstance

Conn.Open strSQLConn

If Err Then wscript.quit

Set colSMBIOS = objWMIService.ExecQuery ("Select * from Win32_SystemEnclosure")
Set colLogicalDrives = objWMIService.ExecQuery ("Select * from Win32_LogicalDisk Where Description = 'Local Fixed Disk'")
Set colDiskDrives = objWMIService.ExecQuery ("Select * from Win32_DiskDrive")
Set colCompSys = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
Set colNetAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter Where AdapterType = 'Ethernet 802.3'")

i = 0

For Each objSMBIOS in colSMBIOS
	With objSMBIOS
		ReDim Preserve strComputerName(i) : strComputerName(i) = .Path_.Server
		ReDim Preserve strSerialNumber(i) : strSerialNumber(i) = .SerialNumber
		ReDim Preserve strBIOSAssetTag(i) : strBIOSAssetTag(i) = .SMBIOSAsSetTag
	End With
	REM MsgBox _
		REM "Computer Name: " & strComputerName(i) & VbCr &_
		REM "Serial Number: " & strSerialNumber(i) & VbCr &_
		REM "Asset Tag: " & strBIOSAssetTag(i)
	i = i + 1
Next

i = 0

For Each objLogicalDrive in colLogicalDrives
	With objLogicalDrive
		ReDim Preserve strLogicalName(i) : strLogicalName(i) = .Name
		ReDim Preserve strLogicalDesc(i) : strLogicalDesc(i) = .Description
		ReDim Preserve strLogicalFS(i) : strLogicalFS(i) = .FileSystem
		ReDim Preserve strlogicalSize(i) : strlogicalSize(i) = .size
		ReDim Preserve strLogicalFree(i) : strLogicalFree(i) = .FreeSpace
	End With
	REM MsgBox _
		REM "Logical Drive Name: " & strLogicalName(i) & VbCr &_
		REM "Logical Drive Description: " & strLogicalDesc(i) & VbCr &_
		REM "Logical Drive File System: " & strLogicalFS(i) & VbCr &_
		REM "Logical Drive Total Size: " & ByteSize(strlogicalSize(i)) & VbCr &_
		REM "Logical Drive Free Space: " & ByteSize(strLogicalFree(i))
	i = i + 1
Next

i = 0

For Each objDiskDrive in colDiskDrives
	With ObjDiskDrive
		ReDim Preserve strDiskName(i) : strDiskName(i) = .Name
		ReDim Preserve strDiskModel(i) : strDiskModel(i) = .Model
		ReDim Preserve strDiskSize(i) : strDiskSize(i) = .size
	End With
	REM msgbox _
		REM "Disk Drive Name: " & strDiskName(i) & VbCr &_
		REM "Disk Drive Model: " & strDiskModel(i) & VbCr &_
		REM "Disk Drive Size: " & ByteSize(strDiskSize(i))
	i = i + 1
Next

i = 0

For Each objCompSys in colCompSys
	With objCompSys
		ReDim Preserve strCompModel(i) : strCompModel(i) = .Model
		ReDim Preserve strCompManuf(i) : strCompManuf(i) = .Manufacturer
		ReDim Preserve strTotalRam(i) : strTotalRam(i) = .TotalPhysicalMemory
	End With
	REM msgbox _
		REM "Computer Model :"  & strCompModel(i) & VbCr &_
		REM "Computer Manufacturer: " & strCompManuf(i) & VbCr &_
		REM "Computer Total Ram: " & ByteSize(strTotalRam(i))
	i = i + 1
Next

i = 0

For Each objNetAdapter in colNetAdapters
	With objNetAdapter
		ReDim Preserve strNetworkName(i) : strNetworkName(i) = .Name
		ReDim Preserve strNetworkManuf(i) : strNetworkManuf(i) = .Manufacturer
		ReDim Preserve strMACAddress(i) : strMACAddress(i) = .MACAddress
		ReDim Preserve strNetworkType(i) : strNetworkType(i) = .AdapterType 
	End With
	REM msgbox _
		REM "Network Adapter Name: " & strNetworkName(i) & VbCr &_
		REM "Network Adapter Manufacturer: " & strNetworkManuf(i) & VbCr &_
		REM "Network Adapter MACAddress: " & strMACAddress(i) & VbCr &_
		REM "Network Adapter Type: " & strNetworkType(i)
	i = i + 1
Next

Set colNetAdapters = Nothing
Set colCompSys = Nothing
Set colDiskDrives = Nothing
Set colLogicalDrives = Nothing
Set colSMBIOS = Nothing

For i = LBound(strComputerName) To UBound(strComputerName)
	' Clean up and update the Computer Table
	bolSQLInsert = True
	strSQL = "Select SerialNumber, TotalRam From " & SQLComputerTable & " Where ComputerName = '" &  strComputerName(i) & "'"
	rs.open strSQL, Conn, 3, 3
	rs.MoveFirst
	While Not rs.EOF
		If inArray(strSerialNumber,rs("SerialNumber")) = -1 Then
			strSQL = "DELETE FROM " & SQLComputerTable & " WHERE SerialNumber = '" & rs("SerialNumber") & "' AND ComputerName = '" &  strComputerName(i) & "'"
			Conn.Execute(strSQL)
		Else 
			bolSQLInsert = False
			If inArray(strTotalRam,rs("TotalRam")) = -1 Then
				strSQL = "UPDATE " & SQLComputerTable & " SET TotalRam = '" & strTotalRam(i) & "' WHERE TotalRam = '" & rs("TotalRam") &_
					"' AND ComputerName = '" &  strComputerName(i) & "' AND SerialNumber = '" & rs("SerialNumber") & "'"
				Conn.Execute(strSQL)
			End If
		End If
		rs.MoveNext
	Wend
	If bolSQLInsert = True Then
		For j = LBound(strTotalRam) to UBound(strTotalRam)
			strSQL = "INSERT INTO " & SQLComputerTable & "(ComputerName,SerialNumber,BiosAsset,ComputerModel,ComputerManufacturer,TotalRam) " &_
					"VALUES ('" & strComputerName(i) & "' ,'" & strSerialNumber(i) & "' ,'" & strBIOSAssetTag(i) & "' ,'" &_
					strCompModel(j) & "' ,'" & strCompManuf(j) & "' ,'" & strTotalRam(j) & "')"
			Conn.Execute(strSQL)
		Next
	End If
	rs.close
	' Clean up and update the Network Table
	strSQL = "Select MacAddress From " & SQLNetworkTable & " Where ComputerName = '" &  strComputerName(i) & "'"
	rs.open strSQL, Conn, 3, 3
	rs.MoveFirst
	While Not rs.EOF
		If inArray(strMACAddress,rs("MacAddress")) = -1 Then
			strSQL = "DELETE FROM " & SQLNetworkTable & " WHERE MacAddress = '" & rs("MacAddress") & "' AND ComputerName = '" &  strComputerName(i) & "'"
			Conn.Execute(strSQL)
		End If
		rs.MoveNext
	Wend
	rs.close
	For j = LBound(strMACAddress) to UBound(strMACAddress)
		strSQL = "Select ComputerName, MacAddress From " & SQLNetworkTable & " WHERE MacAddress = '" & strMACAddress(j) & "'"
		rs.open strSQL, Conn, 3, 3
		If inRecordSet(rs, "MacAddress", strMACAddress(j)) = False Then
			strSQL = "INSERT INTO " & SQLNetworkTable & "(ComputerName,AdapterName,AdapterManufacturer,MacAddress,AdapterType) " &_
				"VALUES ('" & strComputerName(i) & "', '" & strNetworkName(j) & "', '" & strNetworkManuf(j) & "', '" & strMACAddress(j) & "', '" & strNetworkType(j) & "')"
			Conn.Execute(strSQL)
		ElseIf getRecordSet(rs, "MacAddress", strMACAddress(j), "ComputerName") <> strComputerName(i) Then
			strSQL = "UPDATE " & SQLNetworkTable & " " &_
				"SET ComputerName = '" & strComputerName(i) & "' " & "WHERE MacAddress = '" & strMacAddress(j) & "'"
			Conn.Execute(strSQL)
		End If
		rs.close
	Next
	' Clean up and update the Logical Drive Table
	strSQL = "Select LogicalDriveName From " & SQLLogicalDriveTable & " Where ComputerName = '" &  strComputerName(i) & "'"
	rs.open strSQL, Conn, 3, 3
	rs.MoveFirst
	While Not rs.EOF
		If inArray(strLogicalName,rs("LogicalDriveName")) = -1 Then
			strSQL = "DELETE FROM " & SQLLogicalDriveTable & " WHERE LogicalDriveName = '" & rs("LogicalDriveName") & "' AND ComputerName = '" &  strComputerName(i) & "'"
			Conn.Execute(strSQL)
		End If
		rs.MoveNext
	Wend
	rs.close
	strSQL = "Select * From " & SQLLogicalDriveTable & " WHERE ComputerName = '" & strComputerName(i) & "'"
	rs.open strSQL, Conn, 3, 3
	For j = LBound(strLogicalName) to UBound(strLogicalName)		
		If inRecordSet(rs, "LogicalDriveName", strLogicalName(j)) = False Then
			strSQL = "INSERT INTO " & SQLLogicalDriveTable &_
				"(ComputerName,LogicalDriveName,LogicalDriveDescription,LogicalDriveFS,LogicalDriveTotalSize,LogicalDriveFreeSpace) " &_
				"VALUES ('" & strComputerName(i) & "', '" & strLogicalName(j) & "', '" & strLogicalDesc(j) & "', '" & strLogicalFS(j) &_
				"', '" & strlogicalSize(j) & "', '" & strLogicalFree(j) & "')"
			Conn.Execute(strSQL)
		ElseIf (getRecordSet(rs, "LogicalDriveName", strLogicalName(j), "LogicalDriveDescription") <> strLogicalDesc(j)) _
			OR (getRecordSet(rs, "LogicalDriveName", strLogicalName(j), "LogicalDriveFS") <> strLogicalFS(j)) _
			OR (getRecordSet(rs, "LogicalDriveName", strLogicalName(j), "LogicalDriveTotalSize") <> strlogicalSize(j)) _
			OR (getRecordSet(rs, "LogicalDriveName", strLogicalName(j), "LogicalDriveFreeSpace") <> strLogicalFree(j)) Then
				strSQL = "UPDATE " & SQLLogicalDriveTable & " " &_
					"SET LogicalDriveDescription = '" & strLogicalDesc(j) & "', LogicalDriveFS = '" & strLogicalFS(j) & "', " &_
					"LogicalDriveTotalSize = '" & strlogicalSize(j) & "', LogicalDriveFreeSpace = '" & strLogicalFree(j) & "' " &_
					"WHERE LogicalDriveName = '" & strLogicalName(j) & "' AND ComputerName = '" & strComputerName(i) & "'"
				Conn.Execute(strSQL)
		End If
	Next
	rs.close
	' Clean up and update the Disk Drive Table
	strSQL = "Select DiskDriveName From " & SQLDiskDriveTable & " Where ComputerName = '" &  strComputerName(i) & "'"
	rs.open strSQL, Conn, 3, 3
	rs.MoveFirst
	While Not rs.EOF
		If inArray(strDiskName,rs("DiskDriveName")) = -1 Then
			strSQL = "DELETE FROM " & SQLDiskDriveTable & " WHERE DiskDriveName = '" & rs("DiskDriveName") & "' AND ComputerName = '" &  strComputerName(i) & "'"
			Conn.Execute(strSQL)
		End If
		rs.MoveNext
	Wend
	rs.close
	strSQL = "Select * From " & SQLDiskDriveTable & " WHERE ComputerName = '" & strComputerName(i) & "'"
	rs.open strSQL, Conn, 3, 3
	For j = LBound(strDiskName) to UBound(strDiskName)		
		If inRecordSet(rs, "DiskDriveName", strDiskName(j)) = False Then
			strSQL = "INSERT INTO " & SQLDiskDriveTable &_
				"(ComputerName,DiskDriveName,DiskDriveModel,DiskDriveSize) " &_
				"VALUES ('" & strComputerName(i) & "', '" & strDiskName(j) & "', '" & strDiskModel(j) & "', '" & strDiskSize(j) & "')"
			Conn.Execute(strSQL)
		ElseIf (getRecordSet(rs, "DiskDriveName", strDiskName(j), "DiskDriveModel") <> strDiskModel(j)) _
			OR (getRecordSet(rs, "DiskDriveName", strDiskName(j), "DiskDriveSize") <> strDiskSize(j)) Then
				strSQL = "UPDATE " & SQLDiskDriveTable & " SET DiskDriveModel = '" & strDiskModel(j) & "', DiskDriveSize = '" & strDiskSize(j) & "' " &_
					"WHERE DiskDriveName = '" & strDiskName(j) & "' AND ComputerName = '" & strComputerName(i) & "'"
				Conn.Execute(strSQL)
		End If
	Next
	rs.close
Next
Conn.Close

Set rs = Nothing
Set Conn = Nothing
Set objWMIService = Nothing

If strComputer = "." Then
	REM wscript.echo("Finished")
Else
	'wscript.echo("Finished " & strComputer)
End If

Private Function ByteSize(strBytes)
	Select Case True
		Case strBytes > 1099511627776
			ByteSize = Round(((((strBytes / 1024) /1024) /1024) /1024),2) & " TB"
		Case strBytes > 1073741824
			ByteSize = Round((((strBytes / 1024) /1024) /1024),2) & " GB"
		Case strBytes > 1048576
			ByteSize = Round(((strBytes / 1024) /1024),2) & " MB"
		Case strBytes > 1024
			ByteSize = Round((strBytes / 1024),2) & " KB"
		Case strBytes > 0
			ByteSize = strBytes & " BYTES"
		Case Else
			ByteSize = ""
	End Select
End Function

Private Function inArray(arr, obj)
	On Error Resume Next
	Dim x, k
	x = -1
	If isArray(arr) Then
		For k = 0 To UBound(arr)
			If arr(k) = obj Then
				x = k
				Exit For
			End If
		Next
	End If
	Err.Clear()
	inArray = x
End Function

Private Function inRecordSet(recordSet, strColumn, strItem)
	On Error Resume Next
	inRecordSet = False
	recordSet.movefirst
	Do While Not recordSet.EOF
		If recordSet(column) =  strItem Then
			inRecordSet = True
			Exit Do
		End If
		recordSet.MoveNext
	Loop	
End Function

Private Function getRecordSet(recordSet, strColumn, strItem, strColumn2)
	On Error Resume Next
	getRecordSet = "Blank"
	recordSet.movefirst
	Do While Not recordSet.EOF
		If recordSet(strColumn) =  strItem Then
			getRecordSet = recordSet(strColumn2)
			Exit Do
		End If
		recordSet.MoveNext
	Loop	
End Function