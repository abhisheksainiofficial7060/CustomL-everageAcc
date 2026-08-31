'#Reference {3F4DACA7-160D-11D2-A8E9-00104B365C9F}#5.5#0#C:\WINNT\System32\vbscript.dll\3#Microsoft VBScript Regular Expressions 5.5
'#Reference {00020813-0000-0000-C000-000000000046}#1.8#0#C:\Program Files\Microsoft Office\Office15\EXCEL.EXE#Microsoft Excel 15.0 Object Library#Excel

Function EndsWith(str As String, ending As String) As Boolean
     Dim endingLen As Integer
     endingLen = Len(ending)
     EndsWith = (Right(Trim(UCase(str)), endingLen) = UCase(ending))
End Function

Function StartsWith(str As String, start As String) As Boolean
     Dim startLen As Integer
     startLen = Len(start)
     StartsWith = (Left(Trim(UCase(str)), startLen) = UCase(start))
End Function

Sub Main
	PSL.OutputWnd.Clear
	Dim prj As PslProject
	Set prj = PSL.ActiveProject
	Dim totalRowNum As Integer
	Dim upRowNum As Integer
	PSL.Output "wait please..."
	
	Dim dtToday as Date
	dtToday = Date()
	
	Dim xlsxPath As String
	xlsxPath = PSL.ActiveProject.Location & "\" & "lpuData.xlsx"
	
	
	Dim eappReport As Excel.Application
	Set eappReport = CreateObject("Excel.Application")
	Dim ConstantIdWb As Excel.Workbook
	Set ConstantIdWb = eappReport.Workbooks.Open(xlsxPath)
	
	Dim lastRow As Integer, LastColumn As Integer
	lastRow    = ConstantIdWb.Worksheets("Sheet1").Cells.SpecialCells(xlCellTypeLastCell).Row
	LastColumn = ConstantIdWb.Worksheets("Sheet1").Cells.SpecialCells(xlCellTypeLastCell).Column
	totalRowNum = lastRow
	' PSL.Output Cstr(lastRow)
	' PSL.Output Cstr(LastColumn)
	' Exit Sub
	


	' read xlsx to list
	Dim srcListXlsx() As Variant
	' Redim srcListXlsx(how many sub arr(start from 0), items in sub arr(start from 0))
	Redim srcListXlsx(1 To lastRow, 1 To LastColumn)
	' Dim srcListXlsx(lastRow,LastColumn) As String
	Dim intRow As Integer, intColumn As Integer
	For intRow = 1 To lastRow
		For intColumn = 1 To LastColumn
			' PSL.Output ConstantIdWb.Worksheets("Sheet1").Cells(intRow,intColumn).Value
			srcListXlsx(intRow, intColumn) = ConstantIdWb.Worksheets("Sheet1").Cells(intRow,intColumn).Value
		Next intColumn
	Next intRow
	
	ConstantIdWb.Close
	eappReport.Quit
	
	' For intRow = 1 To lastRow
		' PSL.Output srcListXlsx(intRow, 1)
		' PSL.Output srcListXlsx(intRow, 2)
		' PSL.Output srcListXlsx(intRow, 3)
		' PSL.Output srcListXlsx(intRow, 4)
		' PSL.Output srcListXlsx(intRow, 5)
		' PSL.Output srcListXlsx(intRow, 6)
		' PSL.Output "--"
	' Next intRow
	
	Dim endVarReg As RegExp
	Set endVarReg = New RegExp
	endVarReg.Pattern = " [{(]adsk_[a-z_]+}(.*)?$"
	endVarReg.IgnoreCase = True
	endVarReg.Global = True
	Dim Matches As MatchCollection
	Dim targetMatches As MatchCollection
		
		' If ymlReg.Test(trnList.TargetFile) = True Then
			' ' If trnStr.Number = 130 Then
				' If newLineReg.Test(trnStr.SourceText) = True Then
					' Set Matches = newLineReg.Execute(trnStr.SourceText)
	
	' update lpu
	' read lpuData from xlsx
	Dim lpuDataExcel As Excel.Application
	Set lpuDataExcel = CreateObject("Excel.Application")
	Dim lpuDataWb As Excel.Workbook
	Set lpuDataWb = lpuDataExcel.Workbooks.Open(xlsxPath)
	
	upRowNum = 0
	Dim i As Integer
	For i = 1 To prj.TransLists.Count
		Dim trnList As PslTransList
		Set trnList = prj.TransLists.Item(i)
		' PSL.Output Cstr(trnList.StringCount(pslIndex))
		Dim j As Long ' Must be Long to avoid overflow!
		For j = 1 To trnList.StringCount(pslIndex)
			Dim trnStr As PslTransString
			Set trnStr = trnList.String(j, pslIndex)
			' If trnList.Language.LangCode = "chs" And  trnStr.Number = 227 Then
			' If trnList.Language.LangCode = "plk" And  trnStr.Number = 3591 Then
			' PSL.Output trnList.Language.LangCode
			' If trnList.Language.LangCode = "jpn" And trnStr.Number = 219 And trnStr.SourceText = "Reset Fusion" Then '---debug
			
			If endVarReg.Test(trnStr.SourceText) = True Then
				Set Matches = endVarReg.Execute(trnStr.SourceText)
				Dim enLpuVarStr As String
				Dim clearSourceStr As String
				enLpuVarStr = Matches(0).Value
				clearSourceStr = Replace(trnStr.SourceText,enLpuVarStr,"")

				' Set targetMatches = endVarReg.Execute(trnStr.Text)
				' Dim targetLpuVarStr As String
				' Dim clearTargetStr As String
				' If targetMatches.Count = 1 Then
					' targetLpuVarStr = targetMatches(0).Value
					' clearTargetStr = Replace(trnStr.Text,targetLpuVarStr,"")
				' Else
					' targetLpuVarStr = "gggggggggggggggggg"
					' PSL.Output "33333"
				' End If
				
				' PSL.Output "enLpuVarStr: " & enLpuVarStr
				' PSL.Output "targetLpuVarStr: " & targetLpuVarStr
				
				

				' PSL.Output "excel start--"
				' check if only one row in xlsx matched
				Dim MatchNum As Integer
				MatchNum = 0
				For intRow = 1 To lastRow
					If trnList.Language.LangCode = srcListXlsx(intRow, 1) And _
					PSL.ActiveProject.Name & ".lpu" = srcListXlsx(intRow, 2) And _
					EndsWith(trnList.SourceList.SourceFile,srcListXlsx(intRow, 3)) = True And _
					StartsWith(srcListXlsx(intRow, 6),clearSourceStr) Then
						MatchNum = MatchNum + 1
					End If
				Next intRow
				If MatchNum = 1 Then
					' PSL.Output "one matched"
				For intRow = 1 To lastRow
					If trnList.Language.LangCode = srcListXlsx(intRow, 1) And _
					PSL.ActiveProject.Name & ".lpu" = srcListXlsx(intRow, 2) And _
					EndsWith(trnList.SourceList.SourceFile,srcListXlsx(intRow, 3)) = True And _
					StartsWith(srcListXlsx(intRow, 6),clearSourceStr) Then
						' PSL.Output Cstr(intRow)
						' PSL.Output srcListXlsx(intRow, 6)
						
						Dim enXlsxVarStr As String
						Set Matches = endVarReg.Execute(srcListXlsx(intRow, 6))
						enXlsxVarStr = Matches(0).Value
						
						Dim finalTargetStr As String
						If InStr(trnStr.Text, enLpuVarStr) = 0 Then
							enLpuVarStr = LTrim(enLpuVarStr)
						End If

						' PSL.Output "current source: " & srcListXlsx(intRow, 6)
						' PSL.Output "trnStr.Text: " & trnStr.Text
						' PSL.Output "enLpuVarStr: " & enLpuVarStr
						' PSL.Output "enXlsxVarStr: " & enXlsxVarStr
						
						If InStr(trnStr.Text, enLpuVarStr) > 0 Then
							finalTargetStr = Replace(trnStr.Text,enLpuVarStr,enXlsxVarStr)
							' PSL.Output "xlsx target: " & srcListXlsx(intRow, 7)
							' PSL.Output "finalTargetStr: " & finalTargetStr
							If srcListXlsx(intRow, 7) <> finalTargetStr Then
								upRowNum = upRowNum + 1
								lpuDataWb.Worksheets("Sheet1").Cells(intRow,7).Value = finalTargetStr
								lpuDataWb.Worksheets("Sheet1").Cells(intRow,8).Value = "updated"
							End If
						End If
					End If
				Next intRow
				End If
			End If
			' End If '---debug
		Next j
	Next i
	' ConstantIdWb.Save
	lpuDataWb.Save
	lpuDataWb.Close
	lpuDataExcel.Quit
	
	' PSL.Output "source over"
	' Exit Sub
	
	PSL.Output "----"
	PSL.Output "Total rows in xlsx: " & totalRowNum
	PSL.Output "Total updated string in xlsx: " & upRowNum
	If totalRowNum <> upRowNum Then
		PSL.Output "Failed: Not all the row of xlsx update"
	Else
		PSL.Output "Sucess: All the row of xlsx update"
	End If

	PSL.Output "--macro execute success--"
	
	
End Sub
