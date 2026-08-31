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
		' PSL.Output srcListXlsx(intRow, 7)
		' PSL.Output "--"
	' Next intRow
	
	' update lpu
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
			' PSL.Output trnList.Language.LangCode
			' If trnList.Language.LangCode = "jpn" And trnStr.Number = 219 And trnStr.SourceText = "Reset Fusion" Then '---debug
			
			' loop xlsx list data
			For intRow = 1 To lastRow
				' If srcListXlsx(intRow, 1) = "jpn" And srcListXlsx(intRow, 6) = "Reset Fusion" Then '---debug
				' PSL.Output Cstr(intRow)
				' PSL.Output Cstr(trnList.Language.LangCode = srcListXlsx(intRow, 1))
				' PSL.Output Cstr(trnList.Language.LangCode)
				' PSL.Output Cstr(srcListXlsx(intRow, 1))
				' PSL.Output "--"
				
				' PSL.Output Cstr(PSL.ActiveProject.Name & ".lpu" = srcListXlsx(intRow, 2))
				' PSL.Output Cstr(PSL.ActiveProject.Name & ".lpu")
				' PSL.Output Cstr(srcListXlsx(intRow, 2))
				' PSL.Output "--"
				
				' PSL.Output Cstr(EndsWith(trnList.SourceList.SourceFile,srcListXlsx(intRow, 3)) = True)
				' PSL.Output Cstr(EndsWith(trnList.SourceList.SourceFile,srcListXlsx(intRow, 3)))
				' PSL.Output Cstr(trnList.SourceList.SourceFile)
				' PSL.Output Cstr(srcListXlsx(intRow, 3))
				' PSL.Output "--"
				
				' PSL.Output Cstr(trnStr.Number = srcListXlsx(intRow, 4))
				' PSL.Output Cstr(trnStr.Number)
				' PSL.Output Cstr(srcListXlsx(intRow, 4))
				
				' PSL.Output Cstr(trnStr.ID =  srcListXlsx(intRow, 5))
				' PSL.Output Cstr(trnStr.ID)
				' PSL.Output Cstr(srcListXlsx(intRow, 5))
				
				' PSL.Output Cstr(trnStr.SourceText = srcListXlsx(intRow, 6))
				' PSL.Output Cstr(trnStr.SourceText)
				' PSL.Output Cstr(srcListXlsx(intRow, 6))
				' PSL.Output "--"
				
				' PSL.Output Cstr(trnStr.Text <> srcListXlsx(intRow, 7))
				' PSL.Output Cstr(trnStr.Text)
				' PSL.Output Cstr(srcListXlsx(intRow, 7))
				' PSL.Output "--"
				
				' PSL.Output "--"
				If trnList.Language.LangCode = srcListXlsx(intRow, 1) And _
				PSL.ActiveProject.Name & ".lpu" = srcListXlsx(intRow, 2) And _
				EndsWith(trnList.SourceList.SourceFile,srcListXlsx(intRow, 3)) = True And _
				trnStr.Number = srcListXlsx(intRow, 4) And _
				trnStr.ID = srcListXlsx(intRow, 5) And _
				trnStr.SourceText = srcListXlsx(intRow, 6) Then
				' If trnList.Language.LangCode = srcListXlsx(intRow, 1) And _
				' PSL.ActiveProject.Name & ".lpu" = srcListXlsx(intRow, 2) And _
				' EndsWith(trnList.SourceList.SourceFile,srcListXlsx(intRow, 3)) = True And _
				' trnStr.SourceText = srcListXlsx(intRow, 6) Then
					If srcListXlsx(intRow, 8) = "updated" Then
						If trnStr.Text <> srcListXlsx(intRow, 7) Then
							trnStr.Text = srcListXlsx(intRow, 7)
							' trnStr.TransComment = "customer feedback"
							trnStr.TransComment = "update symbol by Macro (" & Cstr(dtToday) & ")"
						Else
							trnStr.TransComment = "validate symbol by Macro (" & Cstr(dtToday) & ")"
						End If
						trnStr.State(pslStateReview) = False
						trnstr.State(pslStateTranslated) = True
						upRowNum = upRowNum + 1
						PSL.Output "Updating " & upRowNum & "/" & (lastRow)
					End If
				End If
				' End If '---debug
			Next intRow
			' End If '---debug
		Next j
		trnlist.Save
	Next i
	' ConstantIdWb.Save
	
	' PSL.Output "source over"
	' Exit Sub
	
	PSL.Output "Total updated string: " & upRowNum

	PSL.Output "--macro execute success--"
	
	
End Sub