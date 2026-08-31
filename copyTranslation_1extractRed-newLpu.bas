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
	Dim rowNum As Integer
	
	PSL.Output "wait please..."
	
	
	' read lpuData from xlsx
	Dim lpuDataXlsxPath As String
	lpuDataXlsxPath = PSL.ActiveProject.Location & "\" & "lpuData.xlsx"
	Dim lpuDataExcel As Excel.Application
	Set lpuDataExcel = CreateObject("Excel.Application")
	Dim lpuDataWb As Excel.Workbook
	Set lpuDataWb = lpuDataExcel.Workbooks.Open(lpuDataXlsxPath)
	Dim lpuDataLastRow As Integer, lpuDataLastColumn As Integer
	lpuDataLastRow    = lpuDataWb.Worksheets("Sheet1").Cells.SpecialCells(xlCellTypeLastCell).Row
	lpuDataLastColumn = lpuDataWb.Worksheets("Sheet1").Cells.SpecialCells(xlCellTypeLastCell).Column
	' PSL.Output Cstr(lpuDataLastRow)
	' PSL.Output Cstr(lpuDataLastColumn)
	' ' Exit Sub
	
	If lpuDataLastRow = 1 And lpuDataLastColumn = 1 Then
		rowNum = 1
		lpuDataWb.Worksheets("Sheet1").Cells(1,1).Value = "Language"
		lpuDataWb.Worksheets("Sheet1").Cells(1,2).Value = "Lpu"
		lpuDataWb.Worksheets("Sheet1").Cells(1,3).Value = "File"
		lpuDataWb.Worksheets("Sheet1").Cells(1,4).Value = "Number"
		lpuDataWb.Worksheets("Sheet1").Cells(1,5).Value = "ID"
		lpuDataWb.Worksheets("Sheet1").Cells(1,6).Value = "Source"
		lpuDataWb.Worksheets("Sheet1").Cells(1,7).Value = "Target"
	End If
	If lpuDataLastRow <> 1 And lpuDataLastColumn <> 1 Then
		rowNum = lpuDataLastRow
	End If
	
	
	
	' PSL.Output "-here1-"
	
	
	' base on conf, extract strings
	' Dim endVarReg As RegExp
	' Set endVarReg = New RegExp
	' endVarReg.Pattern = " [{(]adsk_[a-z_]+}([a-z{} x_-]+)?$"
	' endVarReg.IgnoreCase = True
	' endVarReg.Global = True
	
	Dim i As Integer
	For i = 1 To prj.TransLists.Count
		Dim trnList As PslTransList
		Set trnList = prj.TransLists.Item(i)
		' PSL.Output Cstr(trnList.StringCount(pslIndex))
		' If trnList.Language.LangCode = "chs" Then
		Dim j As Long ' Must be Long to avoid overflow!
		For j = 1 To trnList.StringCount(pslIndex)
			Dim trnStr As PslTransString
			Set trnStr = trnList.String(j, pslIndex)
			Dim trnColor As String
			If trnStr.State(pslStateReadOnly) = False And trnStr.State(pslStateTranslated) = True And trnStr.State(pslStateReview) = True And trnStr.State(pslStateAutoTranslated)  = True Then
				trnColor = "green"
			End If
			If trnStr.State(pslStateReadOnly) = False And trnStr.State(pslStateTranslated) = True And trnStr.State(pslStateReview) = True And trnStr.State(pslStateAutoTranslated)  = False Then
				trnColor = "blue"
			End If
			If trnStr.State(pslStateReadOnly) = False And trnStr.State(pslStateTranslated) = False And trnStr.State(pslStateReview) = False And trnStr.State(pslStateAutoTranslated)  = False Then
				trnColor = "red"
			End If
			If trnStr.State(pslStateReadOnly) = False And trnStr.State(pslStateTranslated) = True And trnStr.State(pslStateReview) = False And trnStr.State(pslStateAutoTranslated)  = False Then
				trnColor = "black"
			End If
			If trnStr.Resource.State(pslStateReadOnly) = True Or trnStr.State(pslStateReadOnly) = True Then
				trnColor = "grey"
			End If
			
			' If trnColor = "red" And trnStr.SourceText <> trnStr.Text Then
			If trnColor = "red" Then
			' If trnColor = "red" And endVarReg.Test(trnStr.SourceText) = True Then
				rowNum = rowNum + 1
				' PSL.Output trnStr.SourceText
				' PSL.Output trnStr.Text
				' PSL.Output "--"
				lpuDataWb.Worksheets("Sheet1").Cells(rowNum,1).Value = trnList.Language.LangCode
				lpuDataWb.Worksheets("Sheet1").Cells(rowNum,2).Value = PSL.ActiveProject.Name & ".lpu"
				lpuDataWb.Worksheets("Sheet1").Cells(rowNum,3).Value = Replace(trnList.SourceList.SourceFile,PSL.ActiveProject.Location,"")
				lpuDataWb.Worksheets("Sheet1").Cells(rowNum,4).Value = trnStr.Number
				lpuDataWb.Worksheets("Sheet1").Cells(rowNum,5).Value = trnStr.ID
				lpuDataWb.Worksheets("Sheet1").Cells(rowNum,6).Value = trnStr.SourceText
				lpuDataWb.Worksheets("Sheet1").Cells(rowNum,7).Value = trnStr.Text
			End If

		Next j
		' End If
	Next i
	
	lpuDataWb.Save
	lpuDataWb.Close
	lpuDataExcel.Quit
	
	PSL.Output "Total rows extract in xlsx: " & rowNum

	PSL.Output "--macro execute success--"
	
	
End Sub