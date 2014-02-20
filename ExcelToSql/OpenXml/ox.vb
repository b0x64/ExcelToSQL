Imports System
Imports System.Data
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Imports System.Collections.Generic
Imports System.Linq
Imports DocumentFormat.OpenXml
Imports WebCreance_V1.OpenXml


Public Class ox
	
	#Region "Microsoft"
	'http://msdn.microsoft.com/en-us/library/office/cc861607.aspx
	
	' Given a document name and text, 
	' inserts a new worksheet and writes the text to cell "A1" of the new worksheet.
	Public Function InsertText(ByVal docName As String, ByVal text As String)
		' Open the document for editing.
		Dim spreadSheet As SpreadsheetDocument = SpreadsheetDocument.Open(docName, True)
		
		Using (spreadSheet)
			' Get the SharedStringTablePart. If it does not exist, create a new one.
			Dim shareStringPart As SharedStringTablePart
			
			If (spreadSheet.WorkbookPart.GetPartsOfType(Of SharedStringTablePart).Count() > 0) Then
				shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType(Of SharedStringTablePart).First()
			Else
				shareStringPart = spreadSheet.WorkbookPart.AddNewPart(Of SharedStringTablePart)()
			End If
			
			' Insert the text into the SharedStringTablePart.
			Dim index As Integer = InsertSharedStringItem(text, shareStringPart)
			
			' Insert a new worksheet.
			Dim worksheetPart As WorksheetPart = InsertWorksheet(spreadSheet.WorkbookPart)
			
			' Insert cell A1 into the new worksheet.
			Dim cell As Cell = InsertCellInWorksheet("A", 1, worksheetPart)
			
			' Set the value of cell A1.
			cell.CellValue = New CellValue(index.ToString)
			cell.DataType = New EnumValue(Of CellValues)(CellValues.SharedString)
			
			' Save the new worksheet.
			worksheetPart.Worksheet.Save()
			
			Return 0
		End Using
	End Function
	
	' Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
	' and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
	Private Shared Function InsertSharedStringItem(ByVal text As String, ByVal shareStringPart As SharedStringTablePart) As Integer
		' If the part does not contain a SharedStringTable, create one.
		If (shareStringPart.SharedStringTable Is Nothing) Then
			shareStringPart.SharedStringTable = New SharedStringTable
		End If
		
		Dim i As Integer = 0
		
		' Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
		For Each item As SharedStringItem In shareStringPart.SharedStringTable.Elements(Of SharedStringItem)()
			If (item.InnerText = text) Then
				Return i
			End If
			i = (i + 1)
		Next
		
		' The text does not exist in the part. Create the SharedStringItem and return its index.
		shareStringPart.SharedStringTable.AppendChild(New SharedStringItem(New DocumentFormat.OpenXml.Spreadsheet.Text(text)))
		shareStringPart.SharedStringTable.Save()
		
		Return i
	End Function
	
	' Given a WorkbookPart, inserts a new worksheet.
	Private Shared Function InsertWorksheet(ByVal workbookPart As WorkbookPart) As WorksheetPart
		' Add a new worksheet part to the workbook.
		Dim newWorksheetPart As WorksheetPart = workbookPart.AddNewPart(Of WorksheetPart)()
		newWorksheetPart.Worksheet = New Worksheet(New SheetData)
		newWorksheetPart.Worksheet.Save()
		Dim sheets As Sheets = workbookPart.Workbook.GetFirstChild(Of Sheets)()
		Dim relationshipId As String = workbookPart.GetIdOfPart(newWorksheetPart)
		
		' Get a unique ID for the new sheet.
		Dim sheetId As UInteger = 1
		If (sheets.Elements(Of Sheet).Count() > 0) Then
			sheetId = sheets.Elements(Of Sheet).Select(Function(s) s.SheetId.Value).Max() + 1
		End If
		
		Dim sheetName As String = ("Sheet" + sheetId.ToString())
		
		' Add the new worksheet and associate it with the workbook.
		Dim sheet As Sheet = New Sheet
		sheet.Id = relationshipId
		sheet.SheetId = sheetId
		sheet.Name = sheetName
		sheets.Append(sheet)
		workbookPart.Workbook.Save()
		
		Return newWorksheetPart
	End Function
	
	' Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
	' If the cell already exists, return it. 
	Private Function InsertCellInWorksheet(ByVal columnName As String, ByVal rowIndex As UInteger, ByVal worksheetPart As WorksheetPart) As Cell
		Dim worksheet As Worksheet = worksheetPart.Worksheet
		Dim sheetData As SheetData = worksheet.GetFirstChild(Of SheetData)()
		Dim cellReference As String = (columnName + rowIndex.ToString())
		
		' If the worksheet does not contain a row with the specified row index, insert one.
		Dim row As Row
		If (sheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value = rowIndex).Count() <> 0) Then
			row = sheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value = rowIndex).First()
		Else
			row = New Row()
			row.RowIndex = rowIndex
			sheetData.Append(row)
		End If
		
		' If there is not a cell with the specified column name, insert one.  
		If (row.Elements(Of Cell).Where(Function(c) c.CellReference.Value = columnName + rowIndex.ToString()).Count() > 0) Then
			Return row.Elements(Of Cell).Where(Function(c) c.CellReference.Value = cellReference).First()
		Else
			' Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
			Dim refCell As Cell = Nothing
			For Each cell As Cell In row.Elements(Of Cell)()
				If (String.Compare(cell.CellReference.Value, cellReference, True) > 0) Then
					refCell = cell
					Exit For
				End If
			Next
			
			Dim newCell As Cell = New Cell
			newCell.CellReference = cellReference
			
			row.InsertBefore(newCell, refCell)
			worksheet.Save()
			
			Return newCell
		End If
	End Function
	
	
	#End Region
	
	#Region "Excel"
	
	Public Shared Function getExcelFileName(Optional prefix As String = "") As String
		Try
			If prefix.Length > 0 Then
				prefix &= "_"
			End If
			
			'Dim TMP_EXCEL = System.Web.HttpContext.Current.Server.MapPath("/WebCreance_V1/TMP_EXCEL")
			
			'fcts.cmd("mkdir """ & TMP_EXCEL & """", True)
			
			'Return TMP_EXCEL & "\" & prefix & Now.Ticks & ".xlsx"
			
			Return   Now.Ticks & ".xlsx"
			
		Catch ex As Exception
			Throw ex
		End Try
	End Function
	
	Public Shared Function CreatExcel(rows As List(Of Row), Optional SHEET_NAME As String = "Sheet1", Optional ByVal PREFIX_FILENAME As String = "") As String
		Try
			
			Dim filename = ox.getExcelFileName(PREFIX_FILENAME)
			
			Using xl = SpreadsheetDocument.Create(filename, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook)
				Dim wbp As WorkbookPart = xl.AddWorkbookPart()
				Dim wsp As WorksheetPart = wbp.AddNewPart(Of WorksheetPart)()
				Dim wb As New Workbook()
				Dim fv As New FileVersion()
				fv.ApplicationName = "WebTreso"
				Dim ws As New Worksheet()
				Dim sd As New SheetData()
				
				Dim wbsp As WorkbookStylesPart = wbp.AddNewPart(Of WorkbookStylesPart)()
				wbsp.Stylesheet = New CustomStylesheet()
				wbsp.Stylesheet.Save()
				
				'columns width ---------------------------------------
				Dim columnsxl As New Columns()
				columnsxl.Append(New ColumnData(8, 8, 25))
				columnsxl.Append(New ColumnData(9, 9, 25))
				columnsxl.Append(New ColumnData(10, 10, 20))
				columnsxl.Append(New ColumnData(11, 11, 20))
				ws.Append(columnsxl)
				'columns width ---------------------------------------
				
				
				For Each newRow In rows
					sd.AppendChild(newRow)
				Next
				
				ws.Append(sd)
				wsp.Worksheet = ws
				wsp.Worksheet.Save()
				Dim sheets As New Sheets()
				Dim sheet As New Sheet()
				sheet.Name = SHEET_NAME
				sheet.SheetId = 1
				sheet.Id = wbp.GetIdOfPart(wsp)
				sheets.Append(sheet)
				wb.Append(fv)
				wb.Append(sheets)
				
				xl.WorkbookPart.Workbook = wb
				xl.WorkbookPart.Workbook.Save()
				xl.Close()
			End Using
			
			
			Return filename
			
		Catch ex As Exception
			Throw ex
		End Try
	End Function
	
	
	Public Shared Function CreatExcel(rows As List(Of Row), cellsGroup As List(Of CellGroup), Optional SHEET_NAME As String = "Sheet1", Optional ByVal PREFIX_FILENAME As String = "") As String
		Try
			
			Dim filename = ox.getExcelFileName(PREFIX_FILENAME)
			
			Using xl = SpreadsheetDocument.Create(filename, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook)
				Dim wbp As WorkbookPart = xl.AddWorkbookPart()
				Dim wsp As WorksheetPart = wbp.AddNewPart(Of WorksheetPart)()
				Dim wb As New Workbook()
				Dim fv As New FileVersion()
				fv.ApplicationName = "WebTreso"
				Dim ws As New Worksheet()
				Dim sd As New SheetData()
				
				Dim wbsp As WorkbookStylesPart = wbp.AddNewPart(Of WorkbookStylesPart)()
				wbsp.Stylesheet = New CustomStylesheet()
				wbsp.Stylesheet.Save()
				
				'columns width ---------------------------------------
				Dim columnsxl As New Columns()
				columnsxl.Append(New ColumnData(8, 8, 25))
				columnsxl.Append(New ColumnData(9, 9, 25))
				columnsxl.Append(New ColumnData(10, 10, 20))
				columnsxl.Append(New ColumnData(11, 11, 20))
				ws.Append(columnsxl)
				'columns width ---------------------------------------
				
				
				For Each newRow In rows
					sd.AppendChild(newRow)
				Next
				
				ws.Append(sd)
				wsp.Worksheet = ws
				wsp.Worksheet.Save()
				Dim sheets As New Sheets()
				Dim sheet As New Sheet()
				sheet.Name = SHEET_NAME
				sheet.SheetId = 1
				sheet.Id = wbp.GetIdOfPart(wsp)
				sheets.Append(sheet)
				wb.Append(fv)
				wb.Append(sheets)
				
				xl.WorkbookPart.Workbook = wb
				xl.WorkbookPart.Workbook.Save()
				
				For Each g In cellsGroup
					FusionnerCells.MergeTwoCells(ws, g.CellName1, g.CellName2)
				Next
				
				xl.Close()
				
			End Using
			
			
			
			
			Return filename
			
		Catch ex As Exception
			Throw ex
		End Try
	End Function
	
	Public Shared Function DataSetToExcel(ds As DataSet) As String
		Try
			
			Dim rows As New List(Of Row)
			Dim newRow As Row
			
			For Each table As DataTable In ds.Tables
				Dim headerRow As New Row()
				Dim columns As New ArrayList
				
				For Each column As DataColumn In table.Columns
					columns.Add(column.ColumnName)
					Dim cell As New StringCell(column.ColumnName)
					cell.StyleIndex = 11
					headerRow.AppendChild(cell)
				Next
				
				rows.Add(headerRow)
				
				For Each dsrow As DataRow In table.Rows
					newRow = New Row()
					
					For Each col As String In columns
						newRow.AppendChild(New StringCell(fcts.getField(dsrow, col).ToString()))
					Next
					
					rows.Add(newRow)
				Next
			Next
			
			Return ox.CreatExcel(rows)
			
		Catch ex As Exception
			Throw ex
		End Try
	End Function
	
	Public Shared Function DataTableToExcel(table As DataTable) As String
		Try
			
			
			Dim rows As New List(Of Row)
			Dim newRow As Row
			
			Dim headerRow As New Row()
			Dim columns As New ArrayList
			
			For Each column As DataColumn In table.Columns
				columns.Add(column.ColumnName)
				Dim cell As New StringCell(column.ColumnName)
				cell.StyleIndex = 11
				headerRow.AppendChild(cell)
			Next
			
			rows.Add(headerRow)
			
			For Each dsrow As DataRow In table.Rows
				newRow = New Row()
				
				For Each col As String In columns
					newRow.AppendChild(New StringCell(fcts.getField(dsrow, col).ToString()))
				Next
				
				rows.Add(newRow)
			Next
			
			Return ox.CreatExcel(rows)
			
		Catch ex As Exception
			Throw ex
		End Try
	End Function
	
	Public Shared Function DataTablesToExcel(tables() As DataTable) As String
		Try
			
			Dim ds As New DataSet
			
			For Each t In tables
				ds.Tables.Add(t)
			Next
			
			Return DataSetToExcel(ds)
			
			
		Catch ex As Exception
			Throw ex
		End Try
	End Function
	
	Public Shared Function ExcelToDataTable(filename As String, t() As Integer) As DataTable
		Try
			
			Dim dt As New DataTable()
			
			Using doc As SpreadsheetDocument = SpreadsheetDocument.Open(filename, False)
				
				Dim workbookPart As WorkbookPart = doc.WorkbookPart
				Dim sheets As IEnumerable(Of Sheet) = doc.WorkbookPart.Workbook.GetFirstChild(Of Sheets)().Elements(Of Sheet)()
				Dim relationshipId As String = sheets.First().Id.Value
				Dim worksheetPart As WorksheetPart = DirectCast(doc.WorkbookPart.GetPartById(relationshipId), WorksheetPart)
				Dim workSheet As Worksheet = worksheetPart.Worksheet
				Dim sheetData As SheetData = workSheet.GetFirstChild(Of SheetData)()
				Dim rows As IEnumerable(Of Row) = sheetData.Descendants(Of Row)()
				
				Dim col_index = 0
				
				
				
				For Each cell As Cell In rows.ElementAt(0)
					
					For i = 0 To t.Length - 1
						If t(i) = col_index Then
							dt.Columns.Add(col_index & "-" & GetCellValue(doc, cell))
						End If
					Next
					col_index += 1
				Next
				
				'For i = 0 To t.Length - 1
				'    dt.Columns.Add(col_index & "-" & GetCellValue(doc, Cell))
				'    col_index += 1
				'Next
				
				Dim col_index_2 = 0
				
				For Each row As Row In rows
					'this will also include your header row ...
					Dim tempRow As DataRow = dt.NewRow()
					
					For i As Integer = 0 To t.Length - 1
						
						tempRow(i) = GetCellValue(doc, row.Descendants(Of Cell)().ElementAt(t(i)))
						
					Next
					
					dt.Rows.Add(tempRow)
				Next
			End Using
			
			dt.Rows.RemoveAt(0)
			
			Return dt
			
		Catch ex As Exception
			Throw ex
		End Try
	End Function
	
	Public Shared Function ExcelToDataTable(filename As String, Optional start_row As Integer = 1) As DataTable
		Try
			
			Dim dt As New DataTable()
			
			Using doc As SpreadsheetDocument = SpreadsheetDocument.Open(filename, False)
				
				Dim workbookPart As WorkbookPart = doc.WorkbookPart
				Dim sheets As IEnumerable(Of Sheet) = doc.WorkbookPart.Workbook.GetFirstChild(Of Sheets)().Elements(Of Sheet)()
				Dim relationshipId As String = sheets.First().Id.Value
				Dim worksheetPart As WorksheetPart = DirectCast(doc.WorkbookPart.GetPartById(relationshipId), WorksheetPart)
				Dim workSheet As Worksheet = worksheetPart.Worksheet
				Dim sheetData As SheetData = workSheet.GetFirstChild(Of SheetData)()
				Dim rows As IEnumerable(Of Row) = sheetData.Descendants(Of Row)()
				
				Dim col_index = 0
				
				For i = 0 To rows.Count - 1
					If rows(i).RowIndex.Value = start_row Then
						start_row = i
						Exit For
					End If
				Next
				
				Dim Alpha As New ArrayList
				
				For Each cell As Cell In rows(start_row).Descendants(Of Cell)()
					dt.Columns.Add(col_index & " - " & GetCellValue(doc, cell))
					'dt.Columns.Add(cell.CellReference.ToString().Substring(0, 1))
					Alpha.Add(cell.CellReference.ToString().Substring(0, 1))
					col_index += 1
				Next
				
				For i = start_row + 1 To rows.Count - 1
					Dim tempRow As DataRow = dt.NewRow()
					Dim row = rows(i)
					
					For j = 0 To dt.Columns.Count - 1
						Dim ref As String = Alpha(j).ToString() & row.RowIndex.Value.ToString()
						Dim col As DataColumn = dt.Columns(j)
						
						Dim cc = (From c In row.Descendants(Of Cell)() Where c.CellReference = ref).FirstOrDefault()
						
						If IsNothing(cc) Then
							tempRow(col.ColumnName) = ""
						Else
							tempRow(col.ColumnName) = GetCellValue(doc, cc)
						End If
					Next
					
					dt.Rows.Add(tempRow)
				Next
			End Using
			
			'dt.Rows.RemoveAt(0)
			
			Return dt
			
		Catch ex As Exception
			Throw ex
		End Try
	End Function
	
	Public Shared Function ExcelToDataTable(filename As String, start_row As Integer, end_row As Integer) As DataTable
		Try
			
			Dim dt As New DataTable()
			
			Using doc As SpreadsheetDocument = SpreadsheetDocument.Open(filename, False)
				
				Dim workbookPart As WorkbookPart = doc.WorkbookPart
				Dim sheets As IEnumerable(Of Sheet) = doc.WorkbookPart.Workbook.GetFirstChild(Of Sheets)().Elements(Of Sheet)()
				Dim relationshipId As String = sheets.First().Id.Value
				Dim worksheetPart As WorksheetPart = DirectCast(doc.WorkbookPart.GetPartById(relationshipId), WorksheetPart)
				Dim workSheet As Worksheet = worksheetPart.Worksheet
				Dim sheetData As SheetData = workSheet.GetFirstChild(Of SheetData)()
				Dim rows As IEnumerable(Of Row) = sheetData.Descendants(Of Row)()
				
				Dim col_index = 0
				
				For i = 0 To rows.Count - 1
					If rows(i).RowIndex.Value = start_row Then
						start_row = i
						Exit For
					End If
				Next
				
				Dim Alpha As New ArrayList
				
				For Each cell As Cell In rows(start_row).Descendants(Of Cell)()
					dt.Columns.Add(col_index & " - " & GetCellValue(doc, cell))
					Alpha.Add(cell.CellReference.ToString().Substring(0, 1))
					col_index += 1
				Next
				
				For i = start_row + 1 To rows.Count - 1
					Dim tempRow As DataRow = dt.NewRow()
					Dim row = rows(i)
					
					If row.RowIndex.Value > end_row Then
						Exit For
					End If
					
					For j = 0 To dt.Columns.Count - 1
						Dim ref As String = Alpha(j).ToString() & row.RowIndex.Value.ToString()
						Dim col As DataColumn = dt.Columns(j)
						
						Dim cc = (From c In row.Descendants(Of Cell)() Where c.CellReference = ref).FirstOrDefault()
						
						If IsNothing(cc) Then
							tempRow(col.ColumnName) = ""
						Else
							tempRow(col.ColumnName) = GetCellValue(doc, cc)
						End If
					Next
					
					dt.Rows.Add(tempRow)
				Next
			End Using
			
			Return dt
			
		Catch ex As Exception
			Throw ex
		End Try
	End Function
	
	Public Shared Function GetCellValue(document As SpreadsheetDocument, cell As Cell) As String
		Try
			
			
			
			If IsNothing(cell.CellValue) Then
				Return ""
			End If
			
			Dim value As String = cell.CellValue.InnerXml
			
			If Not IsNothing(cell.DataType) Then
				If cell.DataType.Value = CellValues.SharedString Then
					Dim stringTablePart As SharedStringTablePart = document.WorkbookPart.SharedStringTablePart
					Return stringTablePart.SharedStringTable.ChildElements(Int32.Parse(value)).InnerText
				End If
			End If
			
			Return value
			
		Catch ex As Exception
			Return ""
		End Try
	End Function
	
	Public Shared Function DateExcel(s) As Date
		Try
			
			Return DateTime.FromOADate(Double.Parse(s)).ToShortDateString()
			
		Catch ex As Exception
			Return Now.Date
		End Try
	End Function
	
	Private Shared Function InsertSharedStringItem2(ByVal text As String, ByVal shareStringPart As SharedStringTablePart) As Integer
		' If the part does not contain a SharedStringTable, create one.
		If (shareStringPart.SharedStringTable Is Nothing) Then
			shareStringPart.SharedStringTable = New SharedStringTable
		End If
		
		Dim i As Integer = 0
		
		' Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
		For Each item As SharedStringItem In shareStringPart.SharedStringTable.Elements(Of SharedStringItem)()
			If (item.InnerText = text) Then
				Return i
			End If
			i = (i + 1)
		Next
		
		' The text does not exist in the part. Create the SharedStringItem and return its index.
		shareStringPart.SharedStringTable.AppendChild(New SharedStringItem(New DocumentFormat.OpenXml.Spreadsheet.Text(text)))
		'shareStringPart.SharedStringTable.Save()
		
		Return i
	End Function
	
	Public Shared Function InsertText(spreadSheet As SpreadsheetDocument, cell As Cell, text As String)
		
		Dim shareStringPart As SharedStringTablePart
		
		If (spreadSheet.WorkbookPart.GetPartsOfType(Of SharedStringTablePart).Count() > 0) Then
			shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType(Of SharedStringTablePart).First()
		Else
			shareStringPart = spreadSheet.WorkbookPart.AddNewPart(Of SharedStringTablePart)()
		End If
		
		' Insert the text into the SharedStringTablePart.
		Dim index As Integer = InsertSharedStringItem2(text, shareStringPart)
		
		' Insert a new worksheet.
		'Dim worksheetPart As WorksheetPart = InsertWorksheet(spreadSheet.WorkbookPart)
		
		
		' Set the value of cell A1.
		cell.CellValue = New CellValue(index.ToString)
		cell.DataType = New EnumValue(Of CellValues)(CellValues.SharedString)
		
		' Save the new worksheet.
		'worksheetPart.Worksheet.Save()
		
		Return 0
		
	End Function
	#End Region
	
End Class

Public Class CellGroup
	Public Property CellName1 As String
		Public Property CellName2 As String
End Class