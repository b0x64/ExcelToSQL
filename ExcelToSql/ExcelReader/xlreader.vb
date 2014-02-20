
Imports System
Imports System.IO
Imports System.Data
Imports Excel

Public Class xlreader
	
	
	Public Shared Function ExcelToDataTable(filename As String)As DataTable
		Try
			
			Using stream As New IO.FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.None)
				
				Dim excelReader As IExcelDataReader = ExcelReaderFactory.CreateBinaryReader(stream)
				excelReader.IsFirstRowAsColumnNames = True
				Dim dt = excelReader.AsDataSet().Tables(0)
				
				Return dt
				
			End Using
			
		Catch ex As Exception
			Throw ex
		End Try
	End Function
	
	
End Class
