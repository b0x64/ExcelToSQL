Imports System
Imports System.Data


Public Class fcts
	
	Public Shared stopp As Boolean
	
	
	Public Shared Sub msg(m As String)
		MessageBox.Show(m)
	End Sub
	
	Public Shared Sub msg(ex As Exception)
		If IsNothing(ex.InnerException) Then
			MessageBox.Show(ex.Message)			
		Else
			MessageBox.Show(ex.InnerException.Message)			
		End If
	End Sub
	
	Public Shared Function strDouble(ByVal str As String) As String
		Dim sp = ""
		If str.IndexOf(",") > -1 And str.IndexOf(".") > -1 Then
			str = str.Replace(",", "")
		End If
		
		Dim separateur = System.Globalization.NumberFormatInfo.CurrentInfo.CurrencyDecimalSeparator
		Return str.Trim.Replace(" ", "").Replace(".", separateur).Replace(",", separateur)
	End Function
	
	#Region "Data"
	Public Shared Function getField(dr As DataRow, fname As String, Optional DBnullValue As Object = "")
		Try
			
			If IsDBNull(dr(fname)) Then
				Return DBnullValue
			End If
			
			Return dr(fname)
			
		Catch ex As Exception
			Return DBnullValue
		End Try
	End Function
	
	Public Shared Function getFieldDate(dr As DataRow, fname As String, Optional DBnullValue As Date = Nothing) As Date
		Try
			
			If DBnullValue = Nothing Then
				DBnullValue = New Date(1900, 1, 1)
			End If
			
			If IsDBNull(dr(fname)) Then
				Return DBnullValue
			End If
			
			If IsDate(dr(fname)) Then
				Return dr(fname)
			End If
			
			Return ox.DateExcel(dr(fname))
			
		Catch ex As Exception
			Return New DateTime(1900, 1, 1)
		End Try
	End Function
	
	Public Shared Function getFieldDouble(dr As DataRow, fname As String) As Double
		Try
			
			If IsDBNull(dr(fname)) Then
				Return -1
			End If
			
			Return CDbl(fcts.strDouble(dr(fname).ToString()))
			
		Catch ex As Exception
			Return -1
		End Try
	End Function
	
	Public Shared Function getFieldBool(dr As DataRow, fname As String) As Boolean
		Try
			
			If IsDBNull(dr(fname)) Then
				Return False
			End If
			
			Return CBool(dr(fname).ToString())
			
		Catch ex As Exception
			Return False
		End Try
	End Function
	
	Public Shared Function getFieldGuid(dr As DataRow, fname As String) As Guid
		Try
			
			If IsDBNull(dr(fname)) Then
				Return Guid.Empty
			End If
			
			Return New Guid(dr(fname).ToString())
			
		Catch ex As Exception
			Return Guid.Empty
		End Try
	End Function
	
	Public Shared Function getFieldAt(dr As DataRow, idx As Integer, Optional DBnullValue As Object = "")
		Try
			
			If IsDBNull(dr(idx)) Then
				Return DBnullValue
			End If
			
			Return dr(idx)
			
		Catch ex As Exception
			Return DBnullValue
		End Try
	End Function
	
	Public Shared Function getFieldDoubleAt(dr As DataRow, idx As Integer, Optional DBnullValue As Double = 0) As Double
		Try
			
			If IsDBNull(dr(idx)) Then
				Return DBnullValue
			End If
			
			Return strDouble(dr(idx).ToString())
			
		Catch ex As Exception
			Return DBnullValue
		End Try
	End Function
	
	Public Shared Function getFieldDateAt(dr As DataRow, idx As Integer, Optional DBnullValue As Date = Nothing) As Date
		Try
			
			If DBnullValue = Nothing Then
				DBnullValue = New Date(1900, 1, 1)
			End If
			
			If IsDBNull(dr(idx)) Then
				Return DBnullValue
			End If
			
			If IsDate(dr(idx)) Then
				Return dr(idx)
			End If
			
			Return ox.DateExcel(dr(idx))
			
			
			
		Catch ex As Exception
			Return DBnullValue
		End Try
	End Function
	#End Region
	
	Public Shared Function DataTableToSql(dt As DataTable) As String
		Try
			
			'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			Dim s As String = ""
			Dim data As String = ""
			Dim sql As String = ""
			
			Dim str_create = "--DROP TABLE TABLE_NAME$$$" & Environment.NewLine &  "CREATE TABLE TABLE_NAME$$$(" & Environment.NewLine
			For j As Integer = 0 To dt.Columns.Count - 1
				If j = dt.Columns.Count - 1 Then
					str_create &= "[" & dt.Columns(j).ColumnName.Replace("[", "_").Replace("]", "_") & "] NVARCHAR(MAX)" & Environment.NewLine
				Else
					str_create &= "[" & dt.Columns(j).ColumnName.Replace("[", "_").Replace("]", "_") & "] NVARCHAR(MAX)," & Environment.NewLine
				End If
			Next
			
			str_create &= ")" & Environment.NewLine & Environment.NewLine & Environment.NewLine
			
			For Each dr As DataRow In dt.Rows
				If stopp Then
					Exit For
				End If
				s = ""
				For j As Integer = 0 To dt.Columns.Count - 1
					data = fcts.getFieldAt(dr, j, "NULL").toString().Replace("'", "''")
					
					If (j = 0) Then
						s &= "INSERT INTO TABLE_NAME$$$ VALUES('" & data & "',"
					ElseIf (j = dt.Columns.Count - 1) Then
						s &= "'" & data & "')"
					Else
						s &= "'" & data & "',"
					End If
				Next
				sql &= s.Replace("'NULL'", "NULL") & Environment.NewLine
			Next
			'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			
			Return str_create & sql
			
		Catch ex As Exception
			Return ""
		End Try
	End Function
	
End Class
