Imports System
Imports System.Data
Imports System.IO
Imports System.Threading
Imports Excel

Public Partial Class MainForm
	Public Sub New()
		' The Me.InitializeComponent call is required for Windows Forms designer support.
		Me.InitializeComponent()
		
		'
		' TODO : Add constructor code after InitializeComponents
		'
	End Sub
	
	
	Private th As Thread
	
	Sub CreateSQL(dt As DataTable)
		fcts.stopp = False
		tbResult.Text = fcts.DataTableToSql(dt)
	End Sub
	
	Sub WaitPoint()
		lbStatus.Text = "*"
		While th.IsAlive
			lbStatus.Text &= "*"
			Thread.Sleep(1000)
		End While
		lbStatus.Text = "OK"
	End Sub
	
	Sub BtnConvertClick(sender As Object, e As EventArgs)
		Try
			
			If openf1.ShowDialog() = DialogResult.OK Then
				Dim filename = openf1.FileName
				lbStatus.Text = filename
				Dim fi = New FileInfo(filename)
				Dim dt As New DataTable
				
				If fi.Extension.ToLower = ".xls" Then
					dt = xlreader.ExcelToDataTable(filename)
				ElseIf fi.Extension.ToLower = ".xlsx" Then
					dt = ox.ExcelToDataTable(filename)
				End If
							
				
				With dgData
					.AutoGenerateColumns = True
					.DataSource = dt
				End With
				
				lbStatus.Text = "ok"
				
				th = New Thread(Sub() CreateSQL(dt))
				th.Start()
				
				Dim th2 As New Thread(Sub() WaitPoint())
				th2.Start()
				
				'CreateSQL(dt)
				
			End If
			
		Catch ex As Exception			
			lbStatus.Text = ex.Message
		End Try
	End Sub
	
	Sub BtnStopClick(sender As Object, e As EventArgs)
		fcts.stopp = True
	End Sub
	
	Sub MainFormFormClosing(sender As Object, e As FormClosingEventArgs)
		fcts.stopp = True
	End Sub
	
	Sub BtnCopyClick(sender As Object, e As EventArgs)
		tbResult.SelectAll()
		tbResult.Copy()
	End Sub
End Class
