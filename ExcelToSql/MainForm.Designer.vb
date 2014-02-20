'
' Created by SharpDevelop.
' User: khaled
' Date: 20/02/2014
' Time: 10:10
' 
' To change this template use Tools | Options | Coding | Edit Standard Headers.
'
Partial Class MainForm
	Inherits System.Windows.Forms.Form
	
	''' <summary>
	''' Designer variable used to keep track of non-visual components.
	''' </summary>
	Private components As System.ComponentModel.IContainer
	
	''' <summary>
	''' Disposes resources used by the form.
	''' </summary>
	''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
	Protected Overrides Sub Dispose(ByVal disposing As Boolean)
		If disposing Then
			If components IsNot Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(disposing)
	End Sub
	
	''' <summary>
	''' This method is required for Windows Forms designer support.
	''' Do not change the method contents inside the source code editor. The Forms designer might
	''' not be able to load this method if it was changed manually.
	''' </summary>
	Private Sub InitializeComponent()
		Me.openf1 = New System.Windows.Forms.OpenFileDialog()
		Me.btnConvert = New System.Windows.Forms.Button()
		Me.lbStatus = New System.Windows.Forms.Label()
		Me.dgData = New System.Windows.Forms.DataGridView()
		Me.tbResult = New System.Windows.Forms.RichTextBox()
		Me.btnStop = New System.Windows.Forms.Button()
		Me.btnCopy = New System.Windows.Forms.Button()
		CType(Me.dgData,System.ComponentModel.ISupportInitialize).BeginInit
		Me.SuspendLayout
		'
		'openf1
		'
		Me.openf1.FileName = "openf1"
		'
		'btnConvert
		'
		Me.btnConvert.Location = New System.Drawing.Point(230, 12)
		Me.btnConvert.Name = "btnConvert"
		Me.btnConvert.Size = New System.Drawing.Size(242, 50)
		Me.btnConvert.TabIndex = 0
		Me.btnConvert.Text = "Convert xls/xlsx ==> Sql"
		Me.btnConvert.UseVisualStyleBackColor = true
		AddHandler Me.btnConvert.Click, AddressOf Me.BtnConvertClick
		'
		'lbStatus
		'
		Me.lbStatus.Location = New System.Drawing.Point(12, 67)
		Me.lbStatus.Name = "lbStatus"
		Me.lbStatus.Size = New System.Drawing.Size(683, 20)
		Me.lbStatus.TabIndex = 1
		Me.lbStatus.Text = "label1"
		'
		'dgData
		'
		Me.dgData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.dgData.Location = New System.Drawing.Point(12, 90)
		Me.dgData.Name = "dgData"
		Me.dgData.Size = New System.Drawing.Size(683, 213)
		Me.dgData.TabIndex = 2
		'
		'tbResult
		'
		Me.tbResult.Location = New System.Drawing.Point(12, 309)
		Me.tbResult.Name = "tbResult"
		Me.tbResult.Size = New System.Drawing.Size(682, 270)
		Me.tbResult.TabIndex = 3
		Me.tbResult.Text = ""
		'
		'btnStop
		'
		Me.btnStop.Location = New System.Drawing.Point(478, 14)
		Me.btnStop.Name = "btnStop"
		Me.btnStop.Size = New System.Drawing.Size(130, 50)
		Me.btnStop.TabIndex = 4
		Me.btnStop.Text = "Stop"
		Me.btnStop.UseVisualStyleBackColor = true
		AddHandler Me.btnStop.Click, AddressOf Me.BtnStopClick
		'
		'btnCopy
		'
		Me.btnCopy.Location = New System.Drawing.Point(266, 589)
		Me.btnCopy.Name = "btnCopy"
		Me.btnCopy.Size = New System.Drawing.Size(130, 50)
		Me.btnCopy.TabIndex = 5
		Me.btnCopy.Text = "Copier"
		Me.btnCopy.UseVisualStyleBackColor = true
		AddHandler Me.btnCopy.Click, AddressOf Me.BtnCopyClick
		'
		'MainForm
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 13!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(708, 651)
		Me.Controls.Add(Me.btnCopy)
		Me.Controls.Add(Me.btnStop)
		Me.Controls.Add(Me.tbResult)
		Me.Controls.Add(Me.dgData)
		Me.Controls.Add(Me.lbStatus)
		Me.Controls.Add(Me.btnConvert)
		Me.Name = "MainForm"
		Me.Text = "ExcelToSql"
		AddHandler FormClosing, AddressOf Me.MainFormFormClosing
		CType(Me.dgData,System.ComponentModel.ISupportInitialize).EndInit
		Me.ResumeLayout(false)
	End Sub
	Private btnCopy As System.Windows.Forms.Button
	Private btnStop As System.Windows.Forms.Button
	Private tbResult As System.Windows.Forms.RichTextBox
	Private dgData As System.Windows.Forms.DataGridView
	Private lbStatus As System.Windows.Forms.Label
	Private btnConvert As System.Windows.Forms.Button
	Private openf1 As System.Windows.Forms.OpenFileDialog
End Class
