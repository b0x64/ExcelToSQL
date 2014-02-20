Imports System
Imports System.Data
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text

Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet




Namespace OpenXml


#Region "Cell"
    Public Class DateCell
        Inherits Cell
        Public Sub New(header As String, dateTime As DateTime, index As Integer)
            MyBase.New()

            Me.DataType = CellValues.Date
            Me.CellReference = header & index
            Me.StyleIndex = 1
            Me.CellValue = New CellValue() With {.Text = dateTime.ToOADate().ToString()}

        End Sub

        Public Sub New(dateTime As DateTime)
            MyBase.New()

            Me.DataType = CellValues.Date
            Me.StyleIndex = 1
            Me.CellValue = New CellValue() With {.Text = dateTime.ToOADate().ToString()}

        End Sub

    End Class

    Public Class NumberCell
        Inherits Cell
        Public Sub New(header As String, text As String, index As Integer)
            MyBase.New()

            Me.DataType = CellValues.Number
            Me.CellReference = header & index
            Me.CellValue = New CellValue(text.Replace(",", "."))

        End Sub

        Public Sub New(text As String)
            MyBase.New()

            Me.DataType = CellValues.Number
            Me.CellValue = New CellValue(text.Replace(",", "."))
        End Sub

    End Class

    Public Class FormatedNumberCell
        Inherits NumberCell
        Public Sub New(header As String, text As String, index As Integer)
            MyBase.New(header, text, index)
            Me.StyleIndex = 2
        End Sub

        Public Sub New(text As String)
            MyBase.New(text)
            Me.StyleIndex = 2
        End Sub

    End Class

    Public Class TextCell
        Inherits Cell

        Public Sub New(header As String, text As String, index As Integer, styleidx As Integer)
            MyBase.New()

            Me.DataType = CellValues.InlineString
            Me.CellReference = header & index
            Me.StyleIndex = styleidx
            Me.InlineString = New InlineString With {.Text = New Text(text)}

        End Sub
        Public Sub New(header As String, text As String, index As Integer)
            MyBase.New()

            Me.DataType = CellValues.InlineString
            Me.CellReference = header & index

            Me.InlineString = New InlineString With {.Text = New Text(text)}

        End Sub

        Public Sub New(text As String)
            MyBase.New()

            Me.DataType = CellValues.InlineString

            Me.InlineString = New InlineString With {.Text = New Text(text)}

        End Sub

    End Class

    Public Class StringCell
        Inherits Cell

        Public Sub New(header As String, str As String, index As Integer, styleidx As Integer)
            MyBase.New()

            Me.DataType = CellValues.String
            Me.CellReference = header & index
            Me.StyleIndex = styleidx
            Me.CellValue = New CellValue(str)

        End Sub
        Public Sub New(header As String, str As String, index As Integer)
            MyBase.New()

            Me.DataType = CellValues.String
            Me.CellReference = header & index
            Me.CellValue = New CellValue(str)

        End Sub

        Public Sub New(str As String)
            MyBase.New()

            Me.DataType = CellValues.String
            Me.CellValue = New CellValue(str)

        End Sub

    End Class
#End Region

    Public Class CustomStylesheet
        Inherits Stylesheet
        Public Sub New()
            MyBase.New()


            'Fonts ------------------------------------------------------------------------------------------------------------
            Dim fts As New Fonts()
            Dim ft As New DocumentFormat.OpenXml.Spreadsheet.Font()
            Dim ftn As New FontName()
            ftn.Val = StringValue.FromString("Calibri")
            Dim ftsz As New FontSize()
            ftsz.Val = DoubleValue.FromDouble(11)
            ft.FontName = ftn
            ft.FontSize = ftsz
            fts.Append(ft)

            ft = New DocumentFormat.OpenXml.Spreadsheet.Font()
            ftn = New FontName()
            ftn.Val = StringValue.FromString("Palatino Linotype")
            ftsz = New FontSize()
            ftsz.Val = DoubleValue.FromDouble(18)
            ft.FontName = ftn
            ft.FontSize = ftsz
            fts.Append(ft)

            ft = New DocumentFormat.OpenXml.Spreadsheet.Font()
            ftn = New FontName()
            ftn.Val = StringValue.FromString("Arial")
            ftsz = New FontSize()
            ftsz.Val = DoubleValue.FromDouble(12)
            ft.FontName = ftn
            ft.FontSize = ftsz
            ft.Bold = New Bold()

            fts.Append(ft)

            fts.Count = UInt32Value.FromUInt32(CUInt(fts.ChildElements.Count))
            'Fonts ------------------------------------------------------------------------------------------------------------



            'Fills ------------------------------------------------------------------------------------------------------------
            Dim fills As New Fills()
            '0
            Dim fill As Fill
            Dim patternFill As PatternFill
            fill = New Fill()
            patternFill = New PatternFill()
            patternFill.PatternType = PatternValues.None
            fill.PatternFill = patternFill
            fills.Append(fill)

            '1
            fill = New Fill()
            patternFill = New PatternFill()
            patternFill.PatternType = PatternValues.Gray125
            fill.PatternFill = patternFill
            fills.Append(fill)

            '2
            fill = New Fill()
            patternFill = New PatternFill()
            patternFill.PatternType = PatternValues.Solid
            patternFill.ForegroundColor = New ForegroundColor()
            patternFill.ForegroundColor.Rgb = HexBinaryValue.FromString("00ff9728")
            patternFill.BackgroundColor = New BackgroundColor()
            patternFill.BackgroundColor.Rgb = patternFill.ForegroundColor.Rgb
            fill.PatternFill = patternFill
            fills.Append(fill)

            '3
            fill = New Fill()
            patternFill = New PatternFill()
            patternFill.PatternType = PatternValues.Solid
            patternFill.ForegroundColor = New ForegroundColor()
            patternFill.ForegroundColor.Rgb = HexBinaryValue.FromString("FFFFFF00")
            patternFill.BackgroundColor = New BackgroundColor()
            patternFill.BackgroundColor.Rgb = patternFill.ForegroundColor.Rgb
            fill.PatternFill = patternFill
            fills.Append(fill)

            fills.Count = UInt32Value.FromUInt32(CUInt(fills.ChildElements.Count))
            'Fills ------------------------------------------------------------------------------------------------------------



            'Borders ------------------------------------------------------------------------------------------------------------
            Dim borders As New Borders()
            Dim border As New Border()
            border.LeftBorder = New LeftBorder()
            border.RightBorder = New RightBorder()
            border.TopBorder = New TopBorder()
            border.BottomBorder = New BottomBorder()
            border.DiagonalBorder = New DiagonalBorder()
            borders.Append(border)

            'Boarder Index 1
            border = New Border()
            border.LeftBorder = New LeftBorder()
            border.LeftBorder.Style = BorderStyleValues.Thin
            border.RightBorder = New RightBorder()
            border.RightBorder.Style = BorderStyleValues.Thin
            border.TopBorder = New TopBorder()
            border.TopBorder.Style = BorderStyleValues.Thin
            border.BottomBorder = New BottomBorder()
            border.BottomBorder.Style = BorderStyleValues.Thin
            border.DiagonalBorder = New DiagonalBorder()
            borders.Append(border)

            'Boarder Index 2
            border = New Border()
            border.LeftBorder = New LeftBorder()
            border.RightBorder = New RightBorder()
            border.TopBorder = New TopBorder()
            border.TopBorder.Style = BorderStyleValues.Thin
            border.BottomBorder = New BottomBorder()
            border.BottomBorder.Style = BorderStyleValues.Thin
            border.DiagonalBorder = New DiagonalBorder()
            borders.Append(border)

            borders.Count = UInt32Value.FromUInt32(CUInt(borders.ChildElements.Count))
            'Borders ------------------------------------------------------------------------------------------------------------


            'CellStyleFormats ------------------------------------------------------------------------------------------------------------			
            Dim csfs As New CellStyleFormats()
            Dim cf As New CellFormat()
            cf.NumberFormatId = 0
            cf.FontId = 0
            cf.FillId = 0
            cf.BorderId = 0
            csfs.Append(cf)
            csfs.Count = UInt32Value.FromUInt32(CUInt(csfs.ChildElements.Count))
            'CellStyleFormats ------------------------------------------------------------------------------------------------------------



            'NumberingFormats ------------------------------------------------------------------------------------------------------------
            'Dim iExcelIndex As UInteger = 164
            Dim iExcelIndex As UInt32 = 164
            Dim nfs As New NumberingFormats()

            Dim nfDateTime As New NumberingFormat()
            nfDateTime.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex)
            nfDateTime.FormatCode = StringValue.FromString("dd/mm/yyyy hh:mm:ss")
            nfs.Append(nfDateTime)

            iExcelIndex += 1
            Dim nf4decimal As New NumberingFormat()
            nf4decimal.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex)
            nf4decimal.FormatCode = StringValue.FromString("#,##0.0000")
            nfs.Append(nf4decimal)

            ' #,##0.00 is also Excel style index 4
            iExcelIndex += 1
            Dim nf2decimal As New NumberingFormat()
            nf2decimal.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex)
            nf2decimal.FormatCode = StringValue.FromString("#,##0.00")
            nfs.Append(nf2decimal)

            ' @ is also Excel style index 49
            iExcelIndex += 1
            Dim nfForcedText As New NumberingFormat()
            nfForcedText.NumberFormatId = UInt32Value.FromUInt32(iExcelIndex)
            nfForcedText.FormatCode = StringValue.FromString("@")
            nfs.Append(nfForcedText)

            nfs.Count = UInt32Value.FromUInt32(CUInt(nfs.ChildElements.Count))
            'NumberingFormats ------------------------------------------------------------------------------------------------------------



            'CellFormats ------------------------------------------------------------------------------------------------------------
            Dim cfs As New CellFormats()
            cf = New CellFormat()
            cf.NumberFormatId = 0
            cf.FontId = 0
            cf.FillId = 0
            cf.BorderId = 0
            cf.FormatId = 0
            cfs.Append(cf)

            ' index 1
            ' Format dd/mm/yyyy
            cf = New CellFormat()
            cf.NumberFormatId = 14
            cf.FontId = 0
            cf.FillId = 0
            cf.BorderId = 0
            cf.FormatId = 0
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(True)
            cfs.Append(cf)

            ' index 2
            ' Format #,##0.00
            cf = New CellFormat()
            cf.NumberFormatId = 4
            cf.FontId = 0
            cf.FillId = 0
            cf.BorderId = 0
            cf.FormatId = 0
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(True)
            cfs.Append(cf)

            ' index 3
            cf = New CellFormat()
            cf.NumberFormatId = nfDateTime.NumberFormatId
            cf.FontId = 0
            cf.FillId = 0
            cf.BorderId = 0
            cf.FormatId = 0
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(True)
            cfs.Append(cf)

            ' index 4
            cf = New CellFormat()
            cf.NumberFormatId = nf4decimal.NumberFormatId
            cf.FontId = 0
            cf.FillId = 0
            cf.BorderId = 0
            cf.FormatId = 0
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(True)
            cfs.Append(cf)

            ' index 5
            cf = New CellFormat()
            cf.NumberFormatId = nf2decimal.NumberFormatId
            cf.FontId = 0
            cf.FillId = 0
            cf.BorderId = 0
            cf.FormatId = 0
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(True)
            cfs.Append(cf)

            ' index 6
            cf = New CellFormat()
            cf.NumberFormatId = nfForcedText.NumberFormatId
            cf.FontId = 0
            cf.FillId = 0
            cf.BorderId = 0
            cf.FormatId = 0
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(True)
            cfs.Append(cf)

            ' index 7
            ' Header text
            cf = New CellFormat()
            cf.NumberFormatId = nfForcedText.NumberFormatId
            cf.FontId = 1
            cf.FillId = 0
            cf.BorderId = 0
            cf.FormatId = 0
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(True)
            cfs.Append(cf)

            ' index 8
            ' column text
            cf = New CellFormat()
            cf.NumberFormatId = nfForcedText.NumberFormatId
            cf.FontId = 0
            cf.FillId = 0
            cf.BorderId = 1
            cf.FormatId = 0
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(True)
            cfs.Append(cf)

            ' index 9
            ' coloured 2 decimal text
            cf = New CellFormat()
            cf.NumberFormatId = nf2decimal.NumberFormatId
            cf.FontId = 0
            cf.FillId = 2
            cf.BorderId = 2
            cf.FormatId = 0
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(True)
            cfs.Append(cf)

            ' index 10
            ' coloured column text
            cf = New CellFormat()
            cf.NumberFormatId = nfForcedText.NumberFormatId
            cf.FontId = 0
            cf.FillId = 2
            cf.BorderId = 2
            cf.FormatId = 0
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(True)
            cfs.Append(cf)

            ' index 11
            cf = New CellFormat()
            cf.NumberFormatId = nfForcedText.NumberFormatId
            cf.FontId = 0
            cf.FillId = 3
            cf.BorderId = 1
            cf.FormatId = 0
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(True)
            cfs.Append(cf)

            ' index 12
            ' Format #,##0.00
            cf = New CellFormat()
            cf.NumberFormatId = 4
            cf.FontId = 0
            cf.FillId = 2
            cf.BorderId = 1
            cf.FormatId = 0
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(True)
            cfs.Append(cf)

            ' index 13
            ' Format #,##0.00
            cf = New CellFormat()
            cf.NumberFormatId = 4
            cf.FontId = 2
            cf.FillId = 0
            cf.BorderId = 1
            cf.FormatId = 0
            cf.Alignment = New Alignment()
            cf.Alignment.Horizontal = HorizontalAlignmentValues.Right
            cf.ApplyAlignment = BooleanValue.FromBoolean(True)
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(True)
            cfs.Append(cf)

            ' index 14 à droite cadré gras
            cf = New CellFormat()
            cf.NumberFormatId = 0
            cf.FontId = 2
            cf.FillId = 0
            cf.BorderId = 1
            cf.FormatId = 0
            cf.Alignment = New Alignment()
            cf.Alignment.Horizontal = HorizontalAlignmentValues.Right
            cf.ApplyAlignment = BooleanValue.FromBoolean(True)
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(True)
            cfs.Append(cf)

            ' index 15 centré cadré gras
            cf = New CellFormat()
            cf.NumberFormatId = 0
            cf.FontId = 2
            cf.FillId = 0
            cf.BorderId = 1
            cf.FormatId = 0
            cf.Alignment = New Alignment() With {.WrapText = True}
            cf.Alignment.Horizontal = HorizontalAlignmentValues.Center
            cf.Alignment.Vertical = VerticalAlignmentValues.Center
            cf.ApplyAlignment = BooleanValue.FromBoolean(True)
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(True)
            cfs.Append(cf)

            ' index 16 centré cadré wrap=false
            cf = New CellFormat()
            cf.NumberFormatId = 0
            cf.FontId = 0
            cf.FillId = 0
            cf.BorderId = 1
            cf.FormatId = 0
            cf.Alignment = New Alignment() With {.WrapText = False}
            cf.Alignment.Horizontal = HorizontalAlignmentValues.Center
            cf.Alignment.Vertical = VerticalAlignmentValues.Center
            cf.ApplyAlignment = BooleanValue.FromBoolean(True)
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(True)
            cfs.Append(cf)

            ' index 17 centré cadré wrap=true
            cf = New CellFormat()
            cf.NumberFormatId = 0
            cf.FontId = 0
            cf.FillId = 0
            cf.BorderId = 1
            cf.FormatId = 0
            cf.Alignment = New Alignment() With {.WrapText = True}
            cf.Alignment.Horizontal = HorizontalAlignmentValues.Center
            cf.Alignment.Vertical = VerticalAlignmentValues.Center
            cf.ApplyAlignment = BooleanValue.FromBoolean(True)
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(True)
            cfs.Append(cf)

            cfs.Count = UInt32Value.FromUInt32(CUInt(cfs.ChildElements.Count))
            'CellFormats ------------------------------------------------------------------------------------------------------------



            'CellStyles ------------------------------------------------------------------------------------------------------------
            Dim css As New CellStyles()
            Dim cs As New CellStyle()
            cs.Name = StringValue.FromString("Normal")
            cs.FormatId = 0
            cs.BuiltinId = 0
            css.Append(cs)
            css.Count = UInt32Value.FromUInt32(CUInt(css.ChildElements.Count))


            Dim dfs As New DifferentialFormats()
            dfs.Count = 0

            'CellStyles ------------------------------------------------------------------------------------------------------------


            'TableStyles ------------------------------------------------------------------------------------------------------------
            Dim tss As New TableStyles()
            tss.Count = 0
            tss.DefaultTableStyle = StringValue.FromString("TableStyleMedium9")
            tss.DefaultPivotStyle = StringValue.FromString("PivotStyleLight16")
            'TableStyles ------------------------------------------------------------------------------------------------------------



            ' L'ordre est omporant 
            Me.Append(nfs)
            Me.Append(fts)
            Me.Append(fills)
            Me.Append(borders)
            Me.Append(csfs)
            Me.Append(cfs)
            Me.Append(css)
            Me.Append(dfs)
            Me.Append(tss)
        End Sub
    End Class


    Public Class ColumnData
        Inherits Column

        Public Sub New(StartColumnIndex As UInt32, EndColumnIndex As UInt32, ColumnWidth As Double)
            MyBase.New()
            Me.Min = StartColumnIndex
            Me.Max = EndColumnIndex
            Me.Width = ColumnWidth
            Me.CustomWidth = True
        End Sub
    End Class







End Namespace

