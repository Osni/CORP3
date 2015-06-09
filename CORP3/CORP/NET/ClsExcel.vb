Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.IO
Imports System.Runtime.InteropServices

Namespace CORP3.NET
    Public Class ClsExcel
        ' Methods
        Public Sub New(Optional ByVal pXLSGenerateType As GenerateType = 1)
            Me.XLSGenerateType = GenerateType.ToFile
            Me.strFileName = "arq.xls"
            Me.BEG_FILE_MARKER.opcode = 9
            Me.BEG_FILE_MARKER.length = 4
            Me.BEG_FILE_MARKER.version = 2
            Me.BEG_FILE_MARKER.ftype = 10
            Me.END_FILE_MARKER.opcode = 10
            Me.XLSGenerateType = pXLSGenerateType
            Me.strFileName = ("ARQ_" & Guid.NewGuid.ToString & ".xls")
            Me.CreateFile(Me.strFileName)
        End Sub

        Public Sub New(ByVal _FileName As String, Optional ByVal pXLSGenerateType As GenerateType = 1)
            Me.XLSGenerateType = GenerateType.ToFile
            Me.strFileName = "arq.xls"
            Me.BEG_FILE_MARKER.opcode = 9
            Me.BEG_FILE_MARKER.length = 4
            Me.BEG_FILE_MARKER.version = 2
            Me.BEG_FILE_MARKER.ftype = 10
            Me.END_FILE_MARKER.opcode = 10
            Me.XLSGenerateType = pXLSGenerateType
            If (_FileName <> String.Empty) Then
                Me.strFileName = _FileName
            Else
                Me.strFileName = ("ARQ_" & Guid.NewGuid.ToString & ".xls")
            End If
            Me.CreateFile(Me.strFileName)
        End Sub

        Private Function CreateFile(ByVal FileName As String) As Boolean
            If (Me.XLSGenerateType = GenerateType.ToFile) Then
                If File.Exists(FileName) Then
                    File.Delete(FileName)
                End If
                Me.strm = New FileStream(FileName, FileMode.CreateNew, FileAccess.Write, FileShare.Read)
            Else
                Me.strm = New MemoryStream
            End If
            Me.writer = New BinaryWriter(Me.strm)
            Me.writer.Write(Me.BEG_FILE_MARKER.opcode)
            Me.writer.Write(Me.BEG_FILE_MARKER.length)
            Me.writer.Write(Me.BEG_FILE_MARKER.version)
            Me.writer.Write(Me.BEG_FILE_MARKER.ftype)
            Me.WriteDefaultFormats()
            Me.HorizPageBreakRows = New Short(1 - 1) {}
            Me.NumHorizPageBreaks = 0
            Return True
        End Function

        Public Function InsertHorizPageBreak(ByVal lrow As Integer) As Boolean
            Dim num As Integer
            If (lrow > &H7FFF) Then
                num = (lrow - &H10000)
            Else
                num = (lrow - 1)
            End If
            Me.NumHorizPageBreaks = CShort((Me.NumHorizPageBreaks + 1))
            Me.HorizPageBreakRows = DirectCast(Utils.CopyArray(DirectCast(Me.HorizPageBreakRows, Array), New Short((Me.NumHorizPageBreaks + 1) - 1) {}), Short())
            Me.HorizPageBreakRows(Me.NumHorizPageBreaks) = CShort(num)
            Return True
        End Function

        Public Function SetColumnWidth(ByVal FirstColumn As Byte, ByVal LastColumn As Byte, ByVal WidthValue As Short) As Boolean
            Dim colwidth_record As COLWIDTH_RECORD
            colwidth_record.opcode = &H24
            colwidth_record.length = 4
            colwidth_record.col1 = CByte((FirstColumn - 1))
            colwidth_record.col2 = CByte((LastColumn - 1))
            colwidth_record.ColumnWidth = CShort((WidthValue * &H100))
            Me.writer.Write(colwidth_record.opcode)
            Me.writer.Write(colwidth_record.length)
            Me.writer.Write(colwidth_record.col1)
            Me.writer.Write(colwidth_record.col2)
            Me.writer.Write(colwidth_record.ColumnWidth)
            Return True
        End Function

        Public Function SetDefaultRowHeight(ByVal HeightValue As Short) As Boolean
            Dim def_rowheight_record As DEF_ROWHEIGHT_RECORD
            def_rowheight_record.opcode = &H25
            def_rowheight_record.length = 2
            def_rowheight_record.RowHeight = CShort((HeightValue * 20))
            Me.writer.Write(def_rowheight_record.opcode)
            Me.writer.Write(def_rowheight_record.length)
            Me.writer.Write(def_rowheight_record.RowHeight)
            Return True
        End Function

        Public Function SetEOF() As Boolean
            If (Me.NumHorizPageBreaks > 0) Then
                Dim num5 As Integer = Information.LBound(Me.HorizPageBreakRows, 1)
                Dim i As Integer = Information.UBound(Me.HorizPageBreakRows, 1)
                Do While (i >= num5)
                    Dim num6 As Integer = i
                    Dim k As Integer = (Information.LBound(Me.HorizPageBreakRows, 1) + 1)
                    Do While (k <= num6)
                        If (Me.HorizPageBreakRows((k - 1)) > Me.HorizPageBreakRows(k)) Then
                            Dim num3 As Integer = Me.HorizPageBreakRows((k - 1))
                            Me.HorizPageBreakRows((k - 1)) = Me.HorizPageBreakRows(k)
                            Me.HorizPageBreakRows(k) = CShort(num3)
                        End If
                        k += 1
                    Loop
                    i = (i + -1)
                Loop
                Me.HORIZ_PAGE_BREAK.opcode = &H1B
                Me.HORIZ_PAGE_BREAK.length = CShort((2 + (Me.NumHorizPageBreaks * 2)))
                Me.HORIZ_PAGE_BREAK.NumPageBreaks = Me.NumHorizPageBreaks
                Me.writer.Write(Me.HORIZ_PAGE_BREAK.opcode)
                Me.writer.Write(Me.HORIZ_PAGE_BREAK.length)
                Me.writer.Write(Me.HORIZ_PAGE_BREAK.NumPageBreaks)
                Dim num7 As Integer = Information.UBound(Me.HorizPageBreakRows, 1)
                Dim j As Integer = 1
                Do While (j <= num7)
                    Me.writer.Write(Me.HorizPageBreakRows(j))
                    j += 1
                Loop
            End If
            Me.writer.Write(Me.END_FILE_MARKER.opcode)
            Me.writer.Write(Me.END_FILE_MARKER.length)
            Return True
        End Function

        Public Function SetFilePassword(ByVal PasswordText As String) As Boolean
            Dim password_record As PASSWORD_RECORD
            Dim num2 As Integer = Strings.Len(PasswordText)
            password_record.opcode = &H2F
            password_record.length = CShort(num2)
            Me.writer.Write(password_record.opcode)
            Me.writer.Write(password_record.length)
            Dim num4 As Integer = num2
            Dim i As Integer = 1
            Do While (i <= num4)
                Dim num As Byte = CByte(Strings.Asc(Strings.Mid(PasswordText, i, 1)))
                Me.writer.Write(num)
                i += 1
            Loop
            Return True
        End Function

        Public Function SetFont(ByVal FontName As String, ByVal FontHeight As Short, ByVal FontFormat As FontFormatting) As Boolean
            Dim font_record As FONT_RECORD
            Dim str As String = Conversions.ToString(Strings.Len(FontName))
            font_record.opcode = &H31
            font_record.length = CShort(Math.Round(CDbl((5 + Conversions.ToDouble(str)))))
            font_record.FontHeight = CShort((FontHeight * 20))
            font_record.FontAttributes1 = CByte(FontFormat)
            font_record.FontAttributes2 = 0
            font_record.FontNameLength = CByte(Strings.Len(FontName))
            Me.writer.Write(font_record.opcode)
            Me.writer.Write(font_record.length)
            Me.writer.Write(font_record.FontHeight)
            Me.writer.Write(font_record.FontAttributes1)
            Me.writer.Write(font_record.FontAttributes2)
            Me.writer.Write(font_record.FontNameLength)
            Dim num3 As Integer = Conversions.ToInteger(str)
            Dim i As Integer = 1
            Do While (i <= num3)
                Dim num As Byte = CByte(Strings.Asc(Strings.Mid(FontName, i, 1)))
                Me.writer.Write(num)
                i += 1
            Loop
            Return True
        End Function

        Public Function SetFooter(ByVal FooterText As String) As Boolean
            Dim header_footer_record As HEADER_FOOTER_RECORD
            Dim num2 As Integer = Strings.Len(FooterText)
            header_footer_record.opcode = &H15
            header_footer_record.length = CShort((1 + num2))
            header_footer_record.TextLength = CByte(Strings.Len(FooterText))
            Me.writer.Write(header_footer_record.opcode)
            Me.writer.Write(header_footer_record.length)
            Me.writer.Write(header_footer_record.TextLength)
            Dim num4 As Integer = num2
            Dim i As Integer = 1
            Do While (i <= num4)
                Dim num As Byte = CByte(Strings.Asc(Strings.Mid(FooterText, i, 1)))
                Me.writer.Write(num)
                i += 1
            Loop
            Return True
        End Function

        Public Function SetHeader(ByVal HeaderText As String) As Boolean
            Dim header_footer_record As HEADER_FOOTER_RECORD
            Dim num2 As Integer = Strings.Len(HeaderText)
            header_footer_record.opcode = 20
            header_footer_record.length = CShort((1 + num2))
            header_footer_record.TextLength = CByte(Strings.Len(HeaderText))
            Me.writer.Write(header_footer_record.opcode)
            Me.writer.Write(header_footer_record.length)
            Me.writer.Write(header_footer_record.TextLength)
            Dim num4 As Integer = num2
            Dim i As Integer = 1
            Do While (i <= num4)
                Dim num As Byte = CByte(Strings.Asc(Strings.Mid(HeaderText, i, 1)))
                Me.writer.Write(num)
                i += 1
            Loop
            Return True
        End Function

        Public Function SetMargin(ByVal Margin As MarginTypes, ByVal MarginValue As Double) As Boolean
            Dim margin_record_layout As MARGIN_RECORD_LAYOUT
            margin_record_layout.opcode = CShort(Margin)
            margin_record_layout.length = 8
            margin_record_layout.MarginValue = MarginValue
            Me.writer.Write(margin_record_layout.opcode)
            Me.writer.Write(margin_record_layout.length)
            Me.writer.Write(margin_record_layout.MarginValue)
            Return True
        End Function

        Public Function SetRowHeight(ByVal lRow As Integer, ByVal HeightValue As Short) As Boolean
            Dim num As Integer
            Dim row_height_record As ROW_HEIGHT_RECORD
            If (lRow > &H7FFF) Then
                num = (lRow - &H10000)
            Else
                num = (lRow - 1)
            End If
            row_height_record.opcode = 8
            row_height_record.length = &H10
            row_height_record.RowNumber = CShort(num)
            row_height_record.FirstColumn = 0
            row_height_record.LastColumn = &H100
            row_height_record.RowHeight = CShort((HeightValue * 20))
            row_height_record.internal = 0
            row_height_record.DefaultAttributes = 0
            row_height_record.FileOffset = 0
            row_height_record.rgbAttr1 = 0
            row_height_record.rgbAttr2 = 0
            row_height_record.rgbAttr3 = 0
            Me.writer.Write(row_height_record.opcode)
            Me.writer.Write(row_height_record.length)
            Me.writer.Write(row_height_record.RowNumber)
            Me.writer.Write(row_height_record.FirstColumn)
            Me.writer.Write(row_height_record.LastColumn)
            Me.writer.Write(row_height_record.RowHeight)
            Me.writer.Write(row_height_record.internal)
            Me.writer.Write(row_height_record.DefaultAttributes)
            Me.writer.Write(row_height_record.FileOffset)
            Me.writer.Write(row_height_record.rgbAttr1)
            Me.writer.Write(row_height_record.rgbAttr2)
            Me.writer.Write(row_height_record.rgbAttr3)
            Return True
        End Function

        Public Function WriteDate(ByVal CellFontUsed As CellFont, ByVal lRow As Integer, ByVal lCol As Integer, ByVal value As Object, Optional ByVal CellFormat As Integer = 20, Optional ByVal Alignment As CellAlignment = 3) As Boolean
            Dim num As Integer
            Dim number As tNumber
            Dim num2 As Integer
            If (lRow > &H7FFF) Then
                num2 = (lRow - &H10000)
            Else
                num2 = (lRow - 1)
            End If
            If (lCol > &H7FFF) Then
                num = (lCol - &H10000)
            Else
                num = (lCol - 1)
            End If
            number.opcode = 3
            number.length = 15
            number.Row = CShort(num2)
            number.col = CShort(num)
            number.rgbAttr1 = 0
            number.rgbAttr2 = CByte((CellFontUsed + CellFormat))
            number.rgbAttr3 = CByte(Alignment)
            If (value.GetType.ToString = "System.TimeSpan") Then
                number.NumberValue = Conversions.ToDate(value.ToString).ToOADate
            Else
                number.NumberValue = Conversions.ToDouble(NewLateBinding.LateGet(value, Nothing, "ToOAdate", New Object(0 - 1) {}, Nothing, Nothing, Nothing))
            End If
            Me.writer.Write(number.opcode)
            Me.writer.Write(number.length)
            Me.writer.Write(number.Row)
            Me.writer.Write(number.col)
            Me.writer.Write(number.rgbAttr1)
            Me.writer.Write(number.rgbAttr2)
            Me.writer.Write(number.rgbAttr3)
            Me.writer.Write(number.NumberValue)
            Return True
        End Function

        Public Function WriteDefaultFormats() As Short
            Dim format_count_record As FORMAT_COUNT_RECORD
            Dim num3 As Short
            Dim array As String() = New String(&H18 - 1) {}
            Dim str As String = """"
            array(0) = "General"
            array(1) = "0"
            array(2) = "0.00"
            array(3) = "#,##0"
            array(4) = "#,##0.00"
            array(5) = String.Concat(New String() {"#,##0\ ", str, "$", str, ";\-#,##0\ ", str, "$", str})
            array(6) = String.Concat(New String() {"#,##0\ ", str, "$", str, ";[Red]\-#,##0\ ", str, "$", str})
            array(7) = String.Concat(New String() {"#,##0.00\ ", str, "$", str, ";\-#,##0.00\ ", str, "$", str})
            array(8) = String.Concat(New String() {"#,##0.00\ ", str, "$", str, ";[Red]\-#,##0.00\ ", str, "$", str})
            array(9) = "0%"
            array(10) = "0.00%"
            array(11) = "0.00E+00"
            array(12) = "dd/mm/yy"
            array(13) = "dd/\ mmm\ yy"
            array(14) = "dd/\ mmm"
            array(15) = "mmm\ yy"
            array(&H10) = "h:mm\ AM/PM"
            array(&H11) = "h:mm:ss\ AM/PM"
            array(&H12) = "hh:mm"
            array(&H13) = "hh:mm:ss"
            array(20) = "dd/mm/yy\ hh:mm"
            array(&H15) = "##0.0E+0"
            array(&H16) = "mm:ss"
            array(&H17) = "@"
            format_count_record.opcode = &H1F
            format_count_record.length = 2
            format_count_record.Count = CShort(Information.UBound(array, 1))
            Me.writer.Write(format_count_record.opcode)
            Me.writer.Write(format_count_record.length)
            Me.writer.Write(format_count_record.Count)
            Dim num6 As Integer = Information.UBound(array, 1)
            Dim i As Integer = Information.LBound(array, 1)
            Do While (i <= num6)
                Dim format_record As FORMAT_RECORD
                Dim num As Integer = Strings.Len(array(i))
                format_record.opcode = 30
                format_record.length = CShort((num + 1))
                format_record.FormatLenght = CByte(num)
                Me.writer.Write(format_record.opcode)
                Me.writer.Write(format_record.length)
                Me.writer.Write(format_record.FormatLenght)
                Dim num7 As Integer = num
                Dim j As Integer = 1
                Do While (j <= num7)
                    Dim num5 As Byte = CByte(Strings.Asc(Strings.Mid(array(i), j, 1)))
                    Me.writer.Write(num5)
                    j += 1
                Loop
                i += 1
            Loop
            Return num3
        End Function

        Public Function WriteInteger(ByVal CellFontUsed As CellFont, ByVal lRow As Integer, ByVal lCol As Integer, ByVal value As Object, Optional ByVal CellFormat As Integer = 0, Optional ByVal Alignment As CellAlignment = 3) As Boolean
            Dim num As Integer
            Dim tInt As tInteger
            Dim num2 As Integer
            If (lRow > &H7FFF) Then
                num2 = (lRow - &H10000)
            Else
                num2 = (lRow - 1)
            End If
            If (lCol > &H7FFF) Then
                num = (lCol - &H10000)
            Else
                num = (lCol - 1)
            End If
            [tInt].opcode = 2
            [tInt].length = 9
            [tInt].Row = CShort(num2)
            [tInt].col = CShort(num)
            [tInt].rgbAttr1 = 0
            [tInt].rgbAttr2 = CByte((CellFontUsed + CellFormat))
            [tInt].rgbAttr3 = CByte(Alignment)
            [tInt].intValue = CShort(Conversions.ToInteger(value))
            Me.writer.Write([tInt].opcode)
            Me.writer.Write([tInt].length)
            Me.writer.Write([tInt].Row)
            Me.writer.Write([tInt].col)
            Me.writer.Write([tInt].rgbAttr1)
            Me.writer.Write([tInt].rgbAttr2)
            Me.writer.Write([tInt].rgbAttr3)
            Me.writer.Write([tInt].intValue)
            Return True
        End Function

        Public Function WriteNumber(ByVal CellFontUsed As CellFont, ByVal lRow As Integer, ByVal lCol As Integer, ByVal value As Object, Optional ByVal CellFormat As Integer = 0, Optional ByVal Alignment As CellAlignment = 3) As Boolean
            Dim num As Integer
            Dim number As tNumber
            Dim num2 As Integer
            If (lRow > &H7FFF) Then
                num2 = (lRow - &H10000)
            Else
                num2 = (lRow - 1)
            End If
            If (lCol > &H7FFF) Then
                num = (lCol - &H10000)
            Else
                num = (lCol - 1)
            End If
            number.opcode = 3
            number.length = 15
            number.Row = CShort(num2)
            number.col = CShort(num)
            number.rgbAttr1 = 0
            number.rgbAttr2 = CByte((CellFontUsed + CellFormat))
            number.rgbAttr3 = CByte(Alignment)
            number.NumberValue = Conversions.ToDouble(value)
            Me.writer.Write(number.opcode)
            Me.writer.Write(number.length)
            Me.writer.Write(number.Row)
            Me.writer.Write(number.col)
            Me.writer.Write(number.rgbAttr1)
            Me.writer.Write(number.rgbAttr2)
            Me.writer.Write(number.rgbAttr3)
            Me.writer.Write(number.NumberValue)
            Return True
        End Function

        Public Function WriteText(ByVal CellFontUsed As CellFont, ByVal lRow As Integer, ByVal lCol As Integer, ByVal value As Object, Optional ByVal CellFormat As Integer = 0, Optional ByVal Alignment As CellAlignment = 1) As Boolean
            Dim num2 As Integer
            Dim num4 As Integer
            Dim text As tText
            If (lRow > &H7FFF) Then
                num4 = (lRow - &H10000)
            Else
                num4 = (lRow - 1)
            End If
            If (lCol > &H7FFF) Then
                num2 = (lCol - &H10000)
            Else
                num2 = (lCol - 1)
            End If
            Dim expression As String = Conversions.ToString(Operators.ConcatenateObject("", value))
            Dim num3 As Integer = Strings.Len(expression)
            [text].opcode = 4
            [text].length = 10
            [text].TextLength = CByte(num3)
            [text].length = CShort((8 + num3))
            [text].Row = CShort(num4)
            [text].col = CShort(num2)
            [text].rgbAttr1 = 0
            [text].rgbAttr2 = CByte((CellFontUsed + CellFormat))
            [text].rgbAttr3 = CByte(Alignment)
            Me.writer.Write([text].opcode)
            Me.writer.Write([text].length)
            Me.writer.Write([text].Row)
            Me.writer.Write([text].col)
            Me.writer.Write([text].rgbAttr1)
            Me.writer.Write([text].rgbAttr2)
            Me.writer.Write([text].rgbAttr3)
            Me.writer.Write([text].TextLength)
            Dim num6 As Integer = num3
            Dim i As Integer = 1
            Do While (i <= num6)
                Dim num As Byte = CByte(Strings.Asc(Strings.Mid(expression, i, 1)))
                Me.writer.Write(num)
                i += 1
            Loop
            Return True
        End Function

        Public Function WriteValue(ByVal ValueType As ValueTypes, ByVal CellFontUsed As CellFont, ByVal Alignment As CellAlignment, ByVal HiddenLocked As CellHiddenLocked, ByVal lRow As Integer, ByVal lCol As Integer, ByVal value As Object, Optional ByVal CellFormat As Integer = 0) As Boolean
            Dim num As Integer
            Dim num2 As Integer
            Dim number As tNumber
            If (lRow > &H7FFF) Then
                num2 = (lRow - &H10000)
            Else
                num2 = (lRow - 1)
            End If
            If (lCol > &H7FFF) Then
                num = (lCol - &H10000)
            Else
                num = (lCol - 1)
            End If
            Select Case ValueType
                Case ValueTypes.xlsInteger
                    Dim tInt As tInteger
                    [tInt].opcode = 2
                    [tInt].length = 9
                    [tInt].Row = CShort(num2)
                    [tInt].col = CShort(num)
                    [tInt].rgbAttr1 = CByte(HiddenLocked)
                    [tInt].rgbAttr2 = CByte((CellFontUsed + CellFormat))
                    [tInt].rgbAttr3 = CByte(Alignment)
                    [tInt].intValue = CShort(Conversions.ToInteger(value))
                    Me.writer.Write([tInt].opcode)
                    Me.writer.Write([tInt].length)
                    Me.writer.Write([tInt].Row)
                    Me.writer.Write([tInt].col)
                    Me.writer.Write([tInt].rgbAttr1)
                    Me.writer.Write([tInt].rgbAttr2)
                    Me.writer.Write([tInt].rgbAttr3)
                    Me.writer.Write([tInt].intValue)
                    GoTo Label_038A
                Case ValueTypes.xlsNumber
                    number.opcode = 3
                    number.length = 15
                    number.Row = CShort(num2)
                    number.col = CShort(num)
                    number.rgbAttr1 = CByte(HiddenLocked)
                    number.rgbAttr2 = CByte((CellFontUsed + CellFormat))
                    number.rgbAttr3 = CByte(Alignment)
                    If (value.GetType.ToString <> "System.DateTime") Then
                        number.NumberValue = Conversions.ToDouble(value)
                        Exit Select
                    End If
                    number.NumberValue = Conversions.ToDouble(NewLateBinding.LateGet(value, Nothing, "ToOAdate", New Object(0 - 1) {}, Nothing, Nothing, Nothing))
                    Exit Select
                Case ValueTypes.xlsText
                    Dim text As tText
                    Dim expression As String = Conversions.ToString(Operators.ConcatenateObject("", value))
                    Dim num4 As Integer = Strings.Len(expression)
                    [text].opcode = 4
                    [text].length = 10
                    [text].TextLength = CByte(num4)
                    [text].length = CShort((8 + num4))
                    [text].Row = CShort(num2)
                    [text].col = CShort(num)
                    [text].rgbAttr1 = CByte(HiddenLocked)
                    [text].rgbAttr2 = CByte((CellFontUsed + CellFormat))
                    [text].rgbAttr3 = CByte(Alignment)
                    Me.writer.Write([text].opcode)
                    Me.writer.Write([text].length)
                    Me.writer.Write([text].Row)
                    Me.writer.Write([text].col)
                    Me.writer.Write([text].rgbAttr1)
                    Me.writer.Write([text].rgbAttr2)
                    Me.writer.Write([text].rgbAttr3)
                    Me.writer.Write([text].TextLength)
                    Dim num6 As Integer = num4
                    Dim i As Integer = 1
                    Do While (i <= num6)
                        Dim num3 As Byte = CByte(Strings.Asc(Strings.Mid(expression, i, 1)))
                        Me.writer.Write(num3)
                        i += 1
                    Loop
                    GoTo Label_038A
                Case Else
                    GoTo Label_038A
            End Select
            Me.writer.Write(number.opcode)
            Me.writer.Write(number.length)
            Me.writer.Write(number.Row)
            Me.writer.Write(number.col)
            Me.writer.Write(number.rgbAttr1)
            Me.writer.Write(number.rgbAttr2)
            Me.writer.Write(number.rgbAttr3)
            Me.writer.Write(number.NumberValue)
Label_038A:
            Return True
        End Function


        ' Properties
        Public Property FileName() As String
            Get
                Return Me.strFileName
            End Get
            Set(ByVal value As String)
                Me.strFileName = value
            End Set
        End Property

        Public ReadOnly Property GetStream() As Stream
            Get
                Return Me.strm
            End Get
        End Property

        Public WriteOnly Property PrintGridLines() As Boolean
            Set(ByVal value As Boolean)
                Dim print_gridlines_record As PRINT_GRIDLINES_RECORD
                print_gridlines_record.opcode = &H2B
                print_gridlines_record.length = 2
                print_gridlines_record.PrintFlag = Conversions.ToShort(Interaction.IIf(value, 1, 0))
                Me.writer.Write(print_gridlines_record.opcode)
                Me.writer.Write(print_gridlines_record.length)
                Me.writer.Write(print_gridlines_record.PrintFlag)
            End Set
        End Property

        Public WriteOnly Property ProtectSpreadsheet() As Boolean
            Set(ByVal value As Boolean)
                Dim protect_spreadsheet_record As PROTECT_SPREADSHEET_RECORD
                protect_spreadsheet_record.opcode = &H12
                protect_spreadsheet_record.length = 2
                protect_spreadsheet_record.Protect = Conversions.ToShort(Interaction.IIf(value, 1, 0))
                Me.writer.Write(protect_spreadsheet_record.opcode)
                Me.writer.Write(protect_spreadsheet_record.length)
                Me.writer.Write(protect_spreadsheet_record.Protect)
            End Set
        End Property


        ' Fields
        Private BEG_FILE_MARKER As BEG_FILE_RECORD
        Private END_FILE_MARKER As END_FILE_RECORD
        Private HORIZ_PAGE_BREAK As HPAGE_BREAK_RECORD
        Private HorizPageBreakRows As Short()
        Private NumHorizPageBreaks As Short
        Private strFileName As String
        Private strm As Stream
        Private writer As BinaryWriter
        Private XLSGenerateType As GenerateType

        ' Nested Types
        <StructLayout(LayoutKind.Sequential)> _
        Private Structure BEG_FILE_RECORD
            Public opcode As Short
            Public length As Short
            Public version As Short
            Public ftype As Short
        End Structure

        Public Enum CellAlignment
            ' Fields
            xlsBottomBorder = &H40
            xlsCentreAlign = 2
            xlsFillCell = 4
            xlsGeneralAlign = 0
            xlsLeftAlign = 1
            xlsLeftBorder = 8
            xlsRightAlign = 3
            xlsRightBorder = &H10
            xlsShaded = &H80
            xlsTopBorder = &H20
        End Enum

        Public Enum CellFont
            ' Fields
            xlsFont0 = 0
            xlsFont1 = &H40
            xlsFont2 = &H80
            xlsFont3 = &HC0
        End Enum

        Public Enum CellHiddenLocked
            ' Fields
            xlsHidden = &H80
            xlsLocked = &H40
            xlsNormal = 0
        End Enum

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure COLWIDTH_RECORD
            Public opcode As Short
            Public length As Short
            Public col1 As Byte
            Public col2 As Byte
            Public ColumnWidth As Short
        End Structure

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure DEF_ROWHEIGHT_RECORD
            Public opcode As Short
            Public length As Short
            Public RowHeight As Short
        End Structure

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure END_FILE_RECORD
            Public opcode As Short
            Public length As Short
        End Structure

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure FONT_RECORD
            Public opcode As Short
            Public length As Short
            Public FontHeight As Short
            Public FontAttributes1 As Byte
            Public FontAttributes2 As Byte
            Public FontNameLength As Byte
        End Structure

        Public Enum FontFormatting
            ' Fields
            xlsBold = 1
            xlsItalic = 2
            xlsNoFormat = 0
            xlsStrikeout = 8
            xlsUnderline = 4
        End Enum

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure FORMAT_COUNT_RECORD
            Public opcode As Short
            Public length As Short
            Public Count As Short
        End Structure

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure FORMAT_RECORD
            Public opcode As Short
            Public length As Short
            Public FormatLenght As Byte
        End Structure

        Public Enum GenerateType
            ' Fields
            ToFile = 1
            ToMemory = 0
        End Enum

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure HEADER_FOOTER_RECORD
            Public opcode As Short
            Public length As Short
            Public TextLength As Byte
        End Structure

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure HPAGE_BREAK_RECORD
            Public opcode As Short
            Public length As Short
            Public NumPageBreaks As Short
        End Structure

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure MARGIN_RECORD_LAYOUT
            Public opcode As Short
            Public length As Short
            Public MarginValue As Double
        End Structure

        Public Enum MarginTypes
            ' Fields
            xlsBottomMargin = &H29
            xlsLeftMargin = &H26
            xlsRightMargin = &H27
            xlsTopMargin = 40
        End Enum

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure PASSWORD_RECORD
            Public opcode As Short
            Public length As Short
        End Structure

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure PRINT_GRIDLINES_RECORD
            Public opcode As Short
            Public length As Short
            Public PrintFlag As Short
        End Structure

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure PROTECT_SPREADSHEET_RECORD
            Public opcode As Short
            Public length As Short
            Public Protect As Short
        End Structure

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure ROW_HEIGHT_RECORD
            Public opcode As Short
            Public length As Short
            Public RowNumber As Short
            Public FirstColumn As Short
            Public LastColumn As Short
            Public RowHeight As Short
            Public internal As Short
            Public DefaultAttributes As Byte
            Public FileOffset As Short
            Public rgbAttr1 As Byte
            Public rgbAttr2 As Byte
            Public rgbAttr3 As Byte
        End Structure

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure tInteger
            Public opcode As Short
            Public length As Short
            Public Row As Short
            Public col As Short
            Public rgbAttr1 As Byte
            Public rgbAttr2 As Byte
            Public rgbAttr3 As Byte
            Public intValue As Short
        End Structure

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure tNumber
            Public opcode As Short
            Public length As Short
            Public Row As Short
            Public col As Short
            Public rgbAttr1 As Byte
            Public rgbAttr2 As Byte
            Public rgbAttr3 As Byte
            Public NumberValue As Double
        End Structure

        <StructLayout(LayoutKind.Sequential)> _
        Private Structure tText
            Public opcode As Short
            Public length As Short
            Public Row As Short
            Public col As Short
            Public rgbAttr1 As Byte
            Public rgbAttr2 As Byte
            Public rgbAttr3 As Byte
            Public TextLength As Byte
        End Structure

        Public Enum ValueTypes
            ' Fields
            xlsInteger = 0
            xlsNumber = 1
            xlsText = 2
        End Enum
    End Class
End Namespace

