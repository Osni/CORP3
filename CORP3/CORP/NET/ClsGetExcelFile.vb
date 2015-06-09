Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Runtime.CompilerServices

Namespace CORP3.NET
    Public Class ClsGetExcelFile
        ' Methods
        Public Sub New()
            Dim hshDeParaDBTypes As Hashtable = Me.hshDeParaDBTypes
            hshDeParaDBTypes.Add("System.Int16", ClsExcel.ValueTypes.xlsInteger)
            hshDeParaDBTypes.Add("System.Boolean", ClsExcel.ValueTypes.xlsInteger)
            hshDeParaDBTypes.Add("System.Int32", ClsExcel.ValueTypes.xlsNumber)
            hshDeParaDBTypes.Add("System.Decimal", ClsExcel.ValueTypes.xlsNumber)
            hshDeParaDBTypes.Add("System.DateTime", ClsExcel.ValueTypes.xlsNumber)
            hshDeParaDBTypes.Add("System.Double", ClsExcel.ValueTypes.xlsNumber)
            hshDeParaDBTypes.Add("System.Single", ClsExcel.ValueTypes.xlsNumber)
            hshDeParaDBTypes.Add("System.TimeSpan", ClsExcel.ValueTypes.xlsText)
            hshDeParaDBTypes.Add("System.String", ClsExcel.ValueTypes.xlsText)
            hshDeParaDBTypes = Nothing
        End Sub

        Public Sub AddColumnTitle(ByVal strColName As String, ByVal strTitle As String)
            Me.mTitleColumn.Add(strColName.ToLower, strTitle.Replace("&nbsp;", " "))
        End Sub

        Public Sub AddGroupColumn(ByVal strColName As String)
            Me.mGroupColumn.Add(strColName.ToLower)
        End Sub

        Public Sub GenerateXLS()
            Dim hashtable As New Hashtable
            If (Me.DataSource.Rows.Count <> 0) Then
                Dim ex As New ClsExcel(Me.FileName, Me.mXLSGenerateType)
                Dim excel2 As ClsExcel = ex
                excel2.PrintGridLines = True
                excel2.SetMargin(ClsExcel.MarginTypes.xlsTopMargin, 1.5)
                excel2.SetMargin(ClsExcel.MarginTypes.xlsLeftMargin, 1.5)
                excel2.SetMargin(ClsExcel.MarginTypes.xlsRightMargin, 1.5)
                excel2.SetMargin(ClsExcel.MarginTypes.xlsBottomMargin, 1.5)
                Dim num As Integer
                For Each num In Me.RetQuebras
                    excel2.InsertHorizPageBreak(num)
                Next
                excel2.SetDefaultRowHeight(14)
                excel2.SetFont("Verdana", 11, ClsExcel.FontFormatting.xlsBold)
                excel2.SetFont("Verdana", 9, ClsExcel.FontFormatting.xlsNoFormat)
                excel2.SetFont("Verdana", 9, ClsExcel.FontFormatting.xlsBold)
                excel2.SetFont("Verdana", 13, ClsExcel.FontFormatting.xlsBold)
                excel2.SetHeader(Me.TextoCabecalho)
                excel2.SetFooter(Me.TextoRodape)
                If (Me.mGroupColumn.Count > 0) Then
                    Me.MakeGroup(ex)
                Else
                    Me.MakeSimple(ex)
                End If
                excel2.ProtectSpreadsheet = Me.ProtegerPlanilha
                excel2.SetEOF()
                Me.strm = excel2.GetStream
                excel2 = Nothing
            End If
        End Sub

        Private Sub MakeGroup(ByRef ex As ClsExcel)
            Dim enumerator As IEnumerator
            Dim num4 As Integer = 1
            Dim lRow As Integer = 1
            Dim lCol As Integer = 0
            Dim str As String = String.Empty
            Dim str2 As String = String.Empty
            Dim excel As ClsExcel = ex
            Dim dtr As DataRow = Me.mDataSource.Rows.Item(0)
            Me.SetTitleGroup(ex, dtr, lRow)
            dtr = Me.mDataSource.Rows.Item(0)
            str = Me.RetGrupo(dtr)
            Try
                enumerator = Me.mDataSource.Rows.GetEnumerator
                Do While enumerator.MoveNext
                    Dim objectValue As Object
                    Dim current As DataRow = DirectCast(enumerator.Current, DataRow)
                    lCol = 0
                    str2 = Me.RetGrupo(current)
                    If (str2 <> str) Then
                        str = str2
                        Me.SetGrupo(ex, current, lRow)
                        num4 = (num4 + 2)
                    End If
                    If (Me.mTitleColumn.Count > 0) Then
                        Dim str3 As String
                        For Each str3 In Me.mTitleColumn.Keys
                            If (Not Me.mGroupColumn.Contains(str3.ToLower) And Me.mTitleColumn.ContainsKey(str3.ToLower)) Then
                                lCol += 1
                                objectValue = RuntimeHelpers.GetObjectValue(Interaction.IIf(Information.IsDBNull(RuntimeHelpers.GetObjectValue(current.Item(str3))), "", RuntimeHelpers.GetObjectValue(current.Item(str3))))
                                Select Case Conversions.ToInteger(Me.hshDeParaDBTypes.Item(objectValue.GetType.ToString))
                                    Case 0
                                        If Operators.ConditionalCompareObjectLessEqual(objectValue, &H7FFF, False) Then
                                            excel.WriteInteger(ClsExcel.CellFont.xlsFont1, lRow, lCol, RuntimeHelpers.GetObjectValue(objectValue), 0, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsRightAlign)))))
                                        Else
                                            excel.WriteNumber(ClsExcel.CellFont.xlsFont1, lRow, lCol, RuntimeHelpers.GetObjectValue(objectValue), 0, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsRightAlign)))))
                                        End If
                                        Continue For
                                    Case 1
                                        If (objectValue.GetType.ToString = "System.DateTime") Then
                                            Dim cellFormat As Short = 20
                                            If (((Conversions.ToDate(objectValue).Hour = 0) And (Conversions.ToDate(objectValue).Minute = 0)) And (Conversions.ToDate(objectValue).Second = 0)) Then
                                                cellFormat = 12
                                            End If
                                            excel.WriteDate(ClsExcel.CellFont.xlsFont1, lRow, lCol, RuntimeHelpers.GetObjectValue(objectValue), cellFormat, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsRightAlign)))))
                                        Else
                                            excel.WriteNumber(ClsExcel.CellFont.xlsFont1, lRow, lCol, RuntimeHelpers.GetObjectValue(objectValue), 0, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsRightAlign)))))
                                        End If
                                        Continue For
                                End Select
                                excel.WriteText(ClsExcel.CellFont.xlsFont1, lRow, lCol, Me.oTool.RemTags(objectValue.ToString), 0, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsLeftAlign)))))
                            End If
                        Next
                    Else
                        Dim enumerator3 As IEnumerator
                        Try
                            enumerator3 = Me.mDataSource.Columns.GetEnumerator
                            Do While enumerator3.MoveNext
                                Dim column As DataColumn = DirectCast(enumerator3.Current, DataColumn)
                                If Not Me.mGroupColumn.Contains(column.ColumnName.ToLower) Then
                                    lCol += 1
                                    objectValue = RuntimeHelpers.GetObjectValue(Interaction.IIf(Information.IsDBNull(RuntimeHelpers.GetObjectValue(current.Item(column.ColumnName))), "", RuntimeHelpers.GetObjectValue(current.Item(column.ColumnName))))
                                    Select Case Conversions.ToInteger(Me.hshDeParaDBTypes.Item(objectValue.GetType.ToString))
                                        Case 0
                                            excel.WriteInteger(ClsExcel.CellFont.xlsFont1, lRow, lCol, RuntimeHelpers.GetObjectValue(objectValue), 0, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsRightAlign)))))
                                            Continue Do
                                        Case 1
                                            If (objectValue.GetType.ToString = "System.DateTime") Then
                                                Dim num6 As Short = 20
                                                If (((Conversions.ToDate(objectValue).Hour = 0) And (Conversions.ToDate(objectValue).Minute = 0)) And (Conversions.ToDate(objectValue).Second = 0)) Then
                                                    num6 = 12
                                                End If
                                                excel.WriteDate(ClsExcel.CellFont.xlsFont1, lRow, lCol, RuntimeHelpers.GetObjectValue(objectValue), num6, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsRightAlign)))))
                                            Else
                                                excel.WriteNumber(ClsExcel.CellFont.xlsFont1, lRow, lCol, RuntimeHelpers.GetObjectValue(objectValue), 0, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsRightAlign)))))
                                            End If
                                            Continue Do
                                    End Select
                                    excel.WriteText(ClsExcel.CellFont.xlsFont1, lRow, lCol, Me.oTool.RemTags(objectValue.ToString), 0, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsLeftAlign)))))
                                End If
                            Loop
                        Finally
                            If TypeOf enumerator3 Is IDisposable Then
                                TryCast(enumerator3, IDisposable).Dispose()
                            End If
                        End Try
                    End If
                    num4 += 1
                    lRow += 1
                    If ((num4 > Me.LinhasPorPagina) And (Me.LinhasPorPagina > 0)) Then
                        Me.SetTitleGroup(ex, current, lRow)
                        If (Me.Titulo <> "") Then
                            num4 = 3
                        Else
                            num4 = 2
                        End If
                    End If
                Loop
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try
            excel = Nothing
        End Sub

        Private Sub MakeSimple(ByRef ex As ClsExcel)
            Dim enumerator As IEnumerator
            Dim num4 As Integer = 1
            Dim lRow As Integer = 1
            Dim lCol As Integer = 0
            Dim excel As ClsExcel = ex
            Dim dtr As DataRow = Me.mDataSource.Rows.Item(0)
            Me.SetTitle(ex, dtr, lRow)
            If (Me.Titulo <> String.Empty) Then
                num4 += 1
            End If
            Try
                enumerator = Me.mDataSource.Rows.GetEnumerator
                Do While enumerator.MoveNext
                    Dim objectValue As Object
                    Dim current As DataRow = DirectCast(enumerator.Current, DataRow)
                    lCol = 0
                    If (Me.mTitleColumn.Count > 0) Then
                        Dim str As String
                        For Each str In Me.mTitleColumn.Keys
                            lCol += 1
                            objectValue = RuntimeHelpers.GetObjectValue(Interaction.IIf(Information.IsDBNull(RuntimeHelpers.GetObjectValue(current.Item(str))), "", RuntimeHelpers.GetObjectValue(current.Item(str))))
                            Select Case Conversions.ToInteger(Me.hshDeParaDBTypes.Item(objectValue.GetType.ToString))
                                Case 0
                                    excel.WriteInteger(ClsExcel.CellFont.xlsFont1, lRow, lCol, RuntimeHelpers.GetObjectValue(objectValue), 0, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsRightAlign)))))
                                    Exit Select
                                Case 1
                                    If (objectValue.GetType.ToString = "System.DateTime") Then
                                        Dim cellFormat As Short = 20
                                        If (((Conversions.ToDate(objectValue).Hour = 0) And (Conversions.ToDate(objectValue).Minute = 0)) And (Conversions.ToDate(objectValue).Second = 0)) Then
                                            cellFormat = 12
                                        End If
                                        excel.WriteDate(ClsExcel.CellFont.xlsFont1, lRow, lCol, RuntimeHelpers.GetObjectValue(objectValue), cellFormat, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsRightAlign)))))
                                    Else
                                        excel.WriteNumber(ClsExcel.CellFont.xlsFont1, lRow, lCol, RuntimeHelpers.GetObjectValue(objectValue), 0, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsRightAlign)))))
                                    End If
                                    Exit Select
                                Case Else
                                    excel.WriteText(ClsExcel.CellFont.xlsFont1, lRow, lCol, Me.oTool.RemTags(objectValue.ToString), 0, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsLeftAlign)))))
                                    Exit Select
                            End Select
                        Next
                    Else
                        Dim enumerator3 As IEnumerator
                        Try
                            enumerator3 = Me.mDataSource.Columns.GetEnumerator
                            Do While enumerator3.MoveNext
                                Dim column As DataColumn = DirectCast(enumerator3.Current, DataColumn)
                                lCol += 1
                                objectValue = RuntimeHelpers.GetObjectValue(Interaction.IIf(Information.IsDBNull(RuntimeHelpers.GetObjectValue(current.Item(column.ColumnName))), "", RuntimeHelpers.GetObjectValue(current.Item(column.ColumnName))))
                                Select Case Conversions.ToInteger(Me.hshDeParaDBTypes.Item(objectValue.GetType.ToString))
                                    Case 0
                                        excel.WriteInteger(ClsExcel.CellFont.xlsFont1, lRow, lCol, RuntimeHelpers.GetObjectValue(objectValue), 0, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsRightAlign)))))
                                        Continue Do
                                    Case 1
                                        If ((objectValue.GetType.ToString = "System.DateTime") Or (objectValue.GetType.ToString = "System.TimeSpan")) Then
                                            Dim num6 As Short = 20
                                            If (((Conversions.ToDate(objectValue).Hour = 0) And (Conversions.ToDate(objectValue).Minute = 0)) And (Conversions.ToDate(objectValue).Second = 0)) Then
                                                num6 = 12
                                            End If
                                            excel.WriteDate(ClsExcel.CellFont.xlsFont1, lRow, lCol, RuntimeHelpers.GetObjectValue(objectValue), num6, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsRightAlign)))))
                                        Else
                                            excel.WriteNumber(ClsExcel.CellFont.xlsFont1, lRow, lCol, RuntimeHelpers.GetObjectValue(objectValue), 0, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsRightAlign)))))
                                        End If
                                        Continue Do
                                End Select
                                excel.WriteText(ClsExcel.CellFont.xlsFont1, lRow, lCol, Me.oTool.RemTags(objectValue.ToString), 0, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsLeftAlign)))))
                            Loop
                        Finally
                            If TypeOf enumerator3 Is IDisposable Then
                                TryCast(enumerator3, IDisposable).Dispose()
                            End If
                        End Try
                    End If
                    num4 += 1
                    lRow += 1
                    If ((num4 = Me.mLimiteLinhas) And (Me.LinhasPorPagina > 0)) Then
                        Me.SetTitle(ex, current, lRow)
                        If (Me.Titulo <> String.Empty) Then
                            num4 = 2
                        Else
                            num4 = 1
                        End If
                    End If
                Loop
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try
            excel = Nothing
        End Sub

        Private Function RetGrupo(ByRef dtr As DataRow) As String
            Dim left As String = String.Empty
            Dim str3 As String
            For Each str3 In Me.mGroupColumn
                left = Conversions.ToString(Operators.ConcatenateObject(left, dtr.Item(str3)))
            Next
            Return left
        End Function

        Private Function RetNumGrupos() As Integer
            Dim num2 As Integer
            Dim enumerator As IEnumerator
            Dim str As String = String.Empty
            Dim str2 As String = String.Empty
            Dim num As Integer = 0
            Try
                enumerator = Me.mDataSource.Rows.GetEnumerator
                Do While enumerator.MoveNext
                    Dim current As DataRow = DirectCast(enumerator.Current, DataRow)
                    str2 = Me.RetGrupo(current)
                    If (str <> str2) Then
                        num += 1
                    End If
                Loop
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try
            Return num2
        End Function

        Private Function RetQuebras() As List(Of Integer)
            Dim list As New List(Of Integer)
            Dim str As String = String.Empty
            Dim str2 As String = String.Empty
            Dim num As Integer = 0
            Dim num2 As Integer = 0
            If (Me.LinhasPorPagina > 0) Then
                Me.mLimiteLinhas = Me.LinhasPorPagina
                If (Me.Titulo <> String.Empty) Then
                    Me.mLimiteLinhas += 1
                    num2 += 1
                    num += 1
                End If
                If (Me.mGroupColumn.Count = 0) Then
                    num2 += 1
                    num += 1
                End If
                Me.mLimiteLinhas += 1
                Dim num4 As Integer = (Me.mDataSource.Rows.Count - 1)
                Dim i As Integer = 0
                Do While (i <= num4)
                    If (Me.mGroupColumn.Count > 0) Then
                        Dim dtr As DataRow = Me.mDataSource.Rows.Item(i)
                        str2 = Me.RetGrupo(dtr)
                        If (str <> str2) Then
                            num = (num + 2)
                            num2 = (num2 + 2)
                            str = str2
                        End If
                    End If
                    num += 1
                    num2 += 1
                    If (num = Me.mLimiteLinhas) Then
                        list.Add((num2 + 1))
                        num = 0
                        str = ""
                        If (Me.Titulo <> String.Empty) Then
                            num2 += 1
                            num += 1
                        End If
                        If (Me.mGroupColumn.Count = 0) Then
                            num2 += 1
                            num += 1
                        End If
                    End If
                    i += 1
                Loop
            End If
            Return list
        End Function

        Private Sub SetGrupo(ByRef ex As ClsExcel, ByRef dtr As DataRow, ByRef lRow As Integer)
            Dim left As String = String.Empty
            Dim right As String = String.Empty
            If (Me.mGroupColumn.Count > 0) Then
                Dim str3 As String
                For Each str3 In Me.mGroupColumn
                    If Me.mTitleColumn.ContainsKey(str3) Then
                        right = Conversions.ToString(Interaction.IIf((Me.mTitleColumn.Item(str3).Trim <> String.Empty), (Me.mTitleColumn.Item(str3) & ": "), ""))
                        left = Conversions.ToString(Operators.ConcatenateObject(left, Operators.ConcatenateObject(Operators.ConcatenateObject(Interaction.IIf((left.Trim <> String.Empty), " - ", ""), right), dtr.Item(str3))))
                    Else
                        left = Conversions.ToString(Operators.ConcatenateObject(left, Operators.ConcatenateObject(Interaction.IIf((left.Trim <> String.Empty), " - ", ""), dtr.Item(str3))))
                    End If
                Next
                ex.WriteValue(ClsExcel.ValueTypes.xlsText, ClsExcel.CellFont.xlsFont0, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsLeftAlign)))), ClsExcel.CellHiddenLocked.xlsNormal, lRow, 1, left, 0)
            End If
            lRow += 1
            Me.SetTitleColumns(ex, lRow)
        End Sub

        Private Sub SetTitle(ByRef ex As ClsExcel, ByRef dtr As DataRow, ByRef lRow As Integer)
            If (Me.Titulo <> String.Empty) Then
                ex.WriteValue(ClsExcel.ValueTypes.xlsText, ClsExcel.CellFont.xlsFont3, ClsExcel.CellAlignment.xlsLeftAlign, ClsExcel.CellHiddenLocked.xlsNormal, lRow, 1, Me.Titulo, 0)
                lRow += 1
            End If
            Me.SetTitleColumns(ex, lRow)
        End Sub

        Private Sub SetTitleColumns(ByRef ex As ClsExcel, ByRef lRow As Integer)
            Dim lCol As Integer = 0
            If (Me.mGroupColumn.Count > 0) Then
                If (Me.mTitleColumn.Count > 0) Then
                    Dim str As String
                    For Each str In Me.mTitleColumn.Keys
                        If Not Me.mGroupColumn.Contains(str) Then
                            lCol += 1
                            ex.WriteValue(ClsExcel.ValueTypes.xlsText, ClsExcel.CellFont.xlsFont2, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsCentreAlign)))), ClsExcel.CellHiddenLocked.xlsNormal, lRow, lCol, Me.mTitleColumn.Item(str), 0)
                        End If
                    Next
                Else
                    Dim enumerator As IEnumerator
                    Try
                        enumerator = Me.mDataSource.Columns.GetEnumerator
                        Do While enumerator.MoveNext
                            Dim current As DataColumn = DirectCast(enumerator.Current, DataColumn)
                            If Not Me.mGroupColumn.Contains(current.ColumnName.ToLower) Then
                                lCol += 1
                                ex.WriteValue(ClsExcel.ValueTypes.xlsText, ClsExcel.CellFont.xlsFont2, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsCentreAlign)))), ClsExcel.CellHiddenLocked.xlsNormal, lRow, lCol, current.ColumnName, 0)
                            End If
                        Loop
                    Finally
                        If TypeOf enumerator Is IDisposable Then
                            TryCast(enumerator, IDisposable).Dispose()
                        End If
                    End Try
                End If
            ElseIf (Me.mTitleColumn.Count > 0) Then
                Dim str2 As String
                For Each str2 In Me.mTitleColumn.Keys
                    If Me.mTitleColumn.ContainsKey(str2.ToLower) Then
                        lCol += 1
                        ex.WriteValue(ClsExcel.ValueTypes.xlsText, ClsExcel.CellFont.xlsFont2, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsCentreAlign)))), ClsExcel.CellHiddenLocked.xlsNormal, lRow, lCol, Me.mTitleColumn.Item(str2), 0)
                    End If
                Next
            Else
                Dim enumerator4 As IEnumerator
                Try
                    enumerator4 = Me.mDataSource.Columns.GetEnumerator
                    Do While enumerator4.MoveNext
                        Dim column2 As DataColumn = DirectCast(enumerator4.Current, DataColumn)
                        lCol += 1
                        ex.WriteValue(ClsExcel.ValueTypes.xlsText, ClsExcel.CellFont.xlsFont2, (ClsExcel.CellAlignment.xlsBottomBorder Or (ClsExcel.CellAlignment.xlsTopBorder Or (ClsExcel.CellAlignment.xlsRightBorder Or (ClsExcel.CellAlignment.xlsLeftBorder Or ClsExcel.CellAlignment.xlsCentreAlign)))), ClsExcel.CellHiddenLocked.xlsNormal, lRow, lCol, column2.ColumnName, 0)
                    Loop
                Finally
                    If TypeOf enumerator4 Is IDisposable Then
                        TryCast(enumerator4, IDisposable).Dispose()
                    End If
                End Try
            End If
            lRow += 1
        End Sub

        Private Sub SetTitleGroup(ByRef ex As ClsExcel, ByRef dtr As DataRow, ByRef lRow As Integer)
            If (Me.Titulo <> String.Empty) Then
                ex.WriteValue(ClsExcel.ValueTypes.xlsText, ClsExcel.CellFont.xlsFont3, ClsExcel.CellAlignment.xlsLeftAlign, ClsExcel.CellHiddenLocked.xlsNormal, lRow, 1, Me.Titulo, 0)
                lRow += 1
            End If
            Me.SetGrupo(ex, dtr, lRow)
        End Sub


        ' Properties
        Public Property DataSource() As DataTable
            Get
                Return Me.mDataSource
            End Get
            Set(ByVal value As DataTable)
                Me.mDataSource = value
            End Set
        End Property

        Public Property FileName() As String
            Get
                If (Me.mFileName.Trim = String.Empty) Then
                    Me.mFileName = ("ARQ_" & Guid.NewGuid.ToString & ".xls")
                End If
                Return Me.mFileName
            End Get
            Set(ByVal value As String)
                Me.mFileName = value
            End Set
        End Property

        Public ReadOnly Property GetStream() As Stream
            Get
                Return Me.strm
            End Get
        End Property

        Public Property LinhasPorPagina() As Short
            Get
                Return Me.mLinhasPorPagina
            End Get
            Set(ByVal value As Short)
                Me.mLinhasPorPagina = value
            End Set
        End Property

        Public Property ProtegerPlanilha() As Boolean
            Get
                Return Me.mProtegerPlanilha
            End Get
            Set(ByVal value As Boolean)
                Me.mProtegerPlanilha = value
            End Set
        End Property

        Public Property TextoCabecalho() As String
            Get
                Return Me.mTextoCabecalho
            End Get
            Set(ByVal value As String)
                Me.mTextoCabecalho = value
            End Set
        End Property

        Public Property TextoRodape() As String
            Get
                Return Me.mTextoRodape.Trim
            End Get
            Set(ByVal value As String)
                Me.mTextoCabecalho = value.Trim
            End Set
        End Property

        Public Property Titulo() As String
            Get
                Return Me.mTitulo.Trim
            End Get
            Set(ByVal value As String)
                Me.mTitulo = value
            End Set
        End Property

        Public Property XLSGenerateType() As ClsExcel.GenerateType
            Get
                Return Me.mXLSGenerateType
            End Get
            Set(ByVal value As ClsExcel.GenerateType)
                Me.mXLSGenerateType = value
            End Set
        End Property


        ' Fields
        Public Const FORMATNORMALCELNUMBER As Byte = &H7B
        Public Const FORMATNORMALCELTEXT As Byte = &H79
        Public Const FORMATROTULOCEL As Byte = &H7A
        Private hshDeParaDBTypes As Hashtable = New Hashtable
        Private mDataSource As DataTable = New DataTable
        Private mFileName As String = String.Empty
        Private mGroupColumn As List(Of String) = New List(Of String)
        Private mLimiteLinhas As Integer = 0
        Private mLinhasPorPagina As Short = 0
        Private mProtegerPlanilha As Boolean = False
        Private mTextoCabecalho As String = String.Empty
        Private mTextoRodape As String = String.Empty
        Private mTitleColumn As Dictionary(Of String, String) = New Dictionary(Of String, String)
        Private mTitulo As String = String.Empty
        Private mXLSGenerateType As ClsExcel.GenerateType = ClsExcel.GenerateType.ToMemory
        Private oTool As ClsTools = New ClsTools
        Private strm As Stream
    End Class
End Namespace

