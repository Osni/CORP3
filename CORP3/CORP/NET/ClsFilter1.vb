Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Collections
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Data
Imports System.Data.Common
Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.HtmlControls
Imports System.Web.UI.WebControls

Namespace CORP3.NET
    <Serializable(), DefaultProperty("Text"), ToolboxData("<{0}:ClsFilter1 runat=server></{0}:ClsFilter1>")> _
    Public Class ClsFilter1
        Inherits WebControl
        Implements IPostBackDataHandler, IPostBackEventHandler
        ' Events
        'Public Custom Event FILTER_CellClick As FILTER_CellClickEventHandler

        'End Event
        ' Methods
        Public Sub New()
            AddHandler MyBase.PreRender, New EventHandler(AddressOf Me.ClsFilter_PreRender)
            Me._GetSQLOrderBy = ""
            Me._FILTER_VIEWSTATE = ""
            Me.x = 0
            Me._FilterQueryString = New DataTable
            Dim columns As DataColumnCollection = Me._FilterQueryString.Columns
            columns.Add("key")
            columns.Add("value")
            columns = Nothing
            Me._FilterCols = New Collection
            Me._ClsFilterCols = New ClsFilterCols("", "", "", True, 10, ClsFilterCols.TypeDB.STRING_T, "", "", "", ClsFilterCols.TStyle.FilterField)
            Me._FilterState = New PropertyCollection
        End Sub

        Public Function AddCol(ByVal pClsFCols As PropertyCollection) As Boolean
            Me._FilterCols.Add(pClsFCols, Conversions.ToString(pClsFCols.Item("Name")), Nothing, Nothing)
            Return True
        End Function

        Public Function AddCol(ByVal sName As String, ByVal sLabel As String, Optional ByVal sTitle As String = "", Optional ByVal bVisible As Boolean = True, Optional ByVal iSize As Integer = 10, Optional ByVal eTypeDB As ClsFilterCols.TypeDB = 1, Optional ByVal sTextValue As String = "", Optional ByVal sPageURLDestino As String = "", Optional ByVal sPageURLColVar As String = "", Optional ByVal Style As ClsFilterCols.TStyle = 0) As Boolean
            Me._FilterCols.Add(New ClsFilterCols(sName, sLabel, sTitle, bVisible, iSize, eTypeDB, sTextValue, sPageURLDestino, sPageURLColVar, Style).FilterColsReadOnly, sName, Nothing, Nothing)
            Return True
        End Function

        Public Sub AddQueryString(ByVal sName As String, Optional ByVal sValue As String = "")
            If (sName.Trim = "") Then
                Throw New Exception("Informe Nome da variável!")
            End If
            If sName.Contains(" ") Then
                Throw New Exception("Nome da Chave inválida!")
            End If
            Dim table As DataTable = Me._FilterQueryString
            Dim row As DataRow = table.NewRow
            row.Item("key") = sName.Trim
            row.Item("value") = sValue.Replace("  ", " ").Trim
            table.Rows.Add(row)
            table = Nothing
        End Sub

        Private Function CheckFieldExists(ByVal strName As String, ByVal Row As DataRowView) As Boolean
            Dim flag As Boolean
            Try
                flag = True
                Dim str As String = Conversions.ToString(Row.Item(strName))
            Catch exception1 As Exception
                ProjectData.SetProjectError(exception1)
                Dim exception As Exception = exception1
                flag = False
                ProjectData.ClearProjectError()
            End Try
            Return flag
        End Function

        Private Sub ClsFilter_PreRender(ByVal sender As Object, ByVal e As EventArgs)
            Dim currentHandler As Page = DirectCast(Me.Context.CurrentHandler, Page)
            Dim styleSheet As IStyleSheet = currentHandler.Header.StyleSheet
            Me.Page.Form.Attributes.Add("autocomplete", "off")
        End Sub

        Protected Overrides Sub CreateChildControls()
            Dim str As String = ""
            If (Me.FilterType = PFilterType.SimpleFilter) Then
                str = (Me.ClientID & "_FILTER_LIST")
                Dim child As New LiteralControl(String.Concat(New String() {"<input style=""width:90%"" type=""text"" name=""", str, "_TEXT"" id=""", str, "_TEXT"" value=""", Me._FilterColSimpleText, """ />&nbsp;&nbsp;&nbsp;"}))
                Me.Controls.Add(child)
                child = New LiteralControl(("<a href=#><img onclick=""javascript:" & Me.Page.ClientScript.GetPostBackEventReference(Me, "filter_list_click") & """ title='List Filtro'  id='filter_list_click_img' name='filter_list_click_img' src='imagens/list_filter.gif' border=0></a>"))
                Me.Controls.Add(child)
                child = New LiteralControl(String.Concat(New String() {"<input type=""hidden"" name=""", str, "_VALUE"" id=""", str, "_VALUE"" value=""", Me._FilterColSimpleValue, """  />"}))
                Me.Controls.Add(child)
            End If
        End Sub

        Public Sub ExportarXLS()
            Dim dataTable As New DataTable
            Dim dataAdapter As DbDataAdapter = New ClsDB1(Me.FilterStrConnection, Me.FilterProvider).GetDataAdapter
            dataAdapter.SelectCommand.CommandText = (Me.GetSQLSelect & " " & Me._GetSQLWhere)
            dataAdapter.Fill(dataTable)
            dataAdapter = Nothing
            Dim file As New ClsGetExcelFile
            Dim response As HttpResponse = HttpContext.Current.Response
            If (dataTable.Rows.Count = 0) Then
                Me.Page.Controls.Clear()
                response.Clear()
                Me.Controls.Add(New LiteralControl("<div id=""rptMsg"" style=""color: red;"">Nenhuma informação foi gerada.</div>"))
            Else
                Dim enumerator As IEnumerator
                Dim file2 As ClsGetExcelFile = file
                file2.DataSource = dataTable
                Try
                    enumerator = Me._FilterCols.GetEnumerator
                    Do While enumerator.MoveNext
                        Dim current As PropertyCollection = DirectCast(enumerator.Current, PropertyCollection)
                        If (Conversions.ToInteger(current.Item("Style")) = 0) Then
                            file2.AddColumnTitle(Conversions.ToString(current.Item("Name")), Conversions.ToString(current.Item("Label")))
                        End If
                    Loop
                Finally
                    If TypeOf enumerator Is IDisposable Then
                        TryCast(enumerator, IDisposable).Dispose()
                    End If
                End Try
                file2.GenerateXLS()
                file2 = Nothing
                Dim buffer As Byte() = DirectCast(file.GetStream, MemoryStream).ToArray
                Dim response2 As HttpResponse = response
                response2.Clear()
                response2.AddHeader("Content-Disposition", ("attachment; filename=" & file.FileName))
                response2.AddHeader("Content-Length", buffer.Length.ToString)
                response2.ContentType = "application/vnd.ms-excel"
                response2.BinaryWrite(buffer)
                response2.End()
                response2 = Nothing
            End If
        End Sub

        Private Sub GetClearSQLSelect()
            Dim count As Integer = Me._FilterCols.Count
            Me.i = 1
            Do While (Me.i <= count)
                NewLateBinding.LateIndexSetComplex(Me._FilterCols.Item(Me.i), New Object() {"TextValue", ""}, Nothing, False, True)
                Me.i += 1
            Loop
            Me._FilterOptionClear = True
        End Sub

        Private Function GetFilterData(ByVal sSQL As String) As DataView
            If (Me.FilterStrConnection = "") Then
                Throw New Exception("Uma String de Conexão é obrigatório em <FilterStrConnection>")
            End If
            Me._PageLast = Me.GetPositionRowPage
            Dim sdb As New ClsDB1(Me.FilterStrConnection, Me.FilterProvider)
            Me._ds = New DataSet
            Dim dataAdapter As DbDataAdapter = sdb.GetDataAdapter
            dataAdapter.SelectCommand = sdb.GetCommand
            dataAdapter.SelectCommand.CommandText = sSQL
            dataAdapter.SelectCommand.Connection = sdb.GetConnection
            dataAdapter.Fill(Me._ds, CInt(Me.FilterPageRowPosition), Me.FilterPageSize, Me.FilterTableName)
            dataAdapter = Nothing
            Return Me._ds.Tables.Item(Me.FilterTableName).DefaultView
        End Function

        Private Sub GetFilterSimpleText()
            Dim request As HttpRequest = HttpContext.Current.Request
            Me._FilterSimpleText = request.Item((Me.ClientID & "_FILTER_LIST_TEXT"))
            Me.FilterParamWhere = (Me.FilterColText & " LIKE '%" & Me._FilterSimpleText.Replace("'", "''") & "%'")
        End Sub

        Private Function GetPositionRowPage() As Long
            If ((Me._TotalRows = 0) And Me.FilterPagePreLoad) Then
                Dim sdb As New ClsDB1(Me.FilterStrConnection, Me.FilterProvider)
                Dim falsePart As String = Conversions.ToString(0)
                Dim sSQL As String = ("SELECT Count(*) as total FROM " & Me.FilterTableName & " " & Me._GetSQLWhere)
                Dim pCon As DbConnection = Nothing
                Dim pTra As DbTransaction = Nothing
                falsePart = sdb.GetDataTable(sSQL, pCon, pTra, 0).Rows.Item(0).Item(0).ToString
                Me._TotalRows = Conversions.ToLong(Interaction.IIf((falsePart = String.Empty), 0, falsePart))
                Me._PageLast = Me._TotalRows
                If ((Me._TotalRows Mod CLng(Me.FilterPageSize)) = 0) Then
                    Me._PageQtde = CLng(Math.Round(Conversion.Int(CDbl((CDbl(Me._TotalRows) / CDbl(Me.FilterPageSize))))))
                Else
                    Me._PageQtde = CLng(Math.Round(CDbl((Conversion.Int(CDbl((CDbl(Me._TotalRows) / CDbl(Me.FilterPageSize)))) + 1))))
                End If
                Me.FilterPageRowPosition = 0
                Me._PageAtiva = 1
            End If
            Return Me._PageQtde
        End Function

        Private Function GetQueryString(Optional ByVal rowView As DataRowView = Nothing) As String
            Dim builder As New StringBuilder
            Dim table As DataTable = Me._FilterQueryString
            Dim num2 As Integer = (table.Rows.Count - 1)
            Dim i As Integer = 0
            Do While (i <= num2)
                Dim row As DataRow = table.Rows.Item(i)
                Dim builder2 As StringBuilder = builder
                If Me.CheckFieldExists(Conversions.ToString(row.Item("key")), rowView) Then
                    builder2.Append("&").Append(RuntimeHelpers.GetObjectValue(row.Item("key"))).Append("=").Append(NewLateBinding.LateGet(rowView, Nothing, "Item", New Object() {RuntimeHelpers.GetObjectValue(row.Item("key"))}, Nothing, Nothing, Nothing).ToString.Replace(" ", "+"))
                Else
                    builder2.Append("&").Append(RuntimeHelpers.GetObjectValue(row.Item("key"))).Append("=").Append(row.Item("value").ToString.Replace(" ", "+"))
                End If
                builder2 = Nothing
                i += 1
            Loop
            table = Nothing
            Return builder.ToString
        End Function

        Private Sub GetSQLOrderBy(ByVal sCol As String, ByVal sPos As String, ByVal sLastCol As String)
            Me._FilterOrderByColName = Me.GetViewStateFilter("ColOrderBy")
            If (sCol <> Me._FilterOrderByColName) Then
                Me._FilterLastOrderByColName = sCol
                Me._FilterColOrderByPos = "ASC"
            ElseIf (sPos = "ASC") Then
                Me._FilterColOrderByPos = "DESC"
            Else
                Me._FilterColOrderByPos = "ASC"
            End If
            Me.SetViewStateFilter("ColOrderBy", sCol)
            Me._FilterOrderByColName = sCol
            Me._GetSQLOrderBy = (" ORDER BY " & Me._FilterOrderByColName & " " & Me._FilterColOrderByPos)
        End Sub

        Private Function GetSQLSelect() As String
            Dim ssql As New ClsSQL1
            Dim ssql2 As ClsSQL1 = ssql
            ssql2.sTable = Me.FilterTableName
            Dim count As Integer = Me._FilterCols.Count
            Me.i = 1
            Do While (Me.i <= count)
                ssql2.AddCol(NewLateBinding.LateIndexGet(Me._FilterCols.Item(Me.i), New Object() {"Name"}, Nothing).ToString, "", ClsSQL1.TypeSQL.EMPTY_T)
                Me.i += 1
            Loop
            ssql2 = Nothing
            Return ssql.GetSELECT("")
        End Function

        Private Function GetSQLWhere() As String
            Dim str4 As String = ""
            Dim request As HttpRequest = HttpContext.Current.Request
            Dim num2 As Integer = (request.Form.Count - 1)
            Me.i = 0
            Do While (Me.i <= num2)
                If Not Information.IsNothing(request.Form.Keys.Item(Me.i)) Then
                    Dim str As String = request.Form.Keys.Item(Me.i).ToString
                    If (Strings.Mid(str, 1, 10).ToLower = "filter_txt") Then
                        Dim sItem As String = request.Form.Item(str).ToString
                        str = str.Substring(11)
                        Dim num As Integer = Conversions.ToInteger(NewLateBinding.LateIndexGet(Me._FilterCols.Item(str), New Object() {"Type"}, Nothing))
                        NewLateBinding.LateIndexSetComplex(Me._FilterCols.Item(str), New Object() {"TextValue", sItem}, Nothing, False, True)
                        If (sItem.Trim <> "") Then
                            str4 = String.Concat(New String() {str4, str, " ", Me.TrataSinal(sItem, DirectCast(num, ClsFilterCols.TypeDB)), " AND "})
                        End If
                    End If
                End If
                Me.i += 1
            Loop
            request = Nothing
            If (str4 = "") Then
                If (Me.FilterParamWhere <> "") Then
                    str4 = (" WHERE " & Me.FilterParamWhere)
                End If
                Return str4
            End If
            If (Me.FilterParamWhere <> "") Then
                Return String.Concat(New String() {" WHERE ", str4.Substring(0, (str4.Length - 4)), " AND (", Me.FilterParamWhere, ")"})
            End If
            Return (" WHERE " & str4.Substring(0, (str4.Length - 4)))
        End Function

        Private Function GetViewStateFilter(ByVal sKey As String) As String
            Return Conversions.ToString(Me._FilterState.Item(sKey))
        End Function

        Protected Overrides Sub LoadControlState(ByVal savedState As Object)
            Me._FilterViewState = New Collection
            Me._FilterViewState = DirectCast(savedState, Collection)
            Me._FilterCols = DirectCast(Me._FilterViewState.Item("FilterCols"), Collection)
            Me._FilterState = DirectCast(Me._FilterViewState.Item("FilterState"), PropertyCollection)
            Me._FilterQueryString = DirectCast(Me._FilterViewState.Item("FilterQueryString"), DataTable)
        End Sub

        Public Function LoadPostData(ByVal PostDataKey As String, ByVal Values As NameValueCollection) As Boolean Implements IPostBackDataHandler.LoadPostData
            Return False
        End Function

        Private Function MyPath() As String
            Dim str3 As String = HttpContext.Current.Request.Url.AbsoluteUri.Replace(HttpContext.Current.Request.Url.Query, "")
            Dim oldValue As String = HttpContext.Current.Request.AppRelativeCurrentExecutionFilePath.Replace("~/", "")
            Return (str3.Replace(oldValue, "") & "js/")
        End Function

        Public Function NullDB(ByRef pExpress As Object, Optional ByVal pReturn As Object = "") As Object
            Return Interaction.IIf(Information.IsDBNull(RuntimeHelpers.GetObjectValue(pExpress)), RuntimeHelpers.GetObjectValue(pReturn), RuntimeHelpers.GetObjectValue(pExpress))
        End Function

        Protected Overrides Sub OnInit(ByVal e As EventArgs)
            Me.SetViewStateFilter("ColOrderBy", "")
            Dim page As Page = Me.Page
            Dim clientScript As ClientScriptManager = page.ClientScript
            clientScript.RegisterClientScriptInclude(Me.Page.GetType, "js_forms", (Me.FilterJsUrl & "forms.js"))
            clientScript.RegisterClientScriptInclude(Me.Page.GetType, "js_filter", (Me.FilterJsUrl & "filter.js"))
            clientScript = Nothing
            page.RegisterRequiresControlState(Me)
            page = Nothing
            MyBase.OnInit(e)
        End Sub

        Public Sub RaisePostBackEvent(ByVal EventArgument As String) Implements IPostBackEventHandler.RaisePostBackEvent
            Dim instance As Object = Strings.Split(EventArgument, "$", -1, CompareMethod.Binary)
            Me._GetSQLWhere = Me.GetSQLWhere
            If (NewLateBinding.LateIndexGet(instance, New Object() {0}, Nothing).ToString.Substring(0, Strings.Len("filter_move")) = "filter_move") Then
                Me._PageQtde = Conversions.ToLong(NewLateBinding.LateIndexGet(instance, New Object() {2}, Nothing))
                Me._TotalRows = Conversions.ToLong(NewLateBinding.LateIndexGet(instance, New Object() {3}, Nothing))
                Me._PageLast = Me._PageQtde
            End If
            Dim left As Object = NewLateBinding.LateIndexGet(instance, New Object() {0}, Nothing)
            If Operators.ConditionalCompareObjectEqual(left, "filter_move_first", False) Then
                Me._PageLast = Me._PageQtde
                Me._PagePrevious = 0
                Me._PageNext = 0
                Me.FilterPageRowPosition = 0
                Me._PageAtiva = 1
            ElseIf Operators.ConditionalCompareObjectEqual(left, "filter_move_next", False) Then
                Me._PageNext = Conversions.ToLong(NewLateBinding.LateIndexGet(instance, New Object() {1}, Nothing))
                Me._PageNext = (Me._PageNext + 1)
                If (Me._PageNext >= Me._PageQtde) Then
                    Me._PageNext = (Me._PageNext - 1)
                End If
                Me.FilterPageRowPosition = (Me._PageNext * Me.FilterPageSize)
                Me._PagePrevious = Me._PageNext
                Me._PageAtiva = (Me._PageNext + 1)
            ElseIf Operators.ConditionalCompareObjectEqual(left, "filter_move_previous", False) Then
                Me._PagePrevious = Conversions.ToLong(NewLateBinding.LateIndexGet(instance, New Object() {1}, Nothing))
                Me._PagePrevious = (Me._PagePrevious - 1)
                If (Me._PagePrevious < 0) Then
                    Me._PagePrevious = (Me._PagePrevious + 1)
                End If
                Me.FilterPageRowPosition = (Me._PagePrevious * Me.FilterPageSize)
                Me._PageNext = Me._PagePrevious
                Me._PageAtiva = (Me._PageNext + 1)
            ElseIf Operators.ConditionalCompareObjectEqual(left, "filter_move_last", False) Then
                Me._PageNext = (Me._PageLast - 1)
                Me._PagePrevious = (Me._PageLast - 1)
                Me.FilterPageRowPosition = CLng(Math.Round(CDbl((Conversion.Int(CDbl((CDbl(Me._TotalRows) / CDbl(Me.FilterPageSize)))) * Me.FilterPageSize))))
                If (Me.FilterPageRowPosition >= Me._TotalRows) Then
                    Me.FilterPageRowPosition = CLng(Math.Round(CDbl(((Conversion.Int(CDbl((CDbl(Me._TotalRows) / CDbl(Me.FilterPageSize)))) - 1) * Me.FilterPageSize))))
                End If
                Me._PageAtiva = Me._PageLast
            ElseIf Operators.ConditionalCompareObjectEqual(left, "filter_click_aplica", False) Then
                Me._TotalRows = 0
                Me._PageNext = 0
                Me._PagePrevious = 0
                Me._PageQtde = 0
                Me._PageAtiva = 0
                Me.FilterPagePreLoad = True
            ElseIf Operators.ConditionalCompareObjectEqual(left, "filter_click_novo", False) Then
                Me.GetClearSQLSelect()
                Me._GetSQLWhere = ""
                Me.FilterPagePreLoad = False
            ElseIf (Not Operators.ConditionalCompareObjectEqual(left, "filter_click_retornar", False) AndAlso Not Operators.ConditionalCompareObjectEqual(left, "filter_click_print", False)) Then
                If Operators.ConditionalCompareObjectEqual(left, "filter_click_excel", False) Then
                    Me.ExportarXLS()
                ElseIf Operators.ConditionalCompareObjectEqual(left, "filter_orderby", False) Then
                    Me.GetSQLOrderBy(Conversions.ToString(NewLateBinding.LateIndexGet(instance, New Object() {1}, Nothing)), Conversions.ToString(NewLateBinding.LateIndexGet(instance, New Object() {2}, Nothing)), Conversions.ToString(NewLateBinding.LateIndexGet(instance, New Object() {3}, Nothing)))
                ElseIf Operators.ConditionalCompareObjectEqual(left, "filter_cellclick", False) Then
                    Me._FilterColSimpleText = Conversions.ToString(NewLateBinding.LateIndexGet(instance, New Object() {1}, Nothing))
                    Me._FilterColSimpleValue = Conversions.ToString(NewLateBinding.LateIndexGet(instance, New Object() {2}, Nothing))
                    Dim handler As FILTER_CellClickEventHandler = Nothing 'Me.FILTER_CellClickEvent
                    If (Not handler Is Nothing) Then
                        handler.Invoke(Conversions.ToString(NewLateBinding.LateIndexGet(instance, New Object() {1}, Nothing)), Conversions.ToString(NewLateBinding.LateIndexGet(instance, New Object() {2}, Nothing)))
                    End If
                    Me._FilterNotVisible = True
                ElseIf Operators.ConditionalCompareObjectEqual(left, "filter_list_click", False) Then
                    Me.GetFilterSimpleText()
                End If
            End If
        End Sub

        Public Sub RaisePostDataChangedEvent() Implements IPostBackDataHandler.RaisePostDataChangedEvent
        End Sub

        Protected Overrides Sub RenderContents(ByVal output As HtmlTextWriter)
            If (Me.FilterPagePreLoad Or (Me.FilterType = PFilterType.FullFilter)) Then
                Me.ShowFilter()
            End If
        End Sub

        Protected Overrides Function SaveControlState() As Object
            Me._FilterViewState = New Collection
            Me._FilterViewState.Add(Me._FilterCols, "FilterCols", Nothing, Nothing)
            Me._FilterViewState.Add(Me._FilterState, "FilterState", Nothing, Nothing)
            Me._FilterViewState.Add(Me._FilterQueryString, "FilterQueryString", Nothing, Nothing)
            Return Me._FilterViewState
        End Function

        Private Function SetViewStateFilter(ByVal sKey As String, ByVal sValue As String) As Boolean
            Dim flag As Boolean
            Dim objectValue As Object = Nothing
            If (Me._FilterState.Count > 0) Then
                objectValue = RuntimeHelpers.GetObjectValue(Me._FilterState.Item(sKey))
            End If
            If (Not objectValue Is Nothing) Then
                Me._FilterState.Item(sKey) = sValue
                Return flag
            End If
            Me._FilterState.Add(sKey, sValue)
            Return flag
        End Function

        Private Function ShowFilter() As Object
            Dim child As New Table
            Dim control As New HtmlGenericControl
            Dim str3 As String = Strings.Replace(Strings.Space(2), " ", "&nbsp;", 1, -1, CompareMethod.Binary)
            Dim writer2 As New StringWriter
            Dim writer As New HtmlTextWriter(writer2)
            Dim str As String = String.Empty
            If ((Me._FilterCols.Count > 0) And Not Me._FilterNotVisible) Then
                Dim current As DataRowView
                If Not Me._FilterOptionClear Then
                    If (Me.FilterType = PFilterType.SimpleFilter) Then
                        Me._GetSQLWhere = Me.GetSQLWhere
                    End If
                    If ((Me._GetSQLWhere = "") AndAlso (Me.FilterParamWhere <> "")) Then
                        Me._GetSQLWhere = Me.GetSQLWhere
                    End If
                End If
                Dim filterData As DataView = Me.GetFilterData(Conversions.ToString(Operators.ConcatenateObject((Me.GetSQLSelect & Me._GetSQLWhere), Interaction.IIf((Me._GetSQLOrderBy <> String.Empty), Me._GetSQLOrderBy, RuntimeHelpers.GetObjectValue(Interaction.IIf((Me.FilterInitialOrderBy <> ""), (" ORDER BY " & Me.FilterInitialOrderBy), ""))))))
                Dim count As Integer = Me._FilterCols.Count
                Dim table2 As Table = child
                table2.ID = "filter_table"
                table2.CssClass = "tblFilter"
                table2.CellPadding = 0
                table2.CellSpacing = 0
                table2 = Nothing
                Dim row As New TableRow
                Dim cell As New TableCell
                Dim builder As New StringBuilder
                Dim builder2 As StringBuilder = builder
                builder2.AppendLine("<div class=""divToolbar"">")
                builder2.AppendLine("<div class=""tButtons"">")
                builder2.AppendLine("<ul>")
                builder2.AppendLine("<li><img src=""img/pagina1/pixToolbar.jpg"" alt=""Movimentar"" /></li>")
                builder2.AppendLine(("<li onclick=""javascript:" & Me.Page.ClientScript.GetPostBackEventReference(Me, "filter_click_novo") & """ id='filter_button_novo' name='filter_button_novo' ><img src=""img/pagina1/btn/btnNovo.gif"" alt=""Novo"" /></li>"))
                builder2.AppendLine("<li><img src=""img/pagina1/sepToolbar.jpg"" alt=""Movimentar"" /></li>")
                builder2.AppendLine(("<li onclick=""javascript:" & Me.Page.ClientScript.GetPostBackEventReference(Me, "filter_click_excel") & """ id='filter_button_excel' name='filter_button_excel' ><img src=""img/pagina1/btn/btnExcel.gif"" alt=""Excel"" /></li>"))
                builder2.AppendLine("<li><img src=""img/pagina1/sepToolbar.jpg"" alt=""Movimentar"" /></li>")
                builder2.AppendLine(("<li onclick=""javascript:" & Me.Page.ClientScript.GetPostBackEventReference(Me, "filter_click_aplica") & """ id='filter_button_aplicar' name='filter_button_aplicar'><img src=""img/pagina1/btn/btnFiltro.gif"" alt=""filtro""  /></li>"))
                builder2.AppendLine("<li><img src=""img/pagina1/sepToolbar.jpg"" alt=""Movimentar"" /></li>")
                builder2.AppendLine("<li onclick=""javascript:window.print()"" title='Imprimir Filtro' id='filter_button_print' name='filter_button_print'><img src=""img/pagina1/btn/btnImprimir.gif"" alt=""Imprimir"" /></li>")
                builder2.AppendLine("<li><img src=""img/pagina1/sepToolbar.jpg"" alt=""Movimentar"" /></li>")
                builder2.AppendLine(("<li onclick=""javascript:" & Me.Page.ClientScript.GetPostBackEventReference(Me, ("filter_move_first$0$" & Conversions.ToString(Me._PageQtde) & "$" & Conversions.ToString(Me._TotalRows))) & """ id='filter_button_first' name='filter_button_first' ><img src=""img/pagina1/btn/btnPrimeiro.gif"" alt=""Primeiro"" /></li>"))
                builder2.AppendLine("<li><img src=""img/pagina1/sepToolbar.jpg"" alt=""Movimentar"" /></li>")
                builder2.AppendLine(("<li onclick=""javascript:" & Me.Page.ClientScript.GetPostBackEventReference(Me, String.Concat(New String() {"filter_move_previous$", Conversions.ToString(Me._PagePrevious), "$", Conversions.ToString(Me._PageQtde), "$", Conversions.ToString(Me._TotalRows)})) & """ id='filter_button_previous' name='filter_button_previous'><img src=""img/pagina1/btn/btnAnterior.gif"" alt=""Anterior"" /></li>"))
                builder2.AppendLine("<li><img src=""img/pagina1/sepToolbar.jpg"" alt=""Movimentar"" /></li>")
                builder2.AppendLine(("<li onclick=""javascript:" & Me.Page.ClientScript.GetPostBackEventReference(Me, String.Concat(New String() {"filter_move_next$", Conversions.ToString(Me._PageNext), "$", Conversions.ToString(Me._PageQtde), "$", Conversions.ToString(Me._TotalRows)})) & """ id='filter_button_next' name='filter_button_next'><img src=""img/pagina1/btn/btnProximo.gif"" alt=""Próximo"" /></li>"))
                builder2.AppendLine("<li><img src=""img/pagina1/sepToolbar.jpg"" alt=""Movimentar"" /></li>")
                builder2.AppendLine(("<li onclick=""javascript:" & Me.Page.ClientScript.GetPostBackEventReference(Me, String.Concat(New String() {"filter_move_last$", Conversions.ToString(Me._PageLast), "$", Conversions.ToString(Me._PageQtde), "$", Conversions.ToString(Me._TotalRows)})) & """ id='filter_button_last' name='filter_button_last'><img src=""img/pagina1/btn/btnUltimo.gif"" alt=""Ultimo""  /></li>"))
                builder2.AppendLine("<li><img src=""img/pagina1/sepToolbar.jpg"" alt=""Movimentar"" /></li>")
                builder2.AppendLine(("<li onclick=""javascript:window.location.href='" & Me.FilterReturnFormName & "'"" ><img src=""img/pagina1/btn/volta.gif"" alt=""Voltar""  /></li>"))
                builder2.AppendLine("<li><img src=""img/pagina1/sepToolbar.jpg"" alt=""Movimentar"" /></li>")
                builder2.AppendLine("<li>")
                builder2.AppendLine(String.Concat(New String() {"Pagina ", Conversions.ToString(Me._PageAtiva), " de ", Conversions.ToString(Me._PageQtde), str3, " &nbsp; - &nbsp;"}))
                builder2.AppendLine(("Qtde Registros (" & Conversions.ToString(Me._TotalRows) & ")"))
                builder2.AppendLine("</li>")
                builder2.AppendLine("</ul>")
                builder2.AppendLine("</div>")
                builder2.AppendLine("</div>")
                builder2 = Nothing
                row = New TableRow
                Dim num4 As Integer = count
                Me.i = 1
                Do While (Me.i <= num4)
                    cell = New TableHeaderCell With { _
                        .ID = ("filter_col_" & NewLateBinding.LateIndexGet(Me._FilterCols.Item(Me.i), New Object() {"Name"}, Nothing).ToString) _
                    }
                    If Conversions.ToBoolean(NewLateBinding.LateIndexGet(Me._FilterCols.Item(Me.i), New Object() {"Visible"}, Nothing)) Then
                        If ((Me.FilterOrderByCols And Me.FilterPagePreLoad) And (Conversions.ToInteger(NewLateBinding.LateIndexGet(Me._FilterCols.Item(Me.i), New Object() {"Style"}, Nothing)) = 0)) Then
                            Dim clientScript As ClientScriptManager = Me.Page.ClientScript
                            cell.Text = String.Concat(New String() {"<a href=""javascript:", clientScript.GetPostBackEventReference(Me, String.Concat(New String() {"filter_orderby$", NewLateBinding.LateIndexGet(Me._FilterCols.Item(Me.i), New Object() {"Name"}, Nothing).ToString, "$", Me._FilterColOrderByPos, "$", Me._FilterLastOrderByColName})), """ title='Ordernar'>", NewLateBinding.LateIndexGet(Me._FilterCols.Item(Me.i), New Object() {"Label"}, Nothing).ToString, "</a>" & ChrW(13) & ChrW(10)})
                            clientScript = Nothing
                        Else
                            cell.Text = NewLateBinding.LateIndexGet(Me._FilterCols.Item(Me.i), New Object() {"Label"}, Nothing).ToString
                        End If
                        cell.ToolTip = NewLateBinding.LateIndexGet(Me._FilterCols.Item(Me.i), New Object() {"Title"}, Nothing).ToString
                        Dim row3 As TableRow = row
                        row3.Cells.Add(cell)
                        row3.CssClass = "trHeader"
                        row3 = Nothing
                        child.Rows.Add(row)
                    End If
                    Me.i += 1
                Loop
                If (Me.FilterType = PFilterType.FullFilter) Then
                    row = New TableRow
                    Dim row2 As New TableRow
                    Dim num5 As Integer = count
                    Me.i = 1
                    Do While (Me.i <= num5)
                        cell = New TableCell
                        If Conversions.ToBoolean(NewLateBinding.LateIndexGet(Me._FilterCols.Item(Me.i), New Object() {"Visible"}, Nothing)) Then
                            Dim str6 As String = NewLateBinding.LateIndexGet(Me._FilterCols.Item(Me.i), New Object() {"Name"}, Nothing).ToString
                            Dim str7 As String = NewLateBinding.LateIndexGet(Me._FilterCols.Item(Me.i), New Object() {"Size"}, Nothing).ToString
                            Dim str8 As String = NewLateBinding.LateIndexGet(Me._FilterCols.Item(Me.i), New Object() {"Title"}, Nothing).ToString
                            If (Conversions.ToInteger(NewLateBinding.LateIndexGet(Me._FilterCols.Item(Me.i), New Object() {"Style"}, Nothing)) = 0) Then
                                str = (str & "'filter_txt_" & str6 & "', ")
                                cell.Text = String.Concat(New String() {"<input title='", str8, "' type=text id='filter_txt_", str6, "' name='filter_txt_", str6, "' size='", str7, "' value='", NewLateBinding.LateIndexGet(Me._FilterCols.Item(Me.i), New Object() {"TextValue"}, Nothing).ToString, "'"">" & ChrW(13) & ChrW(10)})
                            Else
                                Dim cell2 As TableCell = cell
                                cell2.Style.Add("width", (str7 & "px"))
                                cell2.Text = "&nbsp;"
                                cell2 = Nothing
                            End If
                            Dim row4 As TableRow = row
                            row4.CssClass = "trCamposFiltro"
                            row4.Cells.Add(cell)
                            row4 = Nothing
                        End If
                        Me.i += 1
                    Loop
                End If
                child.Rows.Add(row)
                If (Me.FilterType = PFilterType.FullFilter) Then
                    If Me.FilterPagePreLoad Then
                        Dim enumerator As IEnumerator
                        Try
                            enumerator = filterData.GetEnumerator
                            Do While enumerator.MoveNext
                                current = DirectCast(enumerator.Current, DataRowView)
                                Me.x += 1
                                row = New TableRow
                                Dim num6 As Integer = filterData.Table.Columns.Count
                                Me.i = 1
                                Do While (Me.i <= num6)
                                    If Conversions.ToBoolean(NewLateBinding.LateIndexGet(Me._FilterCols.Item(Me.i), New Object() {"Visible"}, Nothing)) Then
                                        Dim view3 As DataRowView
                                        Dim num7 As Integer
                                        Dim objectValue As Object
                                        Dim str5 As String = Conversions.ToString(NewLateBinding.LateIndexGet(Me._FilterCols.Item(Me.i), New Object() {"PageURLDestino"}, Nothing))
                                        Dim str4 As String = Conversions.ToString(NewLateBinding.LateIndexGet(Me._FilterCols.Item(Me.i), New Object() {"PageURLColVar"}, Nothing))
                                        If (str5 <> "") Then
                                            cell = New TableCell
                                            If (str4 = "") Then
                                                str4 = Conversions.ToString(NewLateBinding.LateIndexGet(Me._FilterCols.Item(Me.i), New Object() {"Name"}, Nothing))
                                            End If
                                            If Me.FilterEncryptValue Then
                                                row.Attributes.Add("onclick", String.Concat(New String() {"window.location.href=""", str5, "?", str4, "=", CorpCripto.EncryptString(Conversions.ToString(current.Item((Me.i - 1)))), Me.GetQueryString(current), """"}))
                                            Else
                                                row.Attributes.Add("onclick", Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject((("window.location.href=""" & str5 & "?") & str4 & "="), current.Item((Me.i - 1))), Me.GetQueryString(current)), """")))
                                            End If
                                            view3 = current
                                            num7 = (Me.i - 1)
                                            objectValue = RuntimeHelpers.GetObjectValue(view3.Item(num7))
                                            view3.Item(num7) = RuntimeHelpers.GetObjectValue(objectValue)
                                            cell.Text = Conversions.ToString(Me.NullDB(objectValue, ""))
                                            row.Cells.Add(cell)
                                        Else
                                            cell = New TableCell
                                            view3 = current
                                            num7 = (Me.i - 1)
                                            objectValue = RuntimeHelpers.GetObjectValue(view3.Item(num7))
                                            view3.Item(num7) = RuntimeHelpers.GetObjectValue(objectValue)
                                            cell.Text = Conversions.ToString(Me.NullDB(objectValue, ""))
                                            row.Cells.Add(cell)
                                        End If
                                    End If
                                    Me.i += 1
                                Loop
                                Dim row5 As TableRow = row
                                If ((Me.x Mod 2) <> 0) Then
                                    row5.Attributes.Add("onmouseover", "jsFilter.mouseOverOut(event);")
                                    row5.CssClass = "darkTD"
                                Else
                                    row5.Attributes.Add("onmouseover", "jsFilter.mouseOverOut(event);")
                                    row5.CssClass = "lightTD"
                                End If
                                row5 = Nothing
                                child.Rows.Add(row)
                            Loop
                        Finally
                            If TypeOf enumerator Is IDisposable Then
                                TryCast(enumerator, IDisposable).Dispose()
                            End If
                        End Try
                    End If
                ElseIf Me.FilterPagePreLoad Then
                    Dim enumerator2 As IEnumerator
                    Try
                        enumerator2 = filterData.GetEnumerator
                        Do While enumerator2.MoveNext
                            current = DirectCast(enumerator2.Current, DataRowView)
                            Dim list As New ArrayList
                            row = New TableRow
                            Dim num8 As Integer = filterData.Table.Columns.Count
                            Me.i = 1
                            Do While (Me.i <= num8)
                                If Conversions.ToBoolean(NewLateBinding.LateIndexGet(Me._FilterCols.Item(Me.i), New Object() {"Visible"}, Nothing)) Then
                                    Dim str9 As String = NewLateBinding.LateIndexGet(Me._FilterCols.Item(Me.i), New Object() {"Name"}, Nothing).ToString.ToLower
                                    If (str9 = Me.FilterColText.ToLower) Then
                                        Me._FilterColText = Conversions.ToString(current.Item((Me.i - 1)))
                                    ElseIf (str9 = Me.FilterColValue.ToLower) Then
                                        Me._FilterColValue = Conversions.ToString(current.Item((Me.i - 1)))
                                    End If
                                    list.Add(RuntimeHelpers.GetObjectValue(current.Item((Me.i - 1))))
                                End If
                                Me.i += 1
                            Loop
                            cell = New TableCell With { _
                                .Text = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(("<a href=""javascript:" & Me.Page.ClientScript.GetPostBackEventReference(Me, ("filter_cellclick$" & Me._FilterColText.Replace("""", "") & "$" & Me._FilterColValue)) & """>"), list.Item(0)), "</a>"), ChrW(13) & ChrW(10))) _
                            }
                            row.Cells.Add(cell)
                            Dim num9 As Integer = (list.Count - 1)
                            Me.i = 1
                            Do While (Me.i <= num9)
                                cell = New TableCell With { _
                                    .Text = Conversions.ToString(list.Item(Me.i)) _
                                }
                                row.Cells.Add(cell)
                                Me.i += 1
                            Loop
                            child.Rows.Add(row)
                        Loop
                    Finally
                        If TypeOf enumerator2 Is IDisposable Then
                            TryCast(enumerator2, IDisposable).Dispose()
                        End If
                    End Try
                End If
                child.RenderControl(writer)
                Dim response As HttpResponse = Me.Page.Response
                response.Write(builder.ToString)
                response.Write("<div style=""margin-left:5px;padding:4px"">")
                response.Write(writer2.ToString)
                response.Write("</div>")
                response.Write(builder.ToString)
                response = Nothing
            End If
            Me.Page.ClientScript.RegisterStartupScript(Me.Page.GetType, "js_filter_init", ("jsFilter.init([" & Strings.Mid(str, 1, (str.Length - 2)) & "]);"), True)
            control.Controls.Add(child)
            Return control
        End Function

        Private Function TrataSinal(ByVal sItem As String, ByVal iTipoDado As ClsFilterCols.TypeDB) As String
            Dim flag As Boolean
            Dim message As String = ClsTools.CheckCommandSQL(sItem)
            If (message <> String.Empty) Then
                Throw New Exception(message)
            End If
            Dim num2 As Integer = -1
            Dim instance As Object = Strings.Split("%x<>x>=x<=x<x>", "x", -1, CompareMethod.Binary)
            sItem = sItem.Replace("'", "''")
            Dim num3 As Integer = Information.UBound(DirectCast(instance, Array), 1)
            Dim i As Integer = 0
            Do While (i <= num3)
                If (Strings.InStr(sItem, Conversions.ToString(NewLateBinding.LateIndexGet(instance, New Object() {i}, Nothing)), CompareMethod.Binary) <> 0) Then
                    num2 = i
                    flag = True
                    Exit Do
                End If
                i += 1
            Loop
            If flag Then
                Select Case num2
                    Case 0
                        Return (" LIKE '" & sItem & "'")
                End Select
                Return sItem
            End If
            Select Case iTipoDado
                Case ClsFilterCols.TypeDB.STRING_T
                    If (Strings.UCase(sItem) <> "IS NOT NULL") Then
                        If (Strings.UCase(sItem) = "IS NULL") Then
                            Return " IS NULL"
                        End If
                        If Versioned.IsNumeric(Strings.UCase(sItem)) Then
                            Return (" = '" & sItem & "'")
                        End If
                        If (Strings.UCase(sItem).Contains(">") And Versioned.IsNumeric(Strings.UCase(sItem).Replace(">", ""))) Then
                            Return (" > '" & sItem & "'")
                        End If
                        If (Strings.UCase(sItem).Contains("<") And Versioned.IsNumeric(Strings.UCase(sItem).Replace("<", ""))) Then
                            Return (" < '" & sItem & "'")
                        End If
                        If Information.IsDate(Strings.UCase(sItem)) Then
                            Return String.Concat(New String() {" between '", Strings.Format(Conversions.ToDate(sItem), "yyyy-MM-dd"), " 00:00' and '", Strings.Format(Conversions.ToDate(sItem), "yyyy-MM-dd"), " 23:59'"})
                        End If
                        If Strings.UCase(sItem).Contains("*") Then
                            Return (" LIKE '%" & sItem.Replace("*", "") & "%'")
                        End If
                        Return (" LIKE '%" & sItem & "%'")
                    End If
                    Return " IS NOT NULL"
                Case ClsFilterCols.TypeDB.DATE_T
                    Return (" = '" & Strings.Format(Conversions.ToDate(sItem), "yyyy-MM-dd") & "'")
                Case ClsFilterCols.TypeDB.DATE_TIME_T
                    Return (" = '" & Strings.Format(Conversions.ToDate(sItem), "yyyy-MM-dd H:mm:ss") & "'")
                Case ClsFilterCols.TypeDB.NUMERIC_T, ClsFilterCols.TypeDB.MONEY_T
                    Return (" = " & sItem)
            End Select
            Return ""
        End Function


        ' Properties
        Public Property FilterCol() As ClsFilterCols
            Get
                Return Me._ClsFilterCols
            End Get
            Set(ByVal value As ClsFilterCols)
                Me._ClsFilterCols = value
            End Set
        End Property

        <DefaultValue(""), Category("Appearance"), Bindable(True), Localizable(True)> _
        Public Property FilterColText() As String
            Get
                Dim str2 As String = Conversions.ToString(Me.ViewState.Item("FilterColText"))
                If (str2 Is Nothing) Then
                    Return String.Empty
                End If
                Return str2
            End Get
            Set(ByVal Value As String)
                Me.ViewState.Item("FilterColText") = Value
            End Set
        End Property

        <Bindable(True), Category("Appearance"), Localizable(True), DefaultValue("")> _
        Public Property FilterColValue() As String
            Get
                Dim str2 As String = Conversions.ToString(Me.ViewState.Item("FilterColValue"))
                If (str2 Is Nothing) Then
                    Return String.Empty
                End If
                Return str2
            End Get
            Set(ByVal Value As String)
                Me.ViewState.Item("FilterColValue") = Value
            End Set
        End Property

        <Category("Appearance"), Bindable(True), Localizable(True), DefaultValue(False)> _
        Public Property FilterEncryptValue() As Boolean
            Get
                Dim str As String = Conversions.ToString(Me.ViewState.Item("FilterEncryptValue"))
                If (str Is Nothing) Then
                    Return False
                End If
                Return Conversions.ToBoolean(str)
            End Get
            Set(ByVal Value As Boolean)
                Me.ViewState.Item("FilterEncryptValue") = Value
            End Set
        End Property

        <Localizable(True), DefaultValue("Filter"), Category("Appearance"), Bindable(True)> _
        Public Property FilterInitialOrderBy() As String
            Get
                Dim str2 As String = Conversions.ToString(Me.ViewState.Item("FilterInitialOrderBy"))
                If (str2 Is Nothing) Then
                    Return String.Empty
                End If
                Return str2
            End Get
            Set(ByVal Value As String)
                Me.ViewState.Item("FilterInitialOrderBy") = Value
            End Set
        End Property

        <Localizable(True), Bindable(True), DefaultValue("Filter"), Category("Appearance")> _
        Public Property FilterJsUrl() As String
            Get
                Dim str2 As String = Conversions.ToString(Me.ViewState.Item("FilterJsUrl"))
                If (str2 Is Nothing) Then
                    Return Me.MyPath
                End If
                Return str2
            End Get
            Set(ByVal Value As String)
                Me.ViewState.Item("FilterJsUrl") = Value
            End Set
        End Property

        <Category("Appearance"), Bindable(True), DefaultValue(False), Localizable(True)> _
        Public Property FilterOrderByCols() As Boolean
            Get
                Dim str As String = Conversions.ToString(Me.ViewState.Item("FilterOrderByCols"))
                If (str Is Nothing) Then
                    Return False
                End If
                Return Conversions.ToBoolean(str)
            End Get
            Set(ByVal Value As Boolean)
                Me.ViewState.Item("FilterOrderByCols") = Value
            End Set
        End Property

        <Localizable(True), DefaultValue(False), Category("Appearance"), Bindable(True)> _
        Public Property FilterPagePreLoad() As Boolean
            Get
                Dim str As String = Conversions.ToString(Me.ViewState.Item("FilterPagePreLoad"))
                If (str Is Nothing) Then
                    Return False
                End If
                Return Conversions.ToBoolean(str)
            End Get
            Set(ByVal Value As Boolean)
                Me.ViewState.Item("FilterPagePreLoad") = Value
            End Set
        End Property

        <Localizable(True), DefaultValue(10), Category("Appearance"), Bindable(True)> _
        Public Property FilterPageSize() As Integer
            Get
                Dim str As String = Conversions.ToString(Me.ViewState.Item("FilterPageSize"))
                If (str Is Nothing) Then
                    Return 0
                End If
                Return Conversions.ToInteger(str)
            End Get
            Set(ByVal Value As Integer)
                Me.ViewState.Item("FilterPageSize") = Value
            End Set
        End Property

        <Category("Appearance"), DefaultValue(""), Localizable(True), Bindable(True)> _
        Public Property FilterParamWhere() As String
            Get
                Dim str2 As String = Conversions.ToString(Me.ViewState.Item("FilterParamWhere"))
                If (str2 Is Nothing) Then
                    Return String.Empty
                End If
                Return str2
            End Get
            Set(ByVal Value As String)
                Me.ViewState.Item("FilterParamWhere") = Value
            End Set
        End Property

        Public Property FilterProvider() As ClsDB1.T_PROVIDER
            Get
                If (Me.ViewState.Item("FilterProvider") Is Nothing) Then
                    Me.ViewState.Item("FilterProvider") = ClsDB1.T_PROVIDER.SQL
                End If
                Return DirectCast(Conversions.ToInteger(Me.ViewState.Item("FilterProvider")), ClsDB1.T_PROVIDER)
            End Get
            Set(ByVal value As ClsDB1.T_PROVIDER)
                Me.ViewState.Item("FilterProvider") = value
            End Set
        End Property

        <Bindable(True), Category("Appearance"), Localizable(True), DefaultValue("")> _
        Public Property FilterReturnFormName() As String
            Get
                Dim str2 As String = Conversions.ToString(Me.ViewState.Item("FilterReturnFormName"))
                If (str2 Is Nothing) Then
                    Return String.Empty
                End If
                Return str2
            End Get
            Set(ByVal Value As String)
                Me.ViewState.Item("FilterReturnFormName") = Value
            End Set
        End Property

        Public Property FilterStateView() As Collection
            Get
                Dim objectValue As Object = RuntimeHelpers.GetObjectValue(Me.ViewState.Item("FilterStateView"))
                If Information.IsNothing(RuntimeHelpers.GetObjectValue(objectValue)) Then
                    Return Nothing
                End If
                Return DirectCast(objectValue, Collection)
            End Get
            Set(ByVal Value As Collection)
                Me.ViewState.Add("FilterStateView", Value)
            End Set
        End Property

        <Bindable(True), DefaultValue(""), Category("Appearance"), Localizable(True)> _
        Public Property FilterStrConnection() As String
            Get
                Dim str2 As String = Conversions.ToString(Me.ViewState.Item("FilterStrConnection"))
                If (str2 Is Nothing) Then
                    Return String.Empty
                End If
                Return str2
            End Get
            Set(ByVal Value As String)
                Me.ViewState.Item("FilterStrConnection") = Value
            End Set
        End Property

        <Category("Appearance"), Localizable(True), DefaultValue(""), Bindable(True)> _
        Public Property FilterTableName() As String
            Get
                Dim str2 As String = Conversions.ToString(Me.ViewState.Item("FilterTableName"))
                If (str2 Is Nothing) Then
                    Return String.Empty
                End If
                Return str2
            End Get
            Set(ByVal Value As String)
                Me.ViewState.Item("FilterTableName") = Value
            End Set
        End Property

        <Bindable(True), Localizable(True), DefaultValue("Filter"), Category("Appearance")> _
        Public Property FilterTitle() As String
            Get
                Dim str2 As String = Conversions.ToString(Me.ViewState.Item("FilterTitle"))
                If (str2 Is Nothing) Then
                    Return String.Empty
                End If
                Return str2
            End Get
            Set(ByVal Value As String)
                Me.ViewState.Item("FilterTitle") = Value
            End Set
        End Property

        <Bindable(True), Localizable(True), DefaultValue(1), Category("Appearance")> _
        Public Property FilterType() As PFilterType
            Get
                Dim str As String = Conversions.ToString(Me.ViewState.Item("FilterType"))
                If (str Is Nothing) Then
                    Return DirectCast(0, PFilterType)
                End If
                Return DirectCast(Conversions.ToInteger(str), PFilterType)
            End Get
            Set(ByVal Value As PFilterType)
                Me.ViewState.Item("FilterType") = CInt(Value)
            End Set
        End Property


        ' Fields
        Private _ClsFilterCols As ClsFilterCols
        Private _ds As DataSet
        Private _FILTER_VIEWSTATE As String
        Private _FilterColOrderByPos As String
        Private _FilterCols As Collection
        Private _FilterColSimpleText As String
        Private _FilterColSimpleValue As String
        Private _FilterColText As String
        Private _FilterColValue As String
        Private _FilterLastOrderByColName As String
        Private _FilterNotVisible As Boolean
        Private _FilterOptionClear As Boolean
        Private _FilterOrderByColName As String
        Private _FilterQueryString As DataTable
        Private _FilterSimpleText As String
        Private _FilterState As PropertyCollection
        Private _FilterViewState As Collection
        Private _GetSQLOrderBy As String
        Private _GetSQLWhere As String
        Private _PageAtiva As Long
        Private _PageFirst As Long
        Private _PageLast As Long
        Private _PageLastRows As Long
        Private _PageNext As Long
        Private _PagePrevious As Long
        Private _PageQtde As Long
        Private _TotalRows As Long
        Private FilterPageRowPosition As Long
        Private i As Integer
        Private l As Long
        Private x As Integer

        ' Nested Types
        Public Delegate Sub FILTER_CellClickEventHandler(ByVal sText As String, ByVal sValue As String)

        Public Enum PFilterType
            ' Fields
            FullFilter = 1
            SimpleFilter = 2
        End Enum
    End Class
End Namespace

