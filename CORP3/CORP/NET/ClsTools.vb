Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Collections
Imports System.Drawing
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Web.UI
Imports System.Web.UI.HtmlControls
Imports System.Web.UI.WebControls

Namespace CORP3.NET
    Public Class ClsTools
        Inherits WebControl
        ' Methods
        Public Shared Function CheckCommandSQL(ByVal strExpressao As String) As String
            Dim message As String
            Try
                Dim instance As Object = Strings.Split(strExpressao, " ", -1, CompareMethod.Binary)
                Dim obj3 As Object = Strings.Split("select#insert#update#delete#drop#--#'", "#", -1, CompareMethod.Binary)
                Dim num3 As Integer = Information.UBound(DirectCast(obj3, Array), 1)
                Dim i As Integer = 0
                Do While (i <= num3)
                    Dim num4 As Integer = Information.UBound(DirectCast(instance, Array), 1)
                    Dim j As Integer = 0
                    Do While (j <= num4)
                        Dim arguments As Object() = New Object() {RuntimeHelpers.GetObjectValue(NewLateBinding.LateIndexGet(obj3, New Object() {i}, Nothing))}
                        Dim copyBack As Boolean() = New Boolean() {True}
                        If copyBack(0) Then
                            NewLateBinding.LateIndexSetComplex(obj3, New Object() {i, RuntimeHelpers.GetObjectValue(arguments(0))}, Nothing, True, False)
                        End If
                        Dim objArray7 As Object() = New Object() {RuntimeHelpers.GetObjectValue(NewLateBinding.LateIndexGet(instance, New Object() {j}, Nothing))}
                        Dim flagArray2 As Boolean() = New Boolean() {True}
                        If flagArray2(0) Then
                            NewLateBinding.LateIndexSetComplex(instance, New Object() {j, RuntimeHelpers.GetObjectValue(objArray7(0))}, Nothing, True, False)
                        End If
                        If Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(Nothing, GetType(Strings), "LCase", arguments, Nothing, Nothing, copyBack), NewLateBinding.LateGet(Nothing, GetType(Strings), "LCase", objArray7, Nothing, Nothing, flagArray2), False) Then
                            Throw New Exception("Comando Arbitrário")
                        End If
                        j += 1
                    Loop
                    i += 1
                Loop
                message = String.Empty
            Catch exception1 As Exception
                ProjectData.SetProjectError(exception1)
                Dim exception As Exception = exception1
                message = exception.Message
                ProjectData.ClearProjectError()
            End Try
            Return message
        End Function

        Public Sub LimpaCampos(ByRef Form As Object)
            Dim enumerator As IEnumerator
            Try
                enumerator = DirectCast(NewLateBinding.LateGet(Form, Nothing, "Controls", New Object(0 - 1) {}, Nothing, Nothing, Nothing), IEnumerable).GetEnumerator
                Do While enumerator.MoveNext
                    Dim objectValue As Object = RuntimeHelpers.GetObjectValue(enumerator.Current)
                    Dim str As String = objectValue.GetType.ToString
                    If (str = "System.Web.UI.WebControls.TextBox") Then
                        DirectCast(objectValue, TextBox).Text = String.Empty
                    Else
                        If (str = "System.Web.UI.WebControls.DropDownList") Then
                            DirectCast(objectValue, DropDownList).SelectedIndex = -1
                            Continue Do
                        End If
                        If (str = "System.Web.UI.WebControls.RadioButton") Then
                            DirectCast(objectValue, RadioButton).Checked = False
                            Continue Do
                        End If
                        If (str = "System.Web.UI.WebControls.CheckBox") Then
                            DirectCast(objectValue, CheckBox).Checked = False
                        End If
                    End If
                Loop
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try
        End Sub

        Public Function RemTags(ByVal strHTML As String) As String
            Dim regex As New Regex("<(.|\n)+?>")
            Return regex.Replace(strHTML, "")
        End Function

        Public Sub SetReadOnly(ByRef Form As Object, Optional ByVal blnStatus As Boolean = True)
            Dim enumerator As IEnumerator
            Try
                enumerator = DirectCast(NewLateBinding.LateGet(Form, Nothing, "Controls", New Object(0 - 1) {}, Nothing, Nothing, Nothing), IEnumerable).GetEnumerator
                Do While enumerator.MoveNext
                    Dim objectValue As Object = RuntimeHelpers.GetObjectValue(enumerator.Current)
                    Select Case objectValue.GetType.ToString
                        Case "System.Web.UI.WebControls.TextBox"
                            Dim box As TextBox = DirectCast(objectValue, TextBox)
                            box.ReadOnly = blnStatus
                            If blnStatus Then
                                box.BackColor = Color.FromArgb(0, 240, 240, 240)
                            Else
                                box.BackColor = Color.FromArgb(0, &HFF, &HFF, &HFF)
                            End If
                            box = Nothing
                            Continue Do
                        Case "System.Web.UI.WebControls.RadioButton", "System.Web.UI.WebControls.CheckBox", "System.Web.UI.WebControls.DropDownList"
                            If (objectValue.GetType.ToString = "System.Web.UI.WebControls.DropDownList") Then
                                If blnStatus Then
                                    NewLateBinding.LateSet(objectValue, Nothing, "BackColor", New Object() {Color.FromArgb(0, 240, 240, 240)}, Nothing, Nothing)
                                Else
                                    NewLateBinding.LateSet(objectValue, Nothing, "BackColor", New Object() {Color.FromArgb(0, &HFF, &HFF, &HFF)}, Nothing, Nothing)
                                End If
                            End If
                            NewLateBinding.LateSet(objectValue, Nothing, "Enabled", New Object() {Not blnStatus}, Nothing, Nothing)
                            Exit Select
                    End Select
                Loop
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try
        End Sub

        Public Sub ShowMessage(ByVal strTextoMsg As String, ByVal strTituloMsg As String, ByVal sender As Object, Optional ByVal msgStyle As TMsgStyleIcon = 0, Optional ByVal strScriptTag As String = "", Optional ByVal strScript As String = "")
            Dim objArray As Object()
            Dim flagArray As Boolean()
            Dim table As New Table
            Dim child As New Table
            Dim builder As New StringBuilder
            Dim image As New System.Web.UI.WebControls.Image
            Dim label2 As New Label
            Dim label As New Label
            Dim button As New HtmlInputButton
            Dim table3 As Table = child
            table3.ID = "TabMens"
            table3.Style.Add("width", "350px")
            table3.CellPadding = 0
            table3.CellSpacing = 0
            table3 = Nothing
            Dim row As New TableRow
            Select Case msgStyle
                Case TMsgStyleIcon.MSG_ERROR
                    image.ImageUrl = "imagens/msg/error.gif"
                    Exit Select
                Case TMsgStyleIcon.MSG_WARNING
                    image.ImageUrl = "imagens/msg/warning.gif"
                    Exit Select
                Case TMsgStyleIcon.MSG_INFORMATION
                    image.ImageUrl = "imagens/msg/information.gif"
                    Exit Select
            End Select
            Dim builder2 As StringBuilder = builder
            builder2.AppendLine()
            builder2.AppendLine("function ShowMsg() {")
            builder2.AppendLine(ChrW(9) & "var x = document.getElementById('TabMens');")
            builder2.AppendLine(" " & ChrW(9) & "var ifra = document.getElementById('MsgIframe').style;")
            builder2.AppendLine(" " & ChrW(9) & "var shd = document.getElementById('shd').style;")
            builder2.AppendLine(" " & ChrW(9) & "var vTabCentro = document.getElementById('TabCentro');")
            builder2.AppendLine(ChrW(9) & "vTabCentro.style.height = document.body.offsetHeight + 'px';")
            builder2.AppendLine(ChrW(9) & "vTabCentro.style.width = document.body.offsetWidth + 'px';")
            builder2.AppendLine()
            builder2.AppendLine("   with (shd) {")
            builder2.AppendLine("        top=x.offsetTop + 10 + 'px';")
            builder2.AppendLine("        left=x.offsetLeft + 10  + 'px';")
            builder2.AppendLine("        width=x.offsetWidth + 'px';")
            builder2.AppendLine("        height=x.offsetHeight + 'px';")
            builder2.AppendLine("        visibility='inherit';")
            builder2.AppendLine("        zIndex=0;")
            builder2.AppendLine("    }")
            builder2.AppendLine()
            builder2.AppendLine("    with (ifra) {")
            builder2.AppendLine("    " & ChrW(9) & "top = '0px';")
            builder2.AppendLine("    " & ChrW(9) & "left = '0px';")
            builder2.AppendLine("    " & ChrW(9) & "width = document.body.offsetWidth + 'px';")
            builder2.AppendLine("    " & ChrW(9) & "height = document.body.offsetHeight + 'px';")
            builder2.AppendLine("    " & ChrW(9) & "position='absolute';")
            builder2.AppendLine("    " & ChrW(9) & "visibility='inherit';")
            builder2.AppendLine("    " & ChrW(9) & "zIndex=2;")
            builder2.AppendLine("    " & ChrW(9) & "border =0;")
            builder2.AppendLine(ChrW(9) & "}")
            builder2.AppendLine(ChrW(9) & "x.style.zIndex=9;")
            builder2.AppendLine("}")
            builder2.AppendLine()
            builder2.AppendLine("window.onload = ShowMsg;")
            builder2.AppendLine("window.onresize = ShowMsg; ")
            builder2.AppendLine()
            builder2.AppendLine("function CloseMsg() {")
            builder2.AppendLine(ChrW(9) & "document.getElementById('TabCentro').style.display='none';")
            builder2.AppendLine(ChrW(9) & "document.getElementById('MsgIframe').style.display='none';")
            builder2.AppendLine(ChrW(9) & "document.getElementById('shd').style.display='none';")
            builder2.AppendLine("}")
            builder2 = Nothing
            If Conversions.ToBoolean(Operators.NotObject(NewLateBinding.LateGet(NewLateBinding.LateGet(sender, Nothing, "ClientScript", New Object(0 - 1) {}, Nothing, Nothing, Nothing), Nothing, "IsClientScriptBlockRegistered", New Object() {"msgscript"}, Nothing, Nothing, Nothing))) Then
                NewLateBinding.LateCall(NewLateBinding.LateGet(sender, Nothing, "ClientScript", New Object(0 - 1) {}, Nothing, Nothing, Nothing), Nothing, "RegisterClientScriptBlock", New Object() {Me.GetType, "msgscript", builder.ToString, True}, Nothing, Nothing, Nothing, True)
            End If
            Dim cell As TableCell = New TableHeaderCell
            Dim cell2 As TableCell = cell
            cell2.ColumnSpan = 2
            cell2.VerticalAlign = VerticalAlign.Middle
            cell2.HorizontalAlign = HorizontalAlign.Center
            cell2.Style.Add("height", "31px")
            If (strTituloMsg.Trim = "") Then
                strTituloMsg = "Mensagem"
            End If
            Dim label3 As Label = label2
            label3.Text = strTituloMsg.Replace(ChrW(13) & ChrW(10), "<br />")
            label3 = Nothing
            cell2.Controls.Add(label2)
            cell2 = Nothing
            row.Cells.Add(cell)
            child.Rows.Add(row)
            row = New TableRow
            cell = New TableCell
            Dim cell3 As TableCell = cell
            cell3.ID = "CellIMG"
            cell3.VerticalAlign = VerticalAlign.Top
            cell3.HorizontalAlign = HorizontalAlign.Center
            cell3.Style.Add("height", "70px")
            cell3.Controls.Add(image)
            cell3 = Nothing
            row.Cells.Add(cell)
            cell = New TableCell
            Dim cell4 As TableCell = cell
            cell4.ID = "CellText"
            cell4.VerticalAlign = VerticalAlign.Middle
            cell4.HorizontalAlign = HorizontalAlign.Center
            Dim style As CssStyleCollection = cell4.Style
            style.Add("width", "290px")
            style.Add("text-align", "justify")
            style.Add("vertical-align", "middle")
            style = Nothing
            If (strTextoMsg.Length > 100) Then
                label.Text = ("<div style='height:100px;overflow:auto'>" & strTextoMsg.Replace(ChrW(13) & ChrW(10), "<br />") & "</div>")
            Else
                label.Text = strTextoMsg
            End If
            cell4.Controls.Add(label)
            cell4 = Nothing
            row.Cells.Add(cell)
            child.Rows.Add(row)
            Dim button2 As HtmlInputButton = button
            button2.ID = "btnOK"
            button2.Value = "Ok"
            If (strScript.Trim = "") Then
                button2.Attributes.Add("onclick", "javascript:CloseMsg();")
            ElseIf (strScriptTag.Trim = "") Then
                button2.Attributes.Add("onclick", ("javascript:CloseMsg();" & strScript))
            Else
                button2.Attributes.Add("onclick", ("javascript:CloseMsg();" & strScriptTag))
                objArray = New Object() {sender.GetType, strScriptTag, strScript, True}
                flagArray = New Boolean() {False, True, True, False}
                NewLateBinding.LateCall(NewLateBinding.LateGet(sender, Nothing, "ClientScript", New Object(0 - 1) {}, Nothing, Nothing, Nothing), Nothing, "RegisterClientScriptBlock", objArray, Nothing, Nothing, flagArray, True)
                If flagArray(1) Then
                    strScriptTag = CStr(Conversions.ChangeType(RuntimeHelpers.GetObjectValue(objArray(1)), GetType(String)))
                End If
                If flagArray(2) Then
                    strScript = CStr(Conversions.ChangeType(RuntimeHelpers.GetObjectValue(objArray(2)), GetType(String)))
                End If
            End If
            button2.Attributes.Add("title", " Clique para fechar janela ")
            button2 = Nothing
            row = New TableRow
            cell = New TableCell
            Dim cell5 As TableCell = cell
            cell5.ColumnSpan = 2
            cell5.VerticalAlign = VerticalAlign.Middle
            cell5.HorizontalAlign = HorizontalAlign.Center
            cell5.Style.Add("height", "50px")
            cell5.Controls.Add(button)
            cell5 = Nothing
            row.Cells.Add(cell)
            child.Rows.Add(row)
            Dim table4 As Table = table
            table4.ID = "TabCentro"
            table4.Attributes.Add("onload", "javascript:ShowMsg()")
            Dim styles2 As CssStyleCollection = table4.Style
            styles2.Add("top", "1px")
            styles2.Add("left", "1px")
            styles2.Add("z-index", "3")
            styles2.Add("position", "absolute")
            styles2 = Nothing
            table4.BorderWidth = Unit.Pixel(0)
            table4.CellPadding = 0
            table4.CellSpacing = 0
            table4 = Nothing
            row = New TableRow
            cell = New TableCell
            Dim cell6 As TableCell = cell
            cell6.HorizontalAlign = HorizontalAlign.Center
            cell6.VerticalAlign = VerticalAlign.Middle
            cell6.Height = Unit.Pixel(Conversions.ToInteger("420"))
            cell6.Controls.Add(child)
            cell6 = Nothing
            row.Cells.Add(cell)
            table.Rows.Add(row)
            objArray = New Object() {table}
            flagArray = New Boolean() {True}
            NewLateBinding.LateCall(NewLateBinding.LateGet(sender, Nothing, "Controls", New Object(0 - 1) {}, Nothing, Nothing, Nothing), Nothing, "Add", objArray, Nothing, Nothing, flagArray, True)
            If flagArray(0) Then
                table = DirectCast(Conversions.ChangeType(RuntimeHelpers.GetObjectValue(objArray(0)), GetType(Table)), Table)
            End If
            NewLateBinding.LateCall(NewLateBinding.LateGet(sender, Nothing, "Controls", New Object(0 - 1) {}, Nothing, Nothing, Nothing), Nothing, "Add", New Object() {New LiteralControl("<DIV ID=""shd""></DIV>")}, Nothing, Nothing, Nothing, True)
            NewLateBinding.LateCall(NewLateBinding.LateGet(sender, Nothing, "Controls", New Object(0 - 1) {}, Nothing, Nothing, Nothing), Nothing, "Add", New Object() {New LiteralControl("<IFRAME style='visibility:hidden;z-index:2;FILTER: Alpha(Opacity=0);' src='about:blank' id='MsgIframe'></IFRAME>")}, Nothing, Nothing, Nothing, True)
        End Sub


        ' Nested Types
        Public Enum TMsgStyleIcon
            ' Fields
            MSG_ERROR = 0
            MSG_INFORMATION = 2
            MSG_WARNING = 1
        End Enum
    End Class
End Namespace

