Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Collections
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Web.UI.WebControls

Namespace CORP3.NET
    Public Class ClsSQLParam1
        ' Methods
        Public Sub New(ByVal cmd As DbCommand)
            Me.mCmd = cmd
        End Sub

        Public Sub AddColuna(ByVal Coluna As String, ByRef Value As Object, Optional ByVal TType As DbType = &H10, Optional ByRef Evt As EventoExterno = Nothing)
            Me.mTpSQL = New TpSQL
            Me.mTpSQL.Coluna = Coluna
            Me.mTpSQL.Value = RuntimeHelpers.GetObjectValue(Value)
            Me.mTpSQL.TType = TType
            Me.mTpSQL.EventoExterno = Evt
            Me.mColuna.Add(Me.mTpSQL, Coluna, Nothing, Nothing)
        End Sub

        Public Sub AddColunaWhere(ByVal Coluna As String, ByRef Value As Object, Optional ByVal TType As DbType = &H10, Optional ByRef Evt As EventoExterno = Nothing)
            Me.mTpSQL = New TpSQL
            Me.mTpSQL.Coluna = Coluna
            Me.mTpSQL.Value = RuntimeHelpers.GetObjectValue(Value)
            Me.mTpSQL.TType = TType
            Me.mTpSQL.EventoExterno = Evt
            Me.mColunaWhere.Add(Me.mTpSQL, Nothing, Nothing, Nothing)
        End Sub

        Public Function ALTERAR() As DbCommand
            Dim parameter As DbParameter
            Dim enumerator As IEnumerator
            Dim eventoExterno As Object
            Dim enumerator2 As IEnumerator
            Dim builder As New StringBuilder
            Dim builder2 As New StringBuilder
            Dim num As Short = 0
            Dim str As String = " "
            Try
                enumerator = Me.mColuna.GetEnumerator
                Do While enumerator.MoveNext
                    Dim current As TpSQL = DirectCast(enumerator.Current, TpSQL)
                    If TypeOf Me.mCmd Is SqlCommand Then
                        num = CShort((num + 1))
                        builder.AppendLine(String.Concat(New String() {str, current.Coluna, "=", Me.PipeCommend, current.Coluna, Conversions.ToString(CInt(num))}))
                    Else
                        builder.AppendLine((str & current.Coluna & "=" & Me.PipeCommend))
                    End If
                    str = ","
                    parameter = Me.mCmd.CreateParameter
                    Dim parameter2 As DbParameter = parameter
                    If TypeOf Me.mCmd Is SqlCommand Then
                        parameter2.ParameterName = (Me.PipeCommend & current.Coluna & Conversions.ToString(CInt(num)))
                    Else
                        parameter2.ParameterName = current.Coluna
                    End If
                    parameter2.DbType = current.TType
                    eventoExterno = current.EventoExterno
                    Dim introduced16 As Object = Me.GetValueControle(current.Value, current.TType, eventoExterno)
                    current.EventoExterno = DirectCast(eventoExterno, EventoExterno)
                    parameter2.Value = RuntimeHelpers.GetObjectValue(introduced16)
                    parameter2 = Nothing
                    Me.mCmd.Parameters.Add(parameter)
                Loop
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try
            str = ""
            Try
                enumerator2 = Me.mColunaWhere.GetEnumerator
                Do While enumerator2.MoveNext
                    Dim psql2 As TpSQL = DirectCast(enumerator2.Current, TpSQL)
                    If TypeOf Me.mCmd Is SqlCommand Then
                        num = CShort((num + 1))
                        builder2.AppendLine(String.Concat(New String() {str, psql2.Coluna, "=", Me.PipeCommend, psql2.Coluna, Conversions.ToString(CInt(num))}))
                    Else
                        builder2.AppendLine((str & psql2.Coluna & "=" & Me.PipeCommend))
                    End If
                    str = " and "
                    parameter = Me.mCmd.CreateParameter
                    Dim parameter3 As DbParameter = parameter
                    If TypeOf Me.mCmd Is SqlCommand Then
                        parameter3.ParameterName = (psql2.Coluna & Conversions.ToString(CInt(num)))
                    Else
                        parameter3.ParameterName = psql2.Coluna
                    End If
                    parameter3.DbType = psql2.TType
                    eventoExterno = psql2.EventoExterno
                    Dim introduced17 As Object = Me.GetValueControle(psql2.Value, psql2.TType, eventoExterno)
                    psql2.EventoExterno = DirectCast(eventoExterno, EventoExterno)
                    parameter3.Value = RuntimeHelpers.GetObjectValue(introduced17)
                    parameter3 = Nothing
                    Me.mCmd.Parameters.Add(parameter)
                Loop
            Finally
                If TypeOf enumerator2 Is IDisposable Then
                    TryCast(enumerator2, IDisposable).Dispose()
                End If
            End Try
            Dim mStrSQL As StringBuilder = Me.mStrSQL
            mStrSQL.AppendLine(("UPDATE " & Me.mTabela))
            mStrSQL.AppendLine("SET ")
            mStrSQL.AppendLine(builder.ToString)
            mStrSQL.AppendLine("WHERE ")
            mStrSQL.AppendLine(builder2.ToString)
            mStrSQL = Nothing
            Me.mCmd.CommandText = Me.mStrSQL.ToString
            Return Me.mCmd
        End Function

        Public Overridable Function ChkTipoDB(ByVal Crt As Object, ByVal TType As DbType) As Object
            Select Case TType
                Case DbType.Currency, DbType.Decimal, DbType.Double, DbType.Int16, DbType.Int32, DbType.Int64, DbType.Single
                    If Not Versioned.IsNumeric(RuntimeHelpers.GetObjectValue(Crt)) Then
                        Crt = DBNull.Value
                    End If
                    Return Crt
                Case DbType.Date, DbType.Guid, DbType.Object, DbType.SByte, DbType.String, DbType.Time, DbType.UInt16, DbType.UInt32, DbType.UInt64, DbType.VarNumeric, DbType.AnsiStringFixedLength, DbType.StringFixedLength, (DbType.String Or DbType.Double), DbType.Xml
                    Return Crt
                Case DbType.DateTime, DbType.DateTime2, DbType.DateTimeOffset
                    If Not Information.IsDate(RuntimeHelpers.GetObjectValue(Crt)) Then
                        Crt = DBNull.Value
                    End If
                    Return Crt
            End Select
            Return Crt
        End Function

        Public Function CONTROLS(ByVal tbTableSource As DataTable) As Boolean
            Dim enumerator As IEnumerator
            If (tbTableSource.Rows.Count <> 1) Then
                Return False
            End If
            Try
                enumerator = Me.mColuna.GetEnumerator
                Do While enumerator.MoveNext
                    Dim current As TpSQL = DirectCast(enumerator.Current, TpSQL)
                    Dim objectValue As Object = RuntimeHelpers.GetObjectValue(tbTableSource.Rows.Item(0).Item(current.Coluna))
                    Dim eventoExterno As Object = current.EventoExterno
                    Me.SetValueControle(objectValue, current.Value, current.TType, eventoExterno)
                    current.EventoExterno = DirectCast(eventoExterno, EventoExterno)
                Loop
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try
            Return True
        End Function

        Private Function GetValueControle(ByRef Crt As Object, ByVal TType As DbType, ByRef evt As Object) As Object
            Dim text As Object = Nothing
            Select Case Crt.GetType.ToString
                Case "System.Web.UI.WebControls.TextBox"
                    [text] = DirectCast(Crt, TextBox).Text
                    Exit Select
                Case "System.Web.UI.WebControls.DropDownList"
                    [text] = DirectCast(Crt, DropDownList).SelectedValue.ToString
                    Exit Select
                Case "System.Web.UI.WebControls.ListBox"
                    [text] = DirectCast(Crt, ListBox).SelectedValue.ToString
                    Exit Select
                Case "System.Web.UI.WebControls.RadioButton"
                    [text] = DirectCast(Crt, RadioButton).Checked
                    Exit Select
                Case "System.Web.UI.WebControls.CheckBox"
                    [text] = DirectCast(Crt, CheckBox).Checked
                    Exit Select
                Case Else
                    If (Not evt Is Nothing) Then
                        Dim arguments As Object() = New Object() {RuntimeHelpers.GetObjectValue(Crt), RuntimeHelpers.GetObjectValue([text])}
                        Dim copyBack As Boolean() = New Boolean() {True, True}
                        If copyBack(0) Then
                            Crt = RuntimeHelpers.GetObjectValue(arguments(0))
                        End If
                        If copyBack(1) Then
                            [text] = RuntimeHelpers.GetObjectValue(arguments(1))
                        End If
                        [text] = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(evt, Nothing, "Invoke", arguments, Nothing, Nothing, copyBack))
                    Else
                        [text] = RuntimeHelpers.GetObjectValue(Crt)
                    End If
                    Exit Select
            End Select
            Return Me.ChkTipoDB(RuntimeHelpers.GetObjectValue([text]), TType)
        End Function

        Public Function PESQUISAR() As DbCommand
            Dim enumerator As IEnumerator
            Dim enumerator2 As IEnumerator
            Dim builder2 As New StringBuilder
            Dim builder As New StringBuilder
            Dim str As String = " "
            Dim num As Short = 0
            Try
                enumerator = Me.mColuna.GetEnumerator
                Do While enumerator.MoveNext
                    Dim current As TpSQL = DirectCast(enumerator.Current, TpSQL)
                    builder.AppendLine((str & current.Coluna))
                    str = ","
                Loop
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try
            str = ""
            Try
                enumerator2 = Me.mColunaWhere.GetEnumerator
                Do While enumerator2.MoveNext
                    Dim psql2 As TpSQL = DirectCast(enumerator2.Current, TpSQL)
                    If TypeOf Me.mCmd Is SqlCommand Then
                        num = CShort((num + 1))
                        builder2.AppendLine(String.Concat(New String() {str, psql2.Coluna, "=", Me.PipeCommend, psql2.Coluna, Conversions.ToString(CInt(num))}))
                    Else
                        builder2.AppendLine((str & psql2.Coluna & "=" & Me.PipeCommend))
                    End If
                    str = " and "
                    Dim parameter As DbParameter = Me.mCmd.CreateParameter
                    Dim parameter2 As DbParameter = parameter
                    If TypeOf Me.mCmd Is SqlCommand Then
                        parameter2.ParameterName = (psql2.Coluna & Conversions.ToString(CInt(num)))
                    Else
                        parameter2.ParameterName = psql2.Coluna
                    End If
                    parameter2.DbType = psql2.TType
                    Dim eventoExterno As Object = psql2.EventoExterno
                    Dim introduced15 As Object = Me.GetValueControle(psql2.Value, psql2.TType, eventoExterno)
                    psql2.EventoExterno = DirectCast(eventoExterno, EventoExterno)
                    parameter2.Value = RuntimeHelpers.GetObjectValue(introduced15)
                    parameter2 = Nothing
                    Me.mCmd.Parameters.Add(parameter)
                Loop
            Finally
                If TypeOf enumerator2 Is IDisposable Then
                    TryCast(enumerator2, IDisposable).Dispose()
                End If
            End Try
            Dim mStrSQL As StringBuilder = Me.mStrSQL
            mStrSQL.AppendLine("SELECT ")
            mStrSQL.Append(builder.ToString)
            mStrSQL.AppendLine(("FROM " & Me.mTabela))
            mStrSQL.AppendLine(("WHERE " & builder2.ToString))
            mStrSQL = Nothing
            Me.mCmd.CommandText = Me.mStrSQL.ToString
            Return Me.mCmd
        End Function

        Public Function SALVAR() As DbCommand
            Dim enumerator As IEnumerator
            Dim builder As New StringBuilder
            Dim builder2 As New StringBuilder
            Dim str As String = " "
            Try
                enumerator = Me.mColuna.GetEnumerator
                Do While enumerator.MoveNext
                    Dim current As TpSQL = DirectCast(enumerator.Current, TpSQL)
                    builder.AppendLine((str & current.Coluna))
                    If TypeOf Me.mCmd Is SqlCommand Then
                        builder2.AppendLine((str & Me.PipeCommend & current.Coluna))
                    Else
                        builder2.AppendLine((str & Me.PipeCommend))
                    End If
                    str = ","
                    Dim parameter As DbParameter = Me.mCmd.CreateParameter
                    Dim parameter2 As DbParameter = parameter
                    parameter2.ParameterName = current.Coluna
                    parameter2.DbType = current.TType
                    Dim eventoExterno As Object = current.EventoExterno
                    Dim introduced11 As Object = Me.GetValueControle(current.Value, current.TType, eventoExterno)
                    current.EventoExterno = DirectCast(eventoExterno, EventoExterno)
                    parameter2.Value = RuntimeHelpers.GetObjectValue(introduced11)
                    parameter2 = Nothing
                    Me.mCmd.Parameters.Add(parameter)
                Loop
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try
            Dim mStrSQL As StringBuilder = Me.mStrSQL
            mStrSQL.AppendLine(("INSERT INTO " & Me.mTabela))
            mStrSQL.AppendLine(("(" & builder.ToString & ")"))
            mStrSQL.AppendLine("VALUES ")
            mStrSQL.AppendLine(("(" & builder2.ToString & ")"))
            mStrSQL = Nothing
            Me.mCmd.CommandText = Me.mStrSQL.ToString
            Return Me.mCmd
        End Function

        Private Sub SetValueControle(ByVal Value As Object, ByRef Crt As Object, ByVal TType As DbType, ByRef evt As Object)
            Value = RuntimeHelpers.GetObjectValue(Interaction.IIf(Information.IsDBNull(RuntimeHelpers.GetObjectValue(Value)), String.Empty, RuntimeHelpers.GetObjectValue(Value)))
            Select Case Crt.GetType.ToString
                Case "System.Web.UI.WebControls.TextBox"
                    DirectCast(Crt, TextBox).Text = Conversions.ToString(Value)
                    Exit Select
                Case "System.Web.UI.WebControls.DropDownList"
                    DirectCast(Crt, DropDownList).SelectedValue = Conversions.ToString(Value)
                    Exit Select
                Case "System.Web.UI.WebControls.ListBox"
                    DirectCast(Crt, ListBox).SelectedValue = Conversions.ToString(Value)
                    Exit Select
                Case "System.Web.UI.WebControls.RadioButton"
                    DirectCast(Crt, RadioButton).Checked = Conversions.ToBoolean(Value)
                    Exit Select
                Case "System.Web.UI.WebControls.CheckBox"
                    DirectCast(Crt, CheckBox).Checked = Conversions.ToBoolean(Value)
                    Exit Select
                Case Else
                    If (Not evt Is Nothing) Then
                        Dim arguments As Object() = New Object() {RuntimeHelpers.GetObjectValue(Crt), RuntimeHelpers.GetObjectValue(Value)}
                        Dim copyBack As Boolean() = New Boolean() {True, True}
                        NewLateBinding.LateCall(evt, Nothing, "Invoke", arguments, Nothing, Nothing, copyBack, True)
                        If copyBack(0) Then
                            Crt = RuntimeHelpers.GetObjectValue(arguments(0))
                        End If
                        If copyBack(1) Then
                            Value = RuntimeHelpers.GetObjectValue(arguments(1))
                        End If
                    Else
                        Crt = RuntimeHelpers.GetObjectValue(Value)
                    End If
                    Exit Select
            End Select
        End Sub


        ' Properties
        Public ReadOnly Property Command() As DbCommand
            Get
                Return Me.mCmd
            End Get
        End Property

        Public Overridable Property PipeCommend() As String
            Get
                Return Conversions.ToString(Interaction.IIf((Me.mPipeCommend = ""), "?", Me.mPipeCommend))
            End Get
            Set(ByVal value As String)
                Me.mPipeCommend = value
            End Set
        End Property

        Public Property Tabela() As String
            Get
                Return Me.mTabela
            End Get
            Set(ByVal value As String)
                Me.mTabela = value
            End Set
        End Property


        ' Fields
        Private mCmd As DbCommand
        Private mColuna As Collection = New Collection
        Private mColunaWhere As Collection = New Collection
        Protected mPipeCommend As String = ""
        Public mStrSQL As StringBuilder = New StringBuilder
        Private mTabela As String = ""
        Private mTpSQL As TpSQL

        ' Nested Types
        Public Delegate Function EventoExterno(ByRef Objeto As Object, ByVal Valor As Object) As Object

        <StructLayout(LayoutKind.Sequential)> _
        Public Structure TpSQL
            Public Coluna As String
            Public Value As Object
            Public TType As DbType
            Public EventoExterno As EventoExterno
        End Structure
    End Class
End Namespace

