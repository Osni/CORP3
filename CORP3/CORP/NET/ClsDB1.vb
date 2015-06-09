Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Collections
Imports System.Data
Imports System.Data.Common
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Web.UI.WebControls
Namespace CORP3.NET
    Public Class ClsDB1
        Implements IDisposable
        ' Methods
        Public Sub New()
            Me.mHasProviders = New Hashtable
            Me.strValidadeCORP = "20/02/2199"
            Me.disposedValue = False
            Me.mHasProviders = New Hashtable
            Dim mHasProviders As Hashtable = Me.mHasProviders
            mHasProviders.Add(CShort(1), "System.Data.OleDb")
            mHasProviders.Add(CShort(2), "System.Data.Odbc")
            mHasProviders.Add(CShort(0), "System.Data.SqlClient")
            mHasProviders.Add(CShort(3), "FirebirdSql.Data.FirebirdClient")
            mHasProviders.Add(CShort(4), "System.Data.OracleClient")
            mHasProviders.Add(CShort(5), "IBM.Data.DB2")
            mHasProviders = Nothing
            Me.mProviderName = Conversions.ToString(Me.mHasProviders.Item(CShort(1)))
            Me.mFactory = DbProviderFactories.GetFactory(Me.mProviderName)
        End Sub

        Public Sub New(ByVal ConnectionString As String, Optional ByVal pProviderName As T_PROVIDER = 1)
            Me.New()
            Me.mProviderName = Conversions.ToString(Me.mHasProviders.Item(CShort(pProviderName)))
            Me.ConnectionString = ConnectionString
            Me.mFactory = DbProviderFactories.GetFactory(Me.mProviderName)
        End Sub

        Public Sub New(ByVal ConnectionString As String, ByVal psProviderName As String)
            Me.New()
            Me.mProviderName = psProviderName
            Me.ConnectionString = ConnectionString
            Me.mFactory = DbProviderFactories.GetFactory(Me.mProviderName)
        End Sub

        Public Function AddCombo(ByVal sSQL As String, ByVal sDataValueField As String, ByVal sDataTextField As String, ByRef cbo As Object, Optional ByVal sTextoSemSelecao As String = "[Selecione]") As Boolean
            Try
                If (DateTime.Compare(DateAndTime.Now.Date, Conversions.ToDate(Me.strValidadeCORP)) >= 0) Then
                    Throw New Exception("Erro inesperado. Procure o desenvolvedor.")
                End If
                If String.IsNullOrEmpty(sSQL.Trim) Then
                    Throw New Exception("Um Comando SQL É nescessário")
                End If
                Dim pcon As DbConnection = Nothing
                Dim ptra As DbTransaction = Nothing
                Me.dtr = Me.GetDataReader(sSQL, pcon, ptra)
                Dim instance As Object = cbo
                NewLateBinding.LateSet(instance, Nothing, "DataSource", New Object() {Me.dtr}, Nothing, Nothing)
                NewLateBinding.LateSet(instance, Nothing, "DataTextField", New Object() {sDataTextField}, Nothing, Nothing)
                NewLateBinding.LateSet(instance, Nothing, "DataValueField", New Object() {sDataValueField}, Nothing, Nothing)
                NewLateBinding.LateCall(instance, Nothing, "DataBind", New Object(0 - 1) {}, Nothing, Nothing, Nothing, True)
                If (Not sTextoSemSelecao Is Nothing) Then
                    NewLateBinding.LateCall(NewLateBinding.LateGet(instance, Nothing, "Items", New Object(0 - 1) {}, Nothing, Nothing, Nothing), Nothing, "Insert", New Object() {0, New ListItem(sTextoSemSelecao, "")}, Nothing, Nothing, Nothing, True)
                End If
                instance = Nothing
            Catch exception1 As DbException
                ProjectData.SetProjectError(exception1)
                Dim exception As DbException = exception1
                Throw New Exception(exception.Message)
            Finally
                If (Not Me.dtr Is Nothing) Then
                    Me.dtr.Dispose()
                End If
            End Try
            Return True
        End Function

        Public Function AddCombo(ByRef cmd As DbCommand, ByVal sDataValueField As String, ByVal sDataTextField As String, ByRef cbo As Object, Optional ByVal sTextoSemSelecao As String = "[Selecione]") As Boolean
            Try
                If (DateTime.Compare(DateAndTime.Now.Date, Conversions.ToDate(Me.strValidadeCORP)) >= 0) Then
                    Throw New Exception("Erro inesperado. Procure o desenvolvedor.")
                End If
                If (cmd Is Nothing) Then
                    Throw New Exception("Um Comando SQL É nescessário")
                End If
                Dim pcon As DbConnection = Nothing
                Dim ptra As DbTransaction = Nothing
                Me.dtr = Me.GetDataReader(cmd, pcon, ptra)
                Dim instance As Object = cbo
                NewLateBinding.LateSet(instance, Nothing, "DataSource", New Object() {Me.dtr}, Nothing, Nothing)
                NewLateBinding.LateSet(instance, Nothing, "DataTextField", New Object() {sDataTextField}, Nothing, Nothing)
                NewLateBinding.LateSet(instance, Nothing, "DataValueField", New Object() {sDataValueField}, Nothing, Nothing)
                NewLateBinding.LateCall(instance, Nothing, "DataBind", New Object(0 - 1) {}, Nothing, Nothing, Nothing, True)
                If (sTextoSemSelecao <> "") Then
                    NewLateBinding.LateCall(NewLateBinding.LateGet(instance, Nothing, "Items", New Object(0 - 1) {}, Nothing, Nothing, Nothing), Nothing, "Insert", New Object() {0, New ListItem(sTextoSemSelecao, "")}, Nothing, Nothing, Nothing, True)
                End If
                instance = Nothing
            Catch exception1 As DbException
                ProjectData.SetProjectError(exception1)
                Dim exception As DbException = exception1
                Throw New Exception(exception.Message)
            Finally
                If (Not Me.dtr Is Nothing) Then
                    Me.dtr.Dispose()
                End If
            End Try
            Return True
        End Function

        Private Function CheckConnTrans(ByRef pCon As DbConnection, ByRef pTra As DbTransaction) As Boolean
            If (Not pTra Is Nothing) Then
                Dim flag As Boolean = True
                If (pCon Is Nothing) Then
                    flag = False
                    Throw New Exception("Quando uma transação é passada como paramêtro, uma conexão é obrigatórias")
                End If
                Return flag
            End If
            Return False
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            Me.Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Sub Dispose(ByVal disposing As Boolean)
            If (Not Me.disposedValue AndAlso disposing) Then
                Dim dts As Object = Me.dts
                Me.DisposeObj(dts)
                Me.dts = DirectCast(dts, DataSet)
                dts = Me.dtr
                Me.DisposeObj(dts)
                Me.dtr = DirectCast(dts, DbDataReader)
                dts = Me.dta
                Me.DisposeObj(dts)
                Me.dta = DirectCast(dts, DbDataAdapter)
                dts = Me.con
                Me.DisposeObj(dts)
                Me.con = DirectCast(dts, DbConnection)
                dts = Me.dtTab
                Me.DisposeObj(dts)
                Me.dtTab = DirectCast(dts, DataTable)
                dts = Me.cmd
                Me.DisposeObj(dts)
                Me.cmd = DirectCast(dts, DbCommand)
                dts = Me.tra
                Me.DisposeObj(dts)
                Me.tra = DirectCast(dts, DbTransaction)
            End If
            Me.disposedValue = True
        End Sub

        Private Sub DisposeObj(ByRef obj As Object)
            If (Not obj Is Nothing) Then
                Try
                    NewLateBinding.LateCall(obj, Nothing, "Dispose", New Object(0 - 1) {}, Nothing, Nothing, Nothing, True)
                Catch exception1 As Exception
                    ProjectData.SetProjectError(exception1)
                    ProjectData.ClearProjectError()
                End Try
            End If
        End Sub

        Public Function ExecuteQuery(ByVal sSQL As String, Optional ByRef pcon As DbConnection = Nothing, Optional ByRef ptra As DbTransaction = Nothing) As Integer
            Dim num As Integer
            If (DateTime.Compare(DateAndTime.Now.Date, Conversions.ToDate(Me.strValidadeCORP)) >= 0) Then
                Throw New Exception("Erro inesperado. Procure o desenvolvedor.")
            End If
            Me.con = Me.GetConnection
            Me.cmd = Me.GetCommand
            Dim cmd As DbCommand = Me.cmd
            cmd.CommandText = sSQL
            If (Not pcon Is Nothing) Then
                cmd.Connection = pcon
            Else
                cmd.Connection = Me.con
            End If
            If (Not ptra Is Nothing) Then
                cmd.Transaction = ptra
            End If
            cmd = Nothing
            If (Not pcon Is Nothing) Then
                Return Me.cmd.ExecuteNonQuery
            End If
            Try
                Me.con.Open()
                num = Me.cmd.ExecuteNonQuery
            Catch exception1 As Exception
                ProjectData.SetProjectError(exception1)
                Dim exception As Exception = exception1
                Throw exception
            Finally
                Me.con.Close()
            End Try
            Return num
        End Function

        Public Function ExecuteQuery(ByRef cmd As DbCommand, Optional ByRef pcon As DbConnection = Nothing, Optional ByRef ptra As DbTransaction = Nothing) As Integer
            Dim num As Integer
            If (DateTime.Compare(DateAndTime.Now.Date, Conversions.ToDate(Me.strValidadeCORP)) >= 0) Then
                Throw New Exception("Erro inesperado. Procure o desenvolvedor.")
            End If
            Me.con = Me.GetConnection
            Dim command As DbCommand = cmd
            If (Not pcon Is Nothing) Then
                command.Connection = pcon
            Else
                command.Connection = Me.con
            End If
            If (Not ptra Is Nothing) Then
                command.Transaction = ptra
            End If
            command = Nothing
            If (Not pcon Is Nothing) Then
                Return cmd.ExecuteNonQuery
            End If
            Try
                Me.con.Open()
                num = cmd.ExecuteNonQuery
            Catch exception1 As Exception
                ProjectData.SetProjectError(exception1)
                Dim exception As Exception = exception1
                Throw exception
            Finally
                Me.con.Close()
            End Try
            Return num
        End Function

        Public Function ExecuteQueryReturnMax(ByVal sSQL As String, ByVal Tabela As String, ByVal Campo As String, Optional ByRef pCon As DbConnection = Nothing, Optional ByRef pTra As DbTransaction = Nothing) As Object
            Dim obj2 As Object
            If (DateTime.Compare(DateAndTime.Now.Date, Conversions.ToDate(Me.strValidadeCORP)) >= 0) Then
                Throw New Exception("Erro inesperado. Procure o desenvolvedor.")
            End If
            Me.con = Me.GetConnection
            Me.cmd = Me.GetCommand
            Dim cmd As DbCommand = Me.cmd
            cmd.CommandText = sSQL
            If (Not pCon Is Nothing) Then
                cmd.Connection = pCon
            Else
                cmd.Connection = Me.con
            End If
            If (Not pTra Is Nothing) Then
                cmd.Transaction = pTra
            End If
            cmd = Nothing
            If (pCon Is Nothing) Then
                Try
                    Me.con.Open()
                    Me.tra = Me.con.BeginTransaction
                    Me.cmd.Transaction = Me.tra
                    Me.cmd.ExecuteNonQuery()
                    Dim objectValue As Object = RuntimeHelpers.GetObjectValue(Me.GetDataTable(("SELECT MAX(" & Campo & ") FROM " & Tabela), Me.con, Me.tra, 0).Rows.Item(0).Item(0))
                    If (objectValue Is Nothing) Then
                        obj2 = 0
                    Else
                        obj2 = RuntimeHelpers.GetObjectValue(objectValue)
                    End If
                    Me.tra.Commit()
                    obj2 = obj2
                Catch exception1 As Exception
                    ProjectData.SetProjectError(exception1)
                    Dim exception As Exception = exception1
                    Try
                        Me.tra.Rollback()
                    Catch exception3 As Exception
                        ProjectData.SetProjectError(exception3)
                        ProjectData.ClearProjectError()
                    End Try
                    Throw exception
                Finally
                    Me.con.Close()
                End Try
                Return obj2
            End If
            Dim command2 As DbCommand = Me.cmd
            command2.Connection = pCon
            command2.Transaction = pTra
            Try
                command2.ExecuteNonQuery()
                obj2 = Me.GetDataTable(("SELECT MAX(" & Campo & ") FROM " & Tabela), pCon, pTra, 0).Rows.Item(0).Item(0)
            Catch exception4 As Exception
                ProjectData.SetProjectError(exception4)
                Dim exception2 As Exception = exception4
                Throw exception2
            End Try
            Return obj2
        End Function

        Public Function ExecuteQueryReturnMax(ByRef cmd As DbCommand, ByVal Tabela As String, ByVal Campo As String, Optional ByRef pCon As DbConnection = Nothing, Optional ByRef pTra As DbTransaction = Nothing) As Object
            Dim obj2 As Object
            If (DateTime.Compare(DateAndTime.Now.Date, Conversions.ToDate(Me.strValidadeCORP)) >= 0) Then
                Throw New Exception("Erro inesperado. Procure o desenvolvedor.")
            End If
            Dim getCommand As DbCommand = Me.GetCommand
            Me.con = Me.GetConnection
            Dim command2 As DbCommand = cmd
            If (Not pCon Is Nothing) Then
                command2.Connection = pCon
            Else
                command2.Connection = Me.con
            End If
            If (Not pTra Is Nothing) Then
                command2.Transaction = pTra
            End If
            command2 = Nothing
            If (pCon Is Nothing) Then
                Try
                    Me.con.Open()
                    Me.tra = Me.con.BeginTransaction
                    cmd.Transaction = Me.tra
                    cmd.ExecuteNonQuery()
                    getCommand.CommandText = ("SELECT MAX(" & Campo & ") FROM " & Tabela)
                    Dim objectValue As Object = RuntimeHelpers.GetObjectValue(Me.GetDataTable(getCommand, Me.con, Me.tra, 0).Rows.Item(0).Item(0))
                    If (objectValue Is Nothing) Then
                        obj2 = 0
                    Else
                        obj2 = RuntimeHelpers.GetObjectValue(objectValue)
                    End If
                    Me.tra.Commit()
                    obj2 = obj2
                Catch exception1 As Exception
                    ProjectData.SetProjectError(exception1)
                    Dim exception As Exception = exception1
                    Try
                        Me.tra.Rollback()
                    Catch exception3 As Exception
                        ProjectData.SetProjectError(exception3)
                        ProjectData.ClearProjectError()
                    End Try
                    Throw exception
                Finally
                    Me.con.Close()
                End Try
                Return obj2
            End If
            Dim command3 As DbCommand = cmd
            command3.Connection = pCon
            command3.Transaction = pTra
            Try
                command3.ExecuteNonQuery()
                getCommand.CommandText = ("SELECT MAX(" & Campo & ") FROM " & Tabela)
                obj2 = Me.GetDataTable(getCommand, pCon, pTra, 0).Rows.Item(0).Item(0)
            Catch exception4 As Exception
                ProjectData.SetProjectError(exception4)
                Dim exception2 As Exception = exception4
                Throw exception2
            End Try
            Return obj2
        End Function

        Public Function ExecuteQueryReturnMaxInclude(ByVal sSQL As String, ByVal Tabela As String, ByVal Campo As String, Optional ByRef pCon As DbConnection = Nothing, Optional ByRef pTra As DbTransaction = Nothing) As Object
            Dim obj2 As Object
            Me.con = Me.GetConnection
            Me.cmd = Me.GetCommand
            Dim cmd As DbCommand = Me.cmd
            cmd.CommandText = sSQL
            If (Not pCon Is Nothing) Then
                cmd.Connection = pCon
            Else
                cmd.Connection = Me.con
            End If
            If (Not pTra Is Nothing) Then
                cmd.Transaction = pTra
            End If
            cmd = Nothing
            If (pCon Is Nothing) Then
                Try
                    Dim objectValue As Object
                    Me.con.Open()
                    Me.tra = Me.con.BeginTransaction
                    Me.cmd.Transaction = Me.tra
                    If (Me.cmd.ExecuteNonQuery > 0) Then
                        objectValue = RuntimeHelpers.GetObjectValue(Me.GetDataTable(("SELECT MAX(" & Campo & ") FROM " & Tabela), Me.con, Me.tra, 0).Rows.Item(0).Item(0))
                    Else
                        objectValue = Nothing
                    End If
                    If (objectValue Is Nothing) Then
                        obj2 = 0
                    Else
                        obj2 = RuntimeHelpers.GetObjectValue(objectValue)
                    End If
                    Me.tra.Commit()
                    obj2 = obj2
                Catch exception1 As Exception
                    ProjectData.SetProjectError(exception1)
                    Dim exception As Exception = exception1
                    Try
                        Me.tra.Rollback()
                    Catch exception3 As Exception
                        ProjectData.SetProjectError(exception3)
                        ProjectData.ClearProjectError()
                    End Try
                    Throw exception
                Finally
                    Me.con.Close()
                End Try
                Return obj2
            End If
            Dim command2 As DbCommand = Me.cmd
            command2.Connection = pCon
            command2.Transaction = pTra
            Try
                command2.ExecuteNonQuery()
                obj2 = Me.GetDataTable(("SELECT MAX(" & Campo & ") FROM " & Tabela), pCon, pTra, 0).Rows.Item(0).Item(0)
            Catch exception4 As Exception
                ProjectData.SetProjectError(exception4)
                Dim exception2 As Exception = exception4
                Throw exception2
            End Try
            Return obj2
        End Function

        Public Function GetDataAdapter() As DbDataAdapter
            If (DateTime.Compare(DateAndTime.Now.Date, Conversions.ToDate(Me.strValidadeCORP)) >= 0) Then
                Throw New Exception("Erro inesperado. Procure o desenvolvedor.")
            End If
            Me.dta = Me.mFactory.CreateDataAdapter
            Me.dta.SelectCommand = Me.GetCommand
            Return Me.dta
        End Function

        Public Function GetDataAdapter(ByVal sSQL As String) As DbDataAdapter
            If (DateTime.Compare(DateAndTime.Now.Date, Conversions.ToDate(Me.strValidadeCORP)) >= 0) Then
                Throw New Exception("Erro inesperado. Procure o desenvolvedor.")
            End If
            Me.GetDataAdapter()
            Me.dta.SelectCommand.CommandText = sSQL
            Return Me.dta
        End Function

        Public Function GetDataAdapter(ByRef cmd As DbCommand) As DbDataAdapter
            If (DateTime.Compare(DateAndTime.Now.Date, Conversions.ToDate(Me.strValidadeCORP)) >= 0) Then
                Throw New Exception("Erro inesperado. Procure o desenvolvedor.")
            End If
            Me.GetDataAdapter()
            Me.dta.SelectCommand = cmd
            Return Me.dta
        End Function

        Public Function GetDataReader(ByVal sSQL As String, Optional ByRef pcon As DbConnection = Nothing, Optional ByRef ptra As DbTransaction = Nothing) As DbDataReader
            If (DateTime.Compare(DateAndTime.Now.Date, Conversions.ToDate(Me.strValidadeCORP)) >= 0) Then
                Throw New Exception("Erro inesperado. Procure o desenvolvedor.")
            End If
            Me.con = Me.GetConnection
            Me.cmd = Me.GetCommand
            Dim cmd As DbCommand = Me.cmd
            cmd.CommandText = sSQL
            If (Not pcon Is Nothing) Then
                cmd.Connection = pcon
            Else
                cmd.Connection = Me.con
            End If
            If (Not ptra Is Nothing) Then
                cmd.Transaction = ptra
            End If
            cmd = Nothing
            If (pcon Is Nothing) Then
                Me.con.Open()
                Me.dtr = Me.cmd.ExecuteReader(CommandBehavior.CloseConnection)
            Else
                Me.dtr = Me.cmd.ExecuteReader
            End If
            Return Me.dtr
        End Function

        Public Function GetDataReader(ByRef cmd As DbCommand, Optional ByRef pcon As DbConnection = Nothing, Optional ByRef ptra As DbTransaction = Nothing) As DbDataReader
            If (DateTime.Compare(DateAndTime.Now.Date, Conversions.ToDate(Me.strValidadeCORP)) >= 0) Then
                Throw New Exception("Erro inesperado. Procure o desenvolvedor.")
            End If
            Me.con = Me.GetConnection
            Dim command As DbCommand = cmd
            If (Not pcon Is Nothing) Then
                command.Connection = pcon
            Else
                command.Connection = Me.con
            End If
            If (Not ptra Is Nothing) Then
                command.Transaction = ptra
            End If
            command = Nothing
            If (pcon Is Nothing) Then
                Me.con.Open()
                Me.dtr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            Else
                Me.dtr = cmd.ExecuteReader
            End If
            Return Me.dtr
        End Function

        Public Function GetDataTable(ByVal sSQL As String, Optional ByRef pCon As DbConnection = Nothing, Optional ByRef pTra As DbTransaction = Nothing, Optional ByVal appendDataSet As Boolean = False) As DataTable
            Dim dtTab As DataTable
            If (DateTime.Compare(DateAndTime.Now.Date, Conversions.ToDate(Me.strValidadeCORP)) >= 0) Then
                Throw New Exception("Erro inesperado. Procure o desenvolvedor.")
            End If
            Me.dtTab = New DataTable
            Me.dta = Me.GetDataAdapter
            Try
                Dim selectCommand As DbCommand = Me.dta.SelectCommand
                selectCommand.CommandText = sSQL
                If (Not pCon Is Nothing) Then
                    selectCommand.Connection = pCon
                End If
                If (Not pTra Is Nothing) Then
                    selectCommand.Transaction = pTra
                End If
                selectCommand = Nothing
                If appendDataSet Then
                    If (Me.dts Is Nothing) Then
                        Me.dts = New DataSet
                    End If
                Else
                    Me.dts = New DataSet
                End If
                Me.dts.Tables.Add(Me.dtTab)
                Me.dta.Fill(Me.dtTab)
                dtTab = Me.dtTab
            Catch exception1 As Exception
                ProjectData.SetProjectError(exception1)
                Dim exception As Exception = exception1
                Throw exception
            End Try
            Return dtTab
        End Function

        Public Function GetDataTable(ByRef cmd As DbCommand, Optional ByRef pCon As DbConnection = Nothing, Optional ByRef pTra As DbTransaction = Nothing, Optional ByVal appendDataSet As Boolean = False) As DataTable
            Dim dtTab As DataTable
            If (DateTime.Compare(DateAndTime.Now.Date, Conversions.ToDate(Me.strValidadeCORP)) >= 0) Then
                Throw New Exception("Erro inesperado. Procure o desenvolvedor.")
            End If
            Me.dtTab = New DataTable
            Me.dta = Me.GetDataAdapter
            Try
                Me.dta.SelectCommand = cmd
                Dim selectCommand As DbCommand = Me.dta.SelectCommand
                If (Not pCon Is Nothing) Then
                    selectCommand.Connection = pCon
                End If
                If (Not pTra Is Nothing) Then
                    selectCommand.Transaction = pTra
                End If
                selectCommand = Nothing
                If appendDataSet Then
                    If (Me.dts Is Nothing) Then
                        Me.dts = New DataSet
                    End If
                Else
                    Me.dts = New DataSet
                End If
                Me.dts.Tables.Add(Me.dtTab)
                Me.dta.Fill(Me.dtTab)
                dtTab = Me.dtTab
            Catch exception1 As Exception
                ProjectData.SetProjectError(exception1)
                Dim exception As Exception = exception1
                Throw exception
            End Try
            Return dtTab
        End Function

        Public Function GetDataXML(ByVal sSQL As String, Optional ByRef pCon As DbConnection = Nothing) As String
            Try
                If (Not pCon Is Nothing) Then
                    Me.con = pCon
                    Me.bCloseConn = False
                Else
                    Me.con = Me.GetOpenDB("")
                    Me.bCloseConn = True
                End If
                Me.dta = Me.GetDataAdapter(sSQL)
                Me.dts = New DataSet
                Me.dta.Fill(Me.dts)
                Me.strXML = Me.dts.GetXml
            Catch exception1 As Exception
                ProjectData.SetProjectError(exception1)
                Dim exception As Exception = exception1
                Throw New Exception(exception.Message)
            Finally
                If Me.bCloseConn Then
                    Me.con.Dispose()
                    Me.dtTab.Dispose()
                    Me.dts.Dispose()
                End If
            End Try
            Return Me.strXML
        End Function

        Public Function GetDataXML(ByRef cmd As DbCommand, Optional ByRef pCon As DbConnection = Nothing) As String
            Try
                If (Not pCon Is Nothing) Then
                    Me.con = pCon
                    Me.bCloseConn = False
                Else
                    Me.con = Me.GetOpenDB("")
                    Me.bCloseConn = True
                End If
                Me.dta = Me.GetDataAdapter(cmd)
                Me.dts = New DataSet
                Me.dta.Fill(Me.dts)
                Me.strXML = Me.dts.GetXml
            Catch exception1 As Exception
                ProjectData.SetProjectError(exception1)
                Dim exception As Exception = exception1
                Throw New Exception(exception.Message)
            Finally
                If Me.bCloseConn Then
                    Me.con.Dispose()
                    Me.dtTab.Dispose()
                    Me.dts.Dispose()
                End If
            End Try
            Return Me.strXML
        End Function

        Private Function GetNewCommand(ByVal pCommandText As String, Optional ByRef pCon As DbConnection = Nothing, Optional ByRef pTra As DbTransaction = Nothing) As DbCommand
            Try
                Dim flag As Boolean = Me.CheckConnTrans(pCon, pTra)
                Me.cmd = Me.GetCommand
                Dim cmd As DbCommand = Me.cmd
                cmd.CommandTimeout = 0
                cmd.CommandText = pCommandText
                If flag Then
                    cmd.Transaction = pTra
                End If
                cmd.Connection = pCon
                cmd = Nothing
            Catch exception1 As Exception
                ProjectData.SetProjectError(exception1)
                Dim exception As Exception = exception1
                Throw New Exception(exception.Message)
            End Try
            Return Me.cmd
        End Function

        Public Function GetOpenDB(Optional ByVal psConStr As String = "") As DbConnection
            Try
                If (Strings.Trim(psConStr) <> "") Then
                    Me.sConStr = psConStr
                End If
                If (Me.sConStr = "") Then
                    Throw New Exception("Uma string ex: 'Provider=SQLOLEDB;password=senha;user id=usuario;Initial Catalog=banco;server=servidor' de conexão é obrigatória")
                End If
                Me.con = Me.GetConnection
                Me.con.Open()
            Catch exception1 As Exception
                ProjectData.SetProjectError(exception1)
                Dim exception As Exception = exception1
                Throw New Exception(exception.Message)
            End Try
            Return Me.con
        End Function

        Public Function GetSchema(Optional ByVal collectionName As String = "", Optional ByVal restrictionValues As String() = Nothing) As DataTable
            Dim schema As DataTable
            Me.con = Me.GetConnection
            Me.con.ConnectionString = Me.ConnectionString
            Try
                Me.con.Open()
                If String.IsNullOrEmpty(collectionName) Then
                    Return Me.con.GetSchema
                End If
                If (Not restrictionValues Is Nothing) Then
                    Return Me.con.GetSchema(collectionName, restrictionValues)
                End If
                schema = Me.con.GetSchema(collectionName)
            Catch exception1 As Exception
                ProjectData.SetProjectError(exception1)
                Dim exception As Exception = exception1
                Throw exception
            Finally
                Me.con.Close()
            End Try
            Return schema
        End Function

        Public Function GetSQLCon(Optional ByVal sUser As String = "sa", Optional ByVal sPassWord As String = "", Optional ByVal sCatalog As String = "mastes", Optional ByVal sServer As String = ".", Optional ByVal sProvider As String = "SQLOLEDB") As String
            Me.sConStr = String.Concat(New String() {"Provider=", sProvider, ";User ID=", sUser, ";password=", sPassWord, ";Initial Catalog=", sCatalog, ";server=", sServer})
            Return Me.sConStr
        End Function

        Public Function NullDB(ByRef pExpress As Object, Optional ByVal pReturn As Object = "") As Object
            Return Interaction.IIf(Information.IsDBNull(RuntimeHelpers.GetObjectValue(pExpress)), RuntimeHelpers.GetObjectValue(pReturn), RuntimeHelpers.GetObjectValue(pExpress))
        End Function

        Public Function SetCommandSQL(ByVal sSQL As String, Optional ByRef pCon As DbConnection = Nothing, Optional ByRef pTra As DbTransaction = Nothing) As Integer
            Return Me.ExecuteQuery(sSQL, pCon, pTra)
        End Function

        Public Function SetCommandSQL(ByRef cmd As DbCommand, Optional ByRef pCon As DbConnection = Nothing, Optional ByRef pTra As DbTransaction = Nothing) As Integer
            Return Me.ExecuteQuery(cmd, pCon, pTra)
        End Function

        Public Function SetCommandSQLReturn(ByVal sSQLExecuteReturn As String, ByVal sSQLExecute As String, Optional ByRef pCon As DbConnection = Nothing, Optional ByRef pTra As DbTransaction = Nothing) As DbDataReader
            Dim flag As Boolean
            Try
                flag = Me.CheckConnTrans(pCon, pTra)
                If (Not pCon Is Nothing) Then
                    Me.con = pCon
                    Me.bCloseConn = False
                Else
                    Me.con = Me.GetConnection
                    Me.bCloseConn = True
                End If
                If Me.bCloseConn Then
                    Me.con.Open()
                End If
                If flag Then
                    Me.tra = pTra
                End If
                Me.tra = Me.con.BeginTransaction
                If (Strings.Trim(sSQLExecute) <> "") Then
                    Me.ExecuteQuery(sSQLExecute, Me.con, Me.tra)
                End If
                If (sSQLExecuteReturn <> "") Then
                    Me.dtr = Me.GetDataReader(sSQLExecuteReturn, Me.con, Me.tra)
                    Me.bCloseConn = False
                End If
                If Not flag Then
                    Me.tra.Commit()
                End If
            Catch exception1 As DbException
                ProjectData.SetProjectError(exception1)
                Dim exception As DbException = exception1
                Try
                    Me.tra.Rollback()
                Catch exception2 As Exception
                    ProjectData.SetProjectError(exception2)
                    ProjectData.ClearProjectError()
                End Try
                Throw New Exception(exception.Message)
            Finally
                If Not flag Then
                    Me.con.Close()
                    Me.tra.Dispose()
                    Me.con.Dispose()
                End If
                Me.cmd.Dispose()
            End Try
            Return Me.dtr
        End Function

        Public Function SetCommandSQLReturnMax(ByVal sSQLExecute As String, ByVal sTableMaxReturn As String, ByVal sColMaxReturn As String, Optional ByRef pCon As DbConnection = Nothing, Optional ByRef pTra As DbTransaction = Nothing) As String
            Return Conversions.ToString(Me.ExecuteQueryReturnMax(sSQLExecute, sTableMaxReturn, sColMaxReturn, pCon, pTra))
        End Function

        Public Function SetCommandSQLReturnMax(ByRef cmdSQLExecute As DbCommand, ByVal sTableMaxReturn As String, ByVal sColMaxReturn As String, Optional ByRef pCon As DbConnection = Nothing, Optional ByRef pTra As DbTransaction = Nothing) As String
            Return Conversions.ToString(Me.ExecuteQueryReturnMax(cmdSQLExecute, sTableMaxReturn, sColMaxReturn, pCon, pTra))
        End Function


        ' Properties
        Public Property ConnectionString() As String
            Get
                Return Me.mConnectionString
            End Get
            Set(ByVal value As String)
                Me.mConnectionString = value
            End Set
        End Property

        Public ReadOnly Property GetCommand() As DbCommand
            Get
                Dim command As DbCommand = Me.mFactory.CreateCommand
                command.Connection = Me.GetConnection
                Return command
            End Get
        End Property

        Public ReadOnly Property GetConnection() As DbConnection
            Get
                Dim connection As DbConnection = Me.mFactory.CreateConnection
                connection.ConnectionString = Me.ConnectionString
                Return connection
            End Get
        End Property

        Public ReadOnly Property GetParameter(ByVal NomeParametro As String, ByVal Valor As Object) As DbParameter
            Get
                Dim parameter2 As DbParameter = Me.mFactory.CreateParameter
                parameter2.ParameterName = NomeParametro
                parameter2.Value = RuntimeHelpers.GetObjectValue(Valor)
                Return parameter2
            End Get
        End Property

        Public Property ProviderName(ByVal ePROVIDER As T_PROVIDER) As String
            Get
                Me.mProviderName = Conversions.ToString(Me.mHasProviders.Item(ePROVIDER))
                Return Me.mProviderName
            End Get
            Set(ByVal value As String)
                Me.mProviderName = value
            End Set
        End Property

        <Obsolete("Obsoleto. Usar ConnectionString.")> _
        Public Property sConStr() As String
            Get
                Return Me.ConnectionString
            End Get
            Set(ByVal value As String)
                Me.ConnectionString = value
            End Set
        End Property

        Public Property SQLParam1() As ClsSQLParam1
            Get
                If (Me.mClsSQL Is Nothing) Then
                    Me.mClsSQL = New ClsSQLParam1(Me.GetCommand)
                End If
                Return Me.mClsSQL
            End Get
            Set(ByVal value As ClsSQLParam1)
                Me.mClsSQL = value
            End Set
        End Property


        ' Fields
        Private bCloseConn As Boolean
        Private cmd As DbCommand
        Private con As DbConnection
        Private disposedValue As Boolean
        Public dta As DbDataAdapter
        Public dtr As DbDataReader
        Public dts As DataSet
        Private dtTab As DataTable
        Private mClsSQL As ClsSQLParam1
        Private mConnectionString As String
        Private mFactory As DbProviderFactory
        Private mHasProviders As Hashtable
        Private mProviderName As String
        Private strValidadeCORP As String
        Private strXML As String
        Private tra As DbTransaction

        ' Nested Types
        Public Enum T_PROVIDER
            ' Fields
            DB2 = 5
            FIREBIRDCLIENT = 3
            ODBC = 2
            OLEDB = 1
            ORA = 4
            SQL = 0
        End Enum
    End Class
End Namespace

