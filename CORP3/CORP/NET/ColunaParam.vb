Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Collections
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Web.UI.WebControls

Namespace CORP3.NET
    <Serializable()> _
    Public Class ColunaParam
        ' Methods
        Public Sub New()
            Dim hashtable As Hashtable = Me.HASH_TTipoColuna
            hashtable.Add(CShort(&H1B), TTipoColuna.CARACTER)
            hashtable.Add(CShort(&H1D), TTipoColuna.DATA)
            hashtable.Add(CShort(30), TTipoColuna.MONEY)
            hashtable.Add(CShort(&H1C), TTipoColuna.NUMERICO)
            hashtable = Nothing
            Me.HASH_TTipoLista = New Hashtable
            Dim hashtable2 As Hashtable = Me.HASH_TTipoLista
            hashtable2.Add(CShort(0), TTipoLista.UNDEFINED)
            hashtable2.Add(CShort(12), TTipoLista.LISTA)
            hashtable2.Add(CShort(11), TTipoLista.SQL)
            hashtable2 = Nothing
            Me.HASH_TTipoCampo = New Hashtable
            Dim hashtable3 As Hashtable = Me.HASH_TTipoCampo
            hashtable3.Add(CShort(&H19), TTipoCampo.COMBO)
            hashtable3.Add(CShort(&H16), TTipoCampo.TEXT)
            hashtable3.Add(CShort(&H1A), TTipoCampo.LISTA_MULTISELECT)
            hashtable3 = Nothing
            Me.HASH_Operadores = New Hashtable
            Dim hashtable4 As Hashtable = Me.HASH_Operadores
            hashtable4.Add(CShort(0), "")
            hashtable4.Add(CShort(13), " = ")
            hashtable4.Add(CShort(14), " > ")
            hashtable4.Add(CShort(15), " < ")
            hashtable4.Add(CShort(&H10), " >= ")
            hashtable4.Add(CShort(&H11), " <= ")
            hashtable4.Add(CShort(&H12), " Like '#%' ")
            hashtable4.Add(CShort(&H13), " Like '%#' ")
            hashtable4.Add(CShort(20), " Like '%#%' ")
            hashtable4.Add(CShort(&H15), " IN(#) ")
            hashtable4 = Nothing
            Me.HASH_Obrigatorio = New Hashtable
            Dim hashtable5 As Hashtable = Me.HASH_Operadores
            hashtable5.Add(CShort(&H1F), TObrigatorio.SIM)
            hashtable5.Add(CShort(&H20), TObrigatorio.NAO)
            hashtable5 = Nothing
        End Sub

        Protected Friend Function IncluirOperador() As String
            Select Case Me.mOperadorDefinido
                Case TOperador.A_PARTIR_DE, TOperador.TERMINADO_EM, TOperador.CONTENDO, TOperador.DENTRO_DE
                    Return Conversions.ToString(NewLateBinding.LateGet(Me.HASH_Operadores.Item(CShort(Me.mOperadorDefinido)), Nothing, "Replace", New Object() {"#", Me.Value.Trim}, Nothing, Nothing, Nothing))
            End Select
            Dim mOperadorDefinido As TOperador = Me.mOperadorDefinido
            If ((mOperadorDefinido = TOperador.MENOR_IGUAL) And (Me.mTipoColuna = TTipoColuna.DATA)) Then
                mOperadorDefinido = TOperador.MENOR
            End If
            Return Conversions.ToString(Operators.ConcatenateObject(Me.HASH_Operadores.Item(CShort(mOperadorDefinido)), Interaction.IIf((Me.mTipoCampo = TTipoCampo.TEXT), Me.ValorFormatado(False), Me.Value.Trim)))
        End Function

        Protected Friend Function IsValidParam() As Boolean
            Select Case Me.mTipoColuna
                Case TTipoColuna.CARACTER
                    If (Me.mValue.Trim = String.Empty) Then
                        Exit Select
                    End If
                    Return True
                Case TTipoColuna.NUMERICO, TTipoColuna.MONEY
                    If (Me.mTipoCampo <> TTipoCampo.TEXT) Then
                        If (Me.mValue.Trim <> String.Empty) Then
                            Return True
                        End If
                        Exit Select
                    End If
                    If Not Versioned.IsNumeric(Me.mValue) Then
                        Exit Select
                    End If
                    Return True
                Case TTipoColuna.DATA
                    If Not Information.IsDate(Me.mValue) Then
                        Exit Select
                    End If
                    Return True
            End Select
            Return False
        End Function

        Private Function TrataProp(ByVal ValProp As Short, ByRef hsh As Hashtable) As Integer
            If Not hsh.ContainsKey(ValProp) Then
                Throw New Exception("Valor informado é inválido.")
            End If
            Return ValProp
        End Function

        Public Function ValorFormatado(Optional ByVal IsProc As Boolean = False) As String
            Dim objectValue As Object = RuntimeHelpers.GetObjectValue(New Object)
            Select Case Me.mTipoColuna
                Case TTipoColuna.CARACTER
                    Return ("'" & Me.mValue.Trim & "'")
                Case TTipoColuna.NUMERICO, TTipoColuna.MONEY
                    Return Me.mValue.ToString.Replace(".", "").Replace(",", ".")
                Case TTipoColuna.DATA
                    If (Me.mValue.Trim = String.Empty) Then
                        Return "NULL"
                    End If
                    If Not ((Me.mOperadorDefinido = TOperador.MENOR_IGUAL) And Not IsProc) Then
                        objectValue = Conversions.ToDate(Me.mValue)
                        Exit Select
                    End If
                    objectValue = DateAndTime.DateAdd(DateInterval.Day, 1, Conversions.ToDate(Me.mValue))
                    Exit Select
                Case Else
                    Return Me.mValue
            End Select
            If (Me.mFormatoColuna.Trim = "") Then
                Return ("'" & Strings.Format(RuntimeHelpers.GetObjectValue(objectValue), "yyyy-MM-dd") & "'")
            End If
            Return ("'" & Strings.Format(RuntimeHelpers.GetObjectValue(objectValue), Me.mFormatoColuna) & "'")
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

        Public Property DicaFormatoColuna() As String
            Get
                Return Me.mTipFormatoColuna
            End Get
            Set(ByVal value As String)
                Me.mTipFormatoColuna = value
            End Set
        End Property

        Public Property FormatoColuna() As String
            Get
                Return Me.mFormatoColuna
            End Get
            Set(ByVal value As String)
                Me.mFormatoColuna = value
            End Set
        End Property

        Public Property Height() As Unit
            Get
                Return Me.mHeight
            End Get
            Set(ByVal value As Unit)
                Me.mHeight = value
            End Set
        End Property

        Public Property ListBoxLinhas() As Short
            Get
                Return Me.mListBoxLinhas
            End Get
            Set(ByVal value As Short)
                Me.mListBoxLinhas = value
            End Set
        End Property

        Public Property Nome() As String
            Get
                Return Me.mNome
            End Get
            Set(ByVal value As String)
                Me.mNome = value
            End Set
        End Property

        Public Property NomeColText() As String
            Get
                Return Me.mNomeColText
            End Get
            Set(ByVal value As String)
                Me.mNomeColText = value
            End Set
        End Property

        Public Property NomeColValue() As String
            Get
                Return Me.mNomeColValue
            End Get
            Set(ByVal value As String)
                Me.mNomeColValue = value
            End Set
        End Property

        Public Property Obrigatorio() As TObrigatorio
            Get
                Return Me.mObrigatorio
            End Get
            Set(ByVal value As TObrigatorio)
                Me.mObrigatorio = value
            End Set
        End Property

        Public Property Operador() As TOperador
            Get
                Return Me.mOperador
            End Get
            Set(ByVal value As TOperador)
                Me.mOperador = DirectCast(CShort(Me.TrataProp(CShort(value), Me.HASH_Operadores)), TOperador)
                Me.mOperadorDefinido = Me.mOperador
            End Set
        End Property

        Friend Property OperadorDefinido() As TOperador
            Get
                Return Me.mOperadorDefinido
            End Get
            Set(ByVal value As TOperador)
                Me.mOperadorDefinido = DirectCast(CShort(Me.TrataProp(CShort(value), Me.HASH_Operadores)), TOperador)
            End Set
        End Property

        Public Property Parametros() As String
            Get
                Return Me.mParametros
            End Get
            Set(ByVal value As String)
                Me.mParametros = value
            End Set
        End Property

        Public Property Provider() As ClsDB1.T_PROVIDER
            Get
                Return Me.mProvider
            End Get
            Set(ByVal value As ClsDB1.T_PROVIDER)
                Me.mProvider = value
            End Set
        End Property

        Public Property Separador() As String
            Get
                Return Me.mDivColParametros
            End Get
            Set(ByVal value As String)
                Me.mDivColParametros = value
            End Set
        End Property

        Public Property TipoCampo() As TTipoCampo
            Get
                Return Me.mTipoCampo
            End Get
            Set(ByVal value As TTipoCampo)
                Me.mTipoCampo = DirectCast(Me.TrataProp(CShort(value), Me.HASH_TTipoCampo), TTipoCampo)
            End Set
        End Property

        Public Property TipoColuna() As TTipoColuna
            Get
                Return Me.mTipoColuna
            End Get
            Set(ByVal value As TTipoColuna)
                Me.mTipoColuna = DirectCast(CShort(Me.TrataProp(CShort(value), Me.HASH_TTipoColuna)), TTipoColuna)
            End Set
        End Property

        Public Property TipoLista() As TTipoLista
            Get
                Return Me.mTipoLista
            End Get
            Set(ByVal value As TTipoLista)
                Me.mTipoLista = DirectCast(Me.TrataProp(CShort(value), Me.HASH_TTipoLista), TTipoLista)
            End Set
        End Property

        Public Property Titulo() As String
            Get
                Return Me.mTitulo
            End Get
            Set(ByVal value As String)
                Me.mTitulo = value
            End Set
        End Property

        Public Property ToolTip() As String
            Get
                Return Me.mToolTip
            End Get
            Set(ByVal value As String)
                Me.mToolTip = value
            End Set
        End Property

        Public Property ValorDelimitado() As Boolean
            Get
                Return Me.mValorDelimitado
            End Get
            Set(ByVal value As Boolean)
                Me.mValorDelimitado = value
            End Set
        End Property

        Public Property ValorMaximo() As String
            Get
                Return Me.mValorMaximo
            End Get
            Set(ByVal value As String)
                Me.mValorMaximo = value
            End Set
        End Property

        Public Property ValorMinimo() As String
            Get
                Return Me.mValorMinimo
            End Get
            Set(ByVal value As String)
                Me.mValorMinimo = value
            End Set
        End Property

        Public Property Value() As String
            Get
                If (Me.mValorDelimitado AndAlso ((Me.mValue <> String.Empty) And (Me.mTipoCampo = TTipoCampo.LISTA_MULTISELECT))) Then
                    Return ("'" & Me.mValue.Replace(",", "','") & "'")
                End If
                Return Me.mValue
            End Get
            Set(ByVal value As String)
                Me.mValue = value
            End Set
        End Property

        Public Property Width() As Unit
            Get
                Return Me.mWidth
            End Get
            Set(ByVal value As Unit)
                Me.mWidth = value
            End Set
        End Property


        ' Fields
        Private HASH_Obrigatorio As Hashtable
        Private HASH_Operadores As Hashtable
        Private HASH_TTipoCampo As Hashtable
        Private HASH_TTipoColuna As Hashtable = New Hashtable
        Private HASH_TTipoLista As Hashtable
        Private mConnectionString As String = String.Empty
        Private mDivColParametros As String = String.Empty
        Private mFormatoColuna As String = String.Empty
        Private mHeight As Unit = Unit.Pixel(10)
        Private mListBoxLinhas As Short = 5
        Private mNome As String = String.Empty
        Private mNomeColText As String = String.Empty
        Private mNomeColValue As String = String.Empty
        Private mObrigatorio As TObrigatorio = TObrigatorio.NAO
        Private mOperador As TOperador = TOperador.UNDEFINED
        Private mOperadorDefinido As TOperador = TOperador.UNDEFINED
        Private mParametros As String = String.Empty
        Private mProvider As ClsDB1.T_PROVIDER = ClsDB1.T_PROVIDER.OLEDB
        Private mTipFormatoColuna As String = String.Empty
        Private mTipoCampo As TTipoCampo = TTipoCampo.TEXT
        Private mTipoColuna As TTipoColuna = TTipoColuna.CARACTER
        Private mTipoLista As TTipoLista = TTipoLista.SQL
        Private mTitulo As String = String.Empty
        Private mToolTip As String = String.Empty
        Private mValorDelimitado As Boolean
        Private mValorMaximo As String = String.Empty
        Private mValorMinimo As String = String.Empty
        Private mValue As String = String.Empty
        Private mWidth As Unit = Unit.Pixel(70)

        ' Nested Types
        Public Enum TObrigatorio
            ' Fields
            NAO = &H20
            SIM = &H1F
        End Enum

        Public Enum TOperador As Short
            ' Fields
            A_PARTIR_DE = &H12
            CONTENDO = 20
            DENTRO_DE = &H15
            IGUAL = 13
            MAIOR = 14
            MAIOR_IGUAL = &H10
            MENOR = 15
            MENOR_IGUAL = &H11
            TERMINADO_EM = &H13
            UNDEFINED = 0
        End Enum

        Public Enum TTipoCampo
            ' Fields
            COMBO = &H19
            LISTA_MULTISELECT = &H1A
            [TEXT] = &H16
        End Enum

        Public Enum TTipoColuna As Short
            ' Fields
            CARACTER = &H1B
            DATA = &H1D
            MONEY = 30
            NUMERICO = &H1C
        End Enum

        Public Enum TTipoLista
            ' Fields
            LISTA = 12
            SQL = 11
            UNDEFINED = 0
        End Enum
    End Class
End Namespace

