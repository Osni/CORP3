Imports System
Imports System.Collections
Imports System.Web.UI.WebControls

Namespace CORP3.NET
    Public Class Coluna
        ' Methods
        Public Sub New()
            Dim hashtable As Hashtable = Me.HASH_TTipoColuna
            hashtable.Add(CShort(14), TTipoColuna.DETALHE)
            hashtable.Add(CShort(13), TTipoColuna.GRUPO)
            hashtable = Nothing
            Me.HASH_TAlinhamento = New Hashtable
            Dim hashtable2 As Hashtable = Me.HASH_TAlinhamento
            hashtable2.Add(CShort(2), "center")
            hashtable2.Add(CShort(1), "left")
            hashtable2.Add(CShort(7), "right")
            hashtable2 = Nothing
            Me.HASH_TTipoResumo = New Hashtable
            Dim hashtable3 As Hashtable = Me.HASH_TTipoResumo
            hashtable3.Add(CShort(8), TTipoResumo.SEM_RESUMO)
            hashtable3.Add(CShort(9), TTipoResumo.SOMA)
            hashtable3.Add(CShort(10), TTipoResumo.RESUMO)
            hashtable3 = Nothing
            Me.HASH_TTipoDado = New Hashtable
            Dim hashtable4 As Hashtable = Me.HASH_TTipoResumo
            hashtable4.Add(CShort(&H1B), TTipoDado.CARACTER)
            hashtable4.Add(CShort(&H1D), TTipoDado.DATA)
            hashtable4.Add(CShort(30), TTipoDado.MONEY)
            hashtable4.Add(CShort(&H1C), TTipoDado.NUMERICO)
            hashtable4 = Nothing
        End Sub

        Friend Sub LimpaResumoParcial()
            Me.mResumoSubTotal = 0
        End Sub

        Friend Sub LimpaTodoResumo()
            Me.mResumoSubTotal = 0
            Me.mResumoTotal = 0
        End Sub

        Private Function TrataProp(ByVal ValProp As Short, ByRef hsh As Hashtable) As Integer
            If Not hsh.ContainsKey(ValProp) Then
                Throw New Exception("Valor informado é inválido.")
            End If
            Return ValProp
        End Function


        ' Properties
        Public Property Alinhamento() As TAlinhamento
            Get
                If (Me.mAlinhamento = DirectCast(CShort(Me.TrataProp(CShort(0), Me.HASH_TAlinhamento)), TAlinhamento)) Then
                    Me.mAlinhamento = TAlinhamento.LEFT
                End If
                Return Me.mAlinhamento
            End Get
            Set(ByVal value As TAlinhamento)
                Me.mAlinhamento = DirectCast(CShort(Me.TrataProp(CShort(value), Me.HASH_TAlinhamento)), TAlinhamento)
            End Set
        End Property

        Public Property ColumnMaxLength() As Integer
            Get
                Return Me.mColumnMaxLength
            End Get
            Set(ByVal value As Integer)
                Me.mColumnMaxLength = CShort(value)
            End Set
        End Property

        Public Property ColumnSize() As Unit
            Get
                Return Me.mColumnSize
            End Get
            Set(ByVal value As Unit)
                Me.mColumnSize = value
            End Set
        End Property

        Public Property Formato() As String
            Get
                Return Me.mFormato
            End Get
            Set(ByVal value As String)
                Me.mFormato = value
            End Set
        End Property

        Public Property HeaderStyle() As Hashtable
            Get
                If (Me.colHeaderStyle Is Nothing) Then
                    Me.colHeaderStyle = New Hashtable
                End If
                Return Me.colHeaderStyle
            End Get
            Set(ByVal value As Hashtable)
                Me.colHeaderStyle = value
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

        Public Property QuebrarTexto() As TQuebrarTexto
            Get
                Return Me.mQuebrarTexto
            End Get
            Set(ByVal value As TQuebrarTexto)
                Me.mQuebrarTexto = value
            End Set
        End Property

        Friend Property ResumoSubTotal() As Double
            Get
                Return Me.mResumoSubTotal
            End Get
            Set(ByVal value As Double)
                Me.mResumoSubTotal = (Me.mResumoSubTotal + value)
                Me.mResumoTotal = (Me.mResumoTotal + value)
            End Set
        End Property

        Friend Property ResumoTotal() As Double
            Get
                Return Me.mResumoTotal
            End Get
            Set(ByVal value As Double)
                Me.mResumoTotal = value
            End Set
        End Property

        Public Property RotuloResumoFinal() As String
            Get
                Return Me.mRotuloResumoFinal
            End Get
            Set(ByVal value As String)
                Me.mRotuloResumoFinal = value
            End Set
        End Property

        Public Property RotuloResumoPag() As String
            Get
                Return Me.mRotuloResumoPag
            End Get
            Set(ByVal value As String)
                Me.mRotuloResumoPag = value
            End Set
        End Property

        Public Property TipoColuna() As TTipoColuna
            Get
                If (Me.mTipoColuna = 0) Then
                    Me.mTipoColuna = 14
                End If
                Return DirectCast(Me.mTipoColuna, TTipoColuna)
            End Get
            Set(ByVal value As TTipoColuna)
                Me.mTipoColuna = CShort(Me.TrataProp(CShort(value), Me.HASH_TTipoColuna))
            End Set
        End Property

        Public Property TipoDado() As TTipoDado
            Get
                Return Me.mTipoDado
            End Get
            Set(ByVal value As TTipoDado)
                Me.mTipoDado = value
            End Set
        End Property

        Public Property TipoResumo() As TTipoResumo
            Get
                If (Me.mTipoResumo = 0) Then
                    Me.mTipoResumo = 8
                End If
                Return DirectCast(Me.mTipoResumo, TTipoResumo)
            End Get
            Set(ByVal value As TTipoResumo)
                Me.mTipoResumo = CShort(Me.TrataProp(CShort(value), Me.HASH_TTipoResumo))
            End Set
        End Property

        Public Property Titulo() As String
            Get
                Return Me.mTitulo.Replace(" ", "&nbsp;")
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


        ' Fields
        Private colHeaderStyle As Hashtable
        Private HASH_TAlinhamento As Hashtable
        Private HASH_TTipoColuna As Hashtable = New Hashtable
        Private HASH_TTipoDado As Hashtable
        Private HASH_TTipoResumo As Hashtable
        Private mAlinhamento As TAlinhamento
        Private mColumnMaxLength As Short = 0
        Private mColumnSize As Unit
        Private mFormato As String = String.Empty
        Private mNome As String = String.Empty
        Private mNumLinhasGrupo As Short = 0
        Private mQuebrarTexto As TQuebrarTexto = TQuebrarTexto.NAO
        Private mResumoSubTotal As Double = 0
        Private mResumoTotal As Double = 0
        Private mRotuloResumoFinal As String = String.Empty
        Private mRotuloResumoPag As String = String.Empty
        Private mTipoColuna As Short = 0
        Private mTipoDado As TTipoDado
        Private mTipoResumo As Short = 0
        Private mTitulo As String = String.Empty
        Private mToolTip As String = String.Empty

        ' Nested Types
        Public Enum TAlinhamento As Short
            ' Fields
            CENTER = 2
            LEFT = 1
            RIGHT = 7
        End Enum

        Public Enum TQuebrarTexto As Short
            ' Fields
            NAO = &H24
            SIM = &H23
        End Enum

        Public Enum TTipoColuna As Short
            ' Fields
            DETALHE = 14
            GRUPO = 13
        End Enum

        Public Enum TTipoDado As Short
            ' Fields
            CARACTER = &H1B
            DATA = &H1D
            MONEY = 30
            NUMERICO = &H1C
        End Enum

        Public Enum TTipoResumo As Short
            ' Fields
            RESUMO = 10
            SEM_RESUMO = 8
            SOMA = 9
        End Enum
    End Class
End Namespace

