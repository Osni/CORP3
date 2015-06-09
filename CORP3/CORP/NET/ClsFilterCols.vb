Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Data
Imports System.Runtime.InteropServices

Namespace CORP3.NET
    Public Class ClsFilterCols
        Inherits PropertyCollection
        ' Methods
        Public Sub New(Optional ByVal sName As String = "", Optional ByVal sLabel As String = "", Optional ByVal sTitle As String = "", Optional ByVal bVisible As Boolean = True, Optional ByVal iSize As Integer = 10, Optional ByVal eTypeDB As TypeDB = 1, Optional ByVal sTextValue As String = "", Optional ByVal sPageURLDestino As String = "", Optional ByVal sPageURLColVar As String = "", Optional ByVal Style As TStyle = 0)
            Try
                If (sName <> "") Then
                    Me.prt = New PropertyCollection
                    Dim prt As PropertyCollection = Me.prt
                    prt.Add("Name", sName)
                    prt.Add("Label", sLabel)
                    prt.Add("Title", sTitle)
                    prt.Add("Visible", bVisible)
                    prt.Add("Size", iSize)
                    prt.Add("Type", Conversion.Int(CInt(eTypeDB)))
                    prt.Add("TextValue", sTextValue)
                    prt.Add("PageURLDestino", sPageURLDestino)
                    prt.Add("PageURLColVar", sPageURLColVar)
                    prt.Add("Style", Style)
                    prt = Nothing
                End If
            Catch exception1 As Exception
                ProjectData.SetProjectError(exception1)
                Dim exception As Exception = exception1
                Throw New Exception(exception.Message)
            End Try
        End Sub

        Public Function GetFilterCols() As PropertyCollection
            Dim propertys As PropertyCollection
            Try
                Me.prt = New PropertyCollection
                Dim prt As PropertyCollection = Me.prt
                prt.Add("Label", Me._Label)
                prt.Add("Name", Me._Name)
                prt.Add("Title", Me._Title)
                prt.Add("Visible", Me._Visible)
                prt.Add("Size", Me._Size)
                prt.Add("Type", Conversion.Int(Me._TypeDB))
                prt.Add("TextValue", Me._TextValue)
                prt.Add("PageURLDestino", Me._PageURLDestino)
                prt.Add("PageURLColVar", Me._PageURLColVar)
                prt.Add("Style", Me._Style)
                prt = Nothing
                propertys = Me.prt
            Catch exception1 As Exception
                ProjectData.SetProjectError(exception1)
                Dim exception As Exception = exception1
                Throw New Exception(exception.Message)
            End Try
            Return propertys
        End Function


        ' Properties
        Public ReadOnly Property FilterColsReadOnly() As PropertyCollection
            Get
                Return Me.prt
            End Get
        End Property

        Public Property Label() As String
            Get
                Return Me._Label
            End Get
            Set(ByVal value As String)
                Me._Label = value
            End Set
        End Property

        Public Property Name() As String
            Get
                Return Me._Name
            End Get
            Set(ByVal value As String)
                Me._Name = value
            End Set
        End Property

        Public Property PageURLColVar() As String
            Get
                Return Me._PageURLColVar
            End Get
            Set(ByVal value As String)
                Me._PageURLColVar = value
            End Set
        End Property

        Public Property PageURLDestino() As String
            Get
                Return Me._PageURLDestino
            End Get
            Set(ByVal value As String)
                Me._PageURLDestino = value
            End Set
        End Property

        Public Property Size() As String
            Get
                Return Conversions.ToString(Me._Size)
            End Get
            Set(ByVal value As String)
                Me._Size = Conversions.ToInteger(value)
            End Set
        End Property

        Public Property Style() As TStyle
            Get
                Return Me._Style
            End Get
            Set(ByVal value As TStyle)
                Me._Style = value
            End Set
        End Property

        Public Property TextValue() As String
            Get
                Return Me._TextValue
            End Get
            Set(ByVal value As String)
                Me._TextValue = value
            End Set
        End Property

        Public Property Title() As String
            Get
                Return Me._Title
            End Get
            Set(ByVal value As String)
                Me._Title = value
            End Set
        End Property

        Public Property TypeCOL() As TypeDB
            Get
                Return DirectCast(Me._TypeDB, TypeDB)
            End Get
            Set(ByVal value As TypeDB)
                Me._TypeDB = CInt(value)
            End Set
        End Property

        Public Property Visible() As Boolean
            Get
                Return Me._Visible
            End Get
            Set(ByVal value As Boolean)
                Me._Visible = value
            End Set
        End Property


        ' Fields
        Private _Label As String = String.Empty
        Private _Name As String = String.Empty
        Private _PageURLColVar As String = String.Empty
        Private _PageURLDestino As String = String.Empty
        Private _Size As Integer = 10
        Private _Style As TStyle = TStyle.FilterField
        Private _TextValue As String = String.Empty
        Private _Title As String = String.Empty
        Private _TypeDB As Integer = 1
        Private _Visible As Boolean = True
        Private prt As PropertyCollection

        ' Nested Types
        Public Enum TStyle
            ' Fields
            FilterField = 0
            ImageButton = 1
        End Enum

        Public Enum TypeDB
            ' Fields
            DATE_T = 2
            DATE_TIME_T = 3
            MONEY_T = 5
            NUMERIC_T = 4
            STRING_T = 1
        End Enum
    End Class
End Namespace

