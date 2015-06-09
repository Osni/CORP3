Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Runtime.InteropServices

Namespace CORP3.NET
    Public Class ClsExtendCols
        ' Methods
        Public Sub ComboBox(ByVal sID As String, ByVal dtTab As DataTable, Optional ByVal sATRIBUTO As String = "")
            Me.prt = New PropertyCollection
            Dim prt As PropertyCollection = Me.prt
            prt.Add("ID", sID)
            prt.Add("DATA_TABLE", dtTab)
            prt.Add("ATRIBUTO", sATRIBUTO)
            prt.Add("TYPE", 4)
            prt = Nothing
            Me.ExCols.Add(Me.prt, Nothing, Nothing, Nothing)
        End Sub

        Public Sub ExButton(ByVal sID As String, ByVal sLABEL As String, Optional ByVal sSIZE As Integer = 10, Optional ByVal sTOOLTIPTEXT As String = "", Optional ByVal sATRIBUTO As String = "", Optional ByVal sSCRIPT As String = "")
            Me.prt = New PropertyCollection
            Dim prt As PropertyCollection = Me.prt
            prt.Add("ID", sID)
            prt.Add("LABEL", sLABEL)
            prt.Add("SIZE", sSIZE)
            prt.Add("TOOLTIPTEXT", sTOOLTIPTEXT)
            prt.Add("GRID_EVENT", "ExEventButton")
            prt.Add("ATRIBUTO", sATRIBUTO)
            prt.Add("SCRIPT", sSCRIPT)
            prt.Add("TYPE", 2)
            prt = Nothing
            Me.ExCols.Add(Me.prt, Nothing, Nothing, Nothing)
        End Sub

        Public Sub ExCheckBox(ByVal sID As String, Optional ByVal sLABEL As String = "", Optional ByVal sTOOLTIPTEXT As String = "", Optional ByVal sATRIBUTO As String = "")
            Me.prt = New PropertyCollection
            Dim prt As PropertyCollection = Me.prt
            prt.Add("ID", sID)
            prt.Add("LABEL", sLABEL)
            prt.Add("TOOLTIPTEXT", sTOOLTIPTEXT)
            prt.Add("ATRIBUTO", sATRIBUTO)
            prt.Add("TYPE", 5)
            prt = Nothing
            Me.ExCols.Add(Me.prt, Nothing, Nothing, Nothing)
        End Sub

        Public Sub ExImg(ByVal sID As String, ByVal sSRC As String, Optional ByVal sTOOLTIPTEXT As String = "", Optional ByVal sATRIBUTO As String = "", Optional ByVal sSCRIPT As String = "")
            Me.prt = New PropertyCollection
            Dim prt As PropertyCollection = Me.prt
            prt.Add("ID", sID)
            prt.Add("SRC", sSRC)
            prt.Add("TOOLTIPTEXT", sTOOLTIPTEXT)
            prt.Add("GRID_EVENT", "ExEventImg")
            prt.Add("ATRIBUTO", sATRIBUTO)
            prt.Add("SCRIPT", sSCRIPT)
            prt.Add("TYPE", 3)
            prt = Nothing
            Me.ExCols.Add(Me.prt, Nothing, Nothing, Nothing)
        End Sub

        Public Sub ExTextBox(ByVal sID As String, Optional ByVal sSIZE As Integer = 10, Optional ByVal sMAXLEN As Integer = 10, Optional ByVal sTOOLTIPTEXT As String = "", Optional ByVal sATRIBUTO As String = "")
            Me.prt = New PropertyCollection
            Dim prt As PropertyCollection = Me.prt
            prt.Add("ID", sID)
            prt.Add("SIZE", sSIZE)
            prt.Add("MAXLEN", sMAXLEN)
            prt.Add("TOOLTIPTEXT", sTOOLTIPTEXT)
            prt.Add("ATRIBUTO", sATRIBUTO)
            prt.Add("TYPE", 1)
            prt = Nothing
            Me.ExCols.Add(Me.prt, Nothing, Nothing, Nothing)
        End Sub


        ' Fields
        Public ExCols As Collection = New Collection
        Private prt As PropertyCollection

        ' Nested Types
        Public Enum ExType
            ' Fields
            Button = 2
            CheckBox = 5
            ComboBox = 4
            Img = 3
            TextBox = 1
        End Enum
    End Class
End Namespace

