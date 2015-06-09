Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Collections
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace CORP3.NET
    Public Class ClsSQL1
        ' Methods
        Public Function AddCol(ByVal sName As String, Optional ByVal sValue As Object = "", Optional ByVal TType As TypeSQL = 0) As Boolean
            Dim item As New ArrayList
            Try
                item.Add(sName)
                item.Add(RuntimeHelpers.GetObjectValue(sValue))
                item.Add(TType)
                Me.cCols.Add(item, Nothing, Nothing, Nothing)
            Catch exception1 As Exception
                ProjectData.SetProjectError(exception1)
                Dim exception As Exception = exception1
                Throw New Exception(exception.Message)
            End Try
            Return True
        End Function

        Public Function GetCANCEL(Optional ByVal sSET As String = "Ativo='N'") As String
            If ((Me.sColPrimaryKeyName <> "") And (Me.sColPrimaryKeyValue <> "")) Then
                Return String.Concat(New String() {"UPDATE ", Me.sTable, " SET ", sSET, " WHERE ", Me.sColPrimaryKeyName, "=", Me.sColPrimaryKeyValue})
            End If
            Return String.Concat(New String() {"UPDATE ", Me.sTable, " SET ", sSET, " WHERE ", Me.sWHERE})
        End Function

        Public Function GetDELETE() As String
            If ((Me.sColPrimaryKeyName <> "") And (Me.sColPrimaryKeyValue <> "")) Then
                Return String.Concat(New String() {"DELETE FROM ", Me.sTable, " WHERE ", Me.sColPrimaryKeyName, "=", Me.sColPrimaryKeyValue})
            End If
            Return ("DELETE FROM " & Me.sTable & " WHERE " & Me.sWHERE)
        End Function

        Public Function GetINSERT() As String
            Me.sColName = String.Empty
            Me.sColValue = String.Empty
            Dim count As Integer = Me.cCols.Count
            Me.i = 1
            Do While (Me.i <= count)
                Me.sColName = Conversions.ToString(Operators.ConcatenateObject(Me.sColName, Operators.ConcatenateObject(NewLateBinding.LateIndexGet(Me.cCols.Item(Me.i), New Object() {0}, Nothing), ", ")))
                Me.sColValue = (Me.sColValue & Me.GetValueTType(RuntimeHelpers.GetObjectValue(NewLateBinding.LateIndexGet(Me.cCols.Item(Me.i), New Object() {1}, Nothing)), DirectCast(Conversions.ToInteger(NewLateBinding.LateIndexGet(Me.cCols.Item(Me.i), New Object() {2}, Nothing)), TypeSQL)) & ", ")
                Me.i += 1
            Loop
            Me.sColName = Strings.Left(Me.sColName, (Strings.Len(Me.sColName) - 2))
            Me.sColValue = Strings.Left(Me.sColValue, (Strings.Len(Me.sColValue) - 2))
            Return String.Concat(New String() {"INSERT  INTO ", Me.sTable, "( ", Me.sColName, ") VALUES (", Me.sColValue, ")"})
        End Function

        Public Function GetSELECT(Optional ByVal sColsWHERE As String = "") As String
            Me.sColName = String.Empty
            Dim count As Integer = Me.cCols.Count
            Me.i = 1
            Do While (Me.i <= count)
                Me.sColName = Conversions.ToString(Operators.ConcatenateObject(Me.sColName, Operators.ConcatenateObject(NewLateBinding.LateIndexGet(Me.cCols.Item(Me.i), New Object() {0}, Nothing), ", ")))
                Me.i += 1
            Loop
            Me.sColName = Strings.Left(Me.sColName, (Strings.Len(Me.sColName) - 2))
            If (sColsWHERE <> "") Then
                Me.sWHERE = sColsWHERE
            End If
            Me.sColName = Conversions.ToString(Operators.ConcatenateObject((("SELECT " & Me.sColName) & " FROM " & Me.sTable), Interaction.IIf((Me.sWHERE = ""), "", (" WHERE " & Me.sWHERE))))
            Return Me.SetQSQL(Me.sColName)
        End Function

        Public Function GetUPDATE() As String
            Me.sColValue = String.Empty
            Dim count As Integer = Me.cCols.Count
            Me.i = 1
            Do While (Me.i <= count)
                Me.sColValue = Conversions.ToString(Operators.ConcatenateObject(Me.sColValue, Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(NewLateBinding.LateIndexGet(Me.cCols.Item(Me.i), New Object() {0}, Nothing), "="), Me.GetValueTType(RuntimeHelpers.GetObjectValue(NewLateBinding.LateIndexGet(Me.cCols.Item(Me.i), New Object() {1}, Nothing)), DirectCast(Conversions.ToInteger(NewLateBinding.LateIndexGet(Me.cCols.Item(Me.i), New Object() {2}, Nothing)), TypeSQL))), ", ")))
                Me.i += 1
            Loop
            Me.sColValue = Strings.Left(Me.sColValue, (Strings.Len(Me.sColValue) - 2))
            Return String.Concat(New String() {"UPDATE ", Me.sTable, " SET ", Me.sColValue, " WHERE ", Me.sColPrimaryKeyName, "=", Me.sColPrimaryKeyValue})
        End Function

        Private Function GetValueTType(ByVal sValue As Object, ByVal Type As TypeSQL) As String
            Dim str2 As String = String.Empty
            Select Case Type
                Case TypeSQL.EMPTY_T, TypeSQL.STRING_T
                    Return Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject("'", sValue), "'"))
                Case TypeSQL.DATE_YMD_T
                    Return ("'" & Strings.Format(Conversions.ToDate(sValue), "yyyy-MM-dd") & "'")
                Case TypeSQL.DATE_YMD_HMS_T
                    Return ("'" & Strings.Format(Conversions.ToDate(sValue), "yyyy-MM-dd H:mm:ss") & "'")
                Case TypeSQL.DATE_DMY_T
                    Return ("'" & Strings.Format(Conversions.ToDate(sValue), "dd/MM/yyyy") & "'")
                Case TypeSQL.DATE_DMY__HMS_T
                    Return ("'" & Strings.Format(Conversions.ToDate(sValue), "dd/MM/yyyy H:mm:ss") & "'")
                Case TypeSQL.NUMERIC_T
                    Return Conversions.ToString(sValue)
                Case TypeSQL.MONEY_T
                    Return Strings.Replace(Strings.Replace(Conversions.ToString(sValue), ".", "", 1, -1, CompareMethod.Binary), ",", ".", 1, -1, CompareMethod.Binary)
            End Select
            Return str2
        End Function

        Public Function SetQSQL(ByVal sSQL As String) As String
            Dim list As New ArrayList
            Dim list2 As ArrayList = list
            list2.Add(",")
            list2.Add("SELECT ")
            list2.Add("FROM ")
            list2.Add("WHERE ")
            list2.Add("ORDER BY ")
            list2.Add("GROUP BY ")
            list2.Add("HAVING ")
            list2 = Nothing
            Dim num As Integer = (list.Count - 1)
            Me.i = 0
            Do While (Me.i <= num)
                sSQL = Strings.Replace(sSQL, Conversions.ToString(list.Item(Me.i)), Conversions.ToString(Operators.ConcatenateObject(list.Item(Me.i), ChrW(13) & ChrW(10))), 1, -1, CompareMethod.Binary)
                Me.i += 1
            Loop
            Return sSQL
        End Function


        ' Fields
        Public cCols As Collection = New Collection
        Private i As Integer
        Private l As Long
        Private sColName As String
        Public sColPrimaryKeyName As String
        Public sColPrimaryKeyValue As String
        Private sColType As String
        Private sColValue As String
        Public sTable As String
        Public sWHERE As String

        ' Nested Types
        Public Enum TypeSQL
            ' Fields
            DATE_DMY__HMS_T = 5
            DATE_DMY_T = 4
            DATE_YMD_HMS_T = 3
            DATE_YMD_T = 2
            EMPTY_T = 0
            MONEY_T = 7
            NUMERIC_T = 6
            STRING_T = 1
        End Enum
    End Class
End Namespace

