Imports System.Data.OleDb
Imports System.IO
Imports System.Security
Imports System.Runtime.Serialization

Public Class LookUpTable

    Private mTable As DataTable
    Private mFileName As String
    Private mSheetName As String
    Private mWhereStatement As String
    Private connection As Global.System.Data.OleDb.OleDbConnection

    Sub New(ByVal fileName As String, ByVal sheetName As String)
        mFileName = fileName
        mSheetName = sheetName

        Try
            If File.Exists(mFileName) Then
                mTable = New DataTable
                ' Create a stringbuilder to store connection information
                ' point to the excel spreadsheet (.dataSource) and tell the 
                ' connection what type of driver to use (.provider)
                Dim Builder As New OleDbConnectionStringBuilder _
                With _
                { _
                    .DataSource = mFileName, _
                    .Provider = "Microsoft.ACE.Oledb.12.0" _
                }
                ' Create a new connection object instance
                connection = New Global.System.Data.OleDb.OleDbConnection
                Using connection
                    Builder.Add("Extended Properties", "Excel 12.0;HDR=No;IMEX=1;")
                    connection.ConnectionString = Builder.ConnectionString
                    Using command As OleDbCommand = New OleDbCommand With {.Connection = connection}
                        ' Open the connection to the spreadsheet
                        connection.Open()
                        ' Query the spreadsheet for all of the columns from the 
                        ' sheet mSheetName
                        command.CommandText = "SELECT * FROM [" & mSheetName & "$] "
                        ' put the results of the query into mTable
                        mTable.Load(command.ExecuteReader)
                    End Using
                    connection.Close()
                End Using

                ' Change the generic column names (F1, F2, ... ) to the titles supplied in
                ' the first row of the sheet
                For Each col As DataColumn In mTable.Columns
                    If mTable.Rows(0).Item(col.ColumnName).ToString().Length > 0 Then
                        col.ColumnName = mTable.Rows(0).Item(col.ColumnName).ToString()
                    End If
                Next
                ' Delete the first row of the sheet since it has been copied to the 
                ' column titles
                mTable.Rows(0).Delete()

            Else
                MsgBox("The Lookup table you are attempting to access does not exist at this address: " & _
                       vbCrLf & mFileName & vbCrLf & "You can change the reference to this file in the app configuration file.")
                End
            End If
        Catch ex As Exception
            'MsgBox("An error has occured while trying to access the lookup table at this address: " & _
            '           vbCrLf & mFileName & vbCrLf & "Please check if this table is valid and try again." & vbCrLf & _
            '           ex.ToString())
            End
        End Try

    End Sub

    Public Function lookUp(queryColumn As String, keyColumn As String, keyValue As String, Optional queryColumn1 As String = Nothing,
                                  Optional queryVal1 As String = Nothing, Optional queryColumn2 As String = Nothing,
                                  Optional queryVal2 As String = Nothing) As String()

        Dim results As New List(Of String)

        ' Build the where statement (similar to a SQL Where statement)
        mWhereStatement = keyColumn & " = " & "'" & keyValue & "' "
        If Not IsNothing(queryColumn1) And Not IsNothing(queryVal1) Then
            mWhereStatement = mWhereStatement & "AND " & queryColumn1 & " = " & "'" & queryVal1 & "' "
            If Not IsNothing(queryColumn2) And Not IsNothing(queryVal2) Then
                mWhereStatement = mWhereStatement & "AND " & queryColumn2 & " = " & "'" & queryVal2 & "' "
            End If
        End If

        Try
            ' Store the results of the query in the var rows
            Dim rows As DataRow() = mTable.Select(mWhereStatement)
            If rows.Length > 1 Then
                ' Iterate through the results for a value in the 
                ' queryColumn column, if there in none, add a " "
                ' to the results list
                For Each row As DataRow In rows
                    If Not IsNothing(row.Item(queryColumn)) Then
                        results.Add(row.Item(queryColumn).ToString())
                    Else
                        results.Add(" ")
                    End If
                Next
            Else
                results.Add(" ")
            End If
        Catch ex As Exception
            'MsgBox("an error has occured while trying to access the lookup table at this address: " & _
            '           vbCrLf & mFileName & vbCrLf & "please check if this table is valid and try again." & vbCrLf & _
            '           ex.ToString())
            results.Add("")
        End Try

        Return results.ToArray()

    End Function

End Class
