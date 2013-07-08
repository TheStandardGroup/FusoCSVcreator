Imports System.Data.SqlClient
Public Class database
    Private sConnection As String = ""
    Private dbQuery As String = ""
    Private tableName As String = ""
    Private dt As DataTable

    Property setSconn As String
        Get
            Return sConnection
        End Get
        Set(value As String)
            sConnection = value
        End Set
    End Property

    ReadOnly Property Name() As String
        Get
            Return tableName
        End Get
    End Property

    Public Sub New(ByVal sC As String, ByVal dbN As String)
        sConnection = sC
        tableName = dbN
    End Sub

    Public Function retrieveData(ByVal query) As Boolean
        dbQuery = query
        Try
            Dim conn As New SqlConnection(sConnection)
            Dim command As New SqlCommand(query, conn)
            Dim dataAdapt As New SqlDataAdapter(query, conn)
            dt = New DataTable
            conn.Open()
            dataAdapt.Fill(dt)
            command.Dispose()
            conn.Close()
            Return True
        Catch ex As Exception
            MessageBox.Show("Error retrieving data from the Database, Please contact a System Administrator", "Error")
            Return False
        End Try
    End Function

    Public Function getDataTable() As DataTable
        Return dt
    End Function

    Public Function getCode() As Int32
        Dim res As Int32
        res = dt.Rows(0).Item(0)
        Return res
    End Function

    Private Sub MessageBoxRow()
        MessageBox.Show(dt.TableName & " " & dt.Rows(0).Item(0) & " " & dt.Rows(0).Item(1) & " " & dt.Rows(0).Item(2))
    End Sub

    Public Sub updateTable(ByVal acNum As Int32, ByVal where As String)
        Try
            Dim conn As New SqlConnection(sConnection)
            Dim sqlString As String = "SELECT * FROM " & Name & " WHERE " & where & " = '" & acNum & "'"
            Using da As New SqlDataAdapter(sqlString, sConnection)
                Using cb As New SqlCommandBuilder(da)
                    conn.Open()
                    da.Update(dt)
                    conn.Close()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            MessageBox.Show("Error updateing the Database, please contact your System Adminstrator.", "Error")
        End Try
    End Sub

    Public Sub updateAll(ByVal num As Integer)
        For Each row As DataRow In dt.Rows
            row.Item("MIN LEVEL") = num
        Next
    End Sub

    Public Sub updateSome(ByVal num As Integer, ByVal idxs() As Integer)
        Dim idx As Integer = 0
        Dim count As Integer = 0
        While (idx < dt.Rows.Count And count < idxs.Length)
            If idx = idxs(count) Then
                dt.Rows(idx)("MIN LEVEL") = num
                count += 1
            End If
            idx += 1
        End While
    End Sub
End Class
