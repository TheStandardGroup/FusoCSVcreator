Public Class frmFusoMain
    Dim startFolder As String = "C:\FusoGPS POLK Data Reporting\1 Raw Data"
    Dim db As database

    Private Sub btnQuit_Click(sender As Object, e As EventArgs) Handles btnQuit.Click
        Me.Dispose()
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        db = New database(getConnString, "FusoList_Live")
    End Sub

    Private Function getConnString() As String
        Return "Server=database;DataBase=FusoList_Live; Integrated Security=SSPI"
    End Function

    Private Sub runProgram()
        Dim dir As New System.IO.DirectoryInfo(startFolder)
        Dim fileList = dir.GetFiles("*.txt", System.IO.SearchOption.TopDirectoryOnly)
        Dim test As Boolean = False
        Dim sort As Dictionary(Of String, System.Collections.ArrayList) = sortFiles(fileList)

        'For Each pair As KeyValuePair(Of String, System.Collections.ArrayList) In sort
        '    Dim testMe As String = ""
        '    Dim idx As Integer = 0
        '    testMe = testMe & pair.Key
        '    While idx < pair.Value.Count
        '        testMe = testMe & "," & fileList(pair.Value(idx)).ToString()
        '        idx += 1
        '    End While
        '    MessageBox.Show(testMe)
        'Next
        Dim sorted = createDictionaryFromFiles(sort, fileList)
        createDupedAndCombined(sorted)
        Dim finalList As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)
        For Each pair As KeyValuePair(Of String, ArrayList) In sorted
            Dim sqlQuery = buildQuery(pair.Value)
            spitOutSql(pair.Key, sqlQuery)
            db.retrieveData(sqlQuery)
            Dim newTable As DataTable = db.getDataTable()
            finalList.Add(pair.Key, createFinalCSV(newTable, pair.Key()))
        Next

        createTotalCSV(finalList)
        MessageBox.Show("Finished making csv files, Now Quiting")
        Me.Dispose()

    End Sub

    Function sortFiles(ByRef fileList As System.IO.FileInfo()) As Dictionary(Of String, ArrayList)
        Dim sort As Dictionary(Of String, System.Collections.ArrayList) = New Dictionary(Of String, System.Collections.ArrayList)
        Dim count As Integer = 0
        For Each file As System.IO.FileInfo In fileList
            Dim fName As String = file.ToString().Substring(0, 5)
            'MessageBox.Show(fName)
            If sort.ContainsKey(fName) Then
                sort.Item(fName).Add(count)
            Else
                sort.Add(fName, New System.Collections.ArrayList)
                sort.Item(fName).Add(count)
            End If
            count += 1
        Next
        Return sort
    End Function

    Private Function createDictionaryFromFiles(ByRef dict As Dictionary(Of String, ArrayList), ByRef fileList As System.IO.FileInfo()) As Dictionary(Of String, ArrayList)
        Dim retMe = New Dictionary(Of String, ArrayList)
        For Each pair As KeyValuePair(Of String, System.Collections.ArrayList) In dict
            Dim count = 0
            Dim newDict = New ArrayList
            Dim fName As String = ""
            While count < pair.Value.Count
                Dim file As String = "/" & fileList(pair.Value(count)).ToString()
                fName = fileList(pair.Value(count)).ToString().Substring(0, 5)
                Dim myReader = New Microsoft.VisualBasic.FileIO.TextFieldParser(startFolder & file)
                myReader.SetDelimiters(",")
                Dim currentRow As String()

                While Not myReader.EndOfData
                    Try
                        currentRow = myReader.ReadFields()
                        Dim currentField As String
                        Dim count2 = 0
                        For Each currentField In currentRow
                            If Not newDict.Contains(currentField) Then
                                If count2 > 0 Then
                                    newDict.Add(currentField)
                                End If
                            End If
                            count2 += 1
                        Next
                    Catch ex As Microsoft.VisualBasic.
                                FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message &
                        "is not valid and will be skipped.")
                    End Try
                End While

                count = count + 1
            End While
            retMe.Add(fName, newDict)
            'retMe.Add(pair.Key)
            'MessageBox.Show(newDict(0))
        Next
        Return retMe
    End Function

    Private Sub createDupedAndCombined(ByRef newF As Dictionary(Of String, System.Collections.ArrayList))
        For Each pair As KeyValuePair(Of String, System.Collections.ArrayList) In newF
            Dim newfile = "C:\FusoGPS POLK Data Reporting\2 Sorted and Combined\" & pair.Key & "_deduped.txt"
            System.IO.File.Create(newfile).Dispose()
            Dim newString As String = ""
            Dim count = 0
            While count < pair.Value.Count
                newString += pair.Value(count)
                If count < pair.Value.Count - 1 Then
                    newString += ","
                End If
                count += 1
            End While
            Dim objWriter As New System.IO.StreamWriter(newfile, False)
            objWriter.WriteLine(newString)
            objWriter.Close()
        Next
    End Sub

    Private Function buildQuery(ByRef list As ArrayList)
        Dim retMe = "SELECT REC_ID,BUS_NAME,BUS_ADDRESS1,BUS_ADDRESS2,BUS_ADDRESS3,BUS_CITY,BUS_ST_PROV,BUS_ZIP_CODE,BUS_COUNTY,BUS_COUNTRY,BUS_PHONE1,BUS_EMAIL,TRK_VIN,REC_DLR_NAME,REC_DLR_NUMBER,ListTypeID "
        retMe = retMe + "FROM FusoList_Live.dbo.tblCorpLists WHERE REC_ID IN ("
        Dim intI = 0
        While intI < list.Count
            retMe = retMe + list(intI).ToString()
            If intI < list.Count - 1 Then
                retMe = retMe + ","
            End If
            intI += 1
        End While
        retMe = retMe + ") AND ListTypeID LIKE 'Polk%' ORDER BY ListTypeID asc"
        Return retMe
    End Function

    Private Sub spitOutSql(ByRef name As String, ByRef sql As String)
        Dim newfile = "C:\FusoGPS POLK Data Reporting\3 SQLQUERY\" & name & "_sql.sql"
        System.IO.File.Create(newfile).Dispose()
        Dim objWriter As New System.IO.StreamWriter(newfile, False)
        objWriter.WriteLine(sql)
        objWriter.Close()
    End Sub

    Private Function createFinalCSV(ByRef dt As DataTable, ByRef name As String) As Integer
        Dim fileList As ArrayList = New ArrayList()
        Dim count = 0
        Dim headerString As String = ""
        While count < dt.Columns.Count
            headerString += Chr(34) & dt.Columns(count).ColumnName() & Chr(34)
            If count < dt.Columns.Count - 1 Then
                headerString += ","
            End If
            count += 1
        End While
        fileList.Add(headerString)
        Dim newfile = "C:\FusoGPS POLK Data Reporting\4 Final CSV\" & name & "_03613_061313.csv"
        System.IO.File.Create(newfile).Dispose()
        For Each dr As DataRow In dt.Rows
            Dim myItems = dr.ItemArray()
            Dim newString = ""
            count = 0
            While count < myItems.Count
                Dim addme As String = ""
                Try
                    addme += myItems(count).ToString()
                Catch ex As Exception
                    addme = " "
                End Try
                newString += Chr(34) & addme & Chr(34)
                If (count < myItems.Count - 1) Then
                    newString += ","
                End If
                count += 1
            End While
            fileList.Add(newString)

        Next
        count = 0
        Dim objWriter As New System.IO.StreamWriter(newfile, True)
        While count < fileList.Count
            objWriter.WriteLine(fileList(count))
            count += 1
        End While
        objWriter.Close()
        Return count - 1
    End Function

    Private Sub createTotalCSV(ByRef dict As Dictionary(Of String, Integer))
        Dim fileList As ArrayList = New ArrayList()
        fileList.Add(Chr(34) & "Dealer Code" & Chr(34) & "," & Chr(34) & "Polk Records" & Chr(34))
        For Each pair As KeyValuePair(Of String, Integer) In dict
            fileList.Add(Chr(34) & pair.Key() & Chr(34) & "," & Chr(34) & pair.Value.ToString() & Chr(34))
        Next
        Dim newfile = "C:\FusoGPS POLK Data Reporting\PolkRecordUsage_CountOnly_JUNE13.csv"
        System.IO.File.Create(newfile).Dispose()
        Dim count = 0
        Dim objWriter As New System.IO.StreamWriter(newfile, True)
        While count < fileList.Count
            objWriter.WriteLine(fileList(count))
            count += 1
        End While
        objWriter.Close()
    End Sub

    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click
        runProgram()
    End Sub
End Class
