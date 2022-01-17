Imports System.Text
Imports System.IO
Module Module3

    Public Sub DataGridToCSV(ByRef dt As DataGridView, Qualifier As String, ByVal xFile As String)
        'Dim TempDirectory As String = "A temp Directory"
        'System.IO.Directory.CreateDirectory(TempDirectory)
        Dim oWrite As System.IO.StreamWriter
        'Dim file As String = System.IO.Path.GetRandomFileName & ".csv"
        'oWrite = IO.File.CreateText(TempDirectory & "\" & file)
        oWrite = IO.File.CreateText(xFile)

        Dim CSV As StringBuilder = New StringBuilder()

        Dim i As Integer = 1
        Dim CSVHeader As StringBuilder = New StringBuilder()
        For Each c As DataGridViewColumn In dt.Columns
            If i = 1 Then
                CSVHeader.Append(Qualifier & c.HeaderText.ToString() & Qualifier)
            Else
                CSVHeader.Append("," & Qualifier & c.HeaderText.ToString() & Qualifier)
            End If
            i += 1
        Next

        'CSV.AppendLine(CSVHeader.ToString())
        oWrite.WriteLine(CSVHeader.ToString())
        oWrite.Flush()

        For r As Integer = 0 To dt.Rows.Count - 2

            Dim CSVLine As StringBuilder = New StringBuilder()
            Dim s As String = ""
            For c As Integer = 0 To dt.Columns.Count - 1
                If c = 0 Then
                    'CSVLine.Append(Qualifier & gridResults.Rows(r).Cells(c).Value.ToString() & Qualifier)
                    If IsNothing(dt.Rows(r).Cells(c).Value) = True Then
                        s = s & Qualifier & "" & Qualifier
                    Else
                        s = s & Qualifier & dt.Rows(r).Cells(c).Value.ToString() & Qualifier
                    End If

                Else
                    'CSVLine.Append("," & Qualifier & gridResults.Rows(r).Cells(c).Value.ToString() & Qualifier)
                    If IsNothing(dt.Rows(r).Cells(c).Value) = True Then
                        s = s & "," & Qualifier & "" & Qualifier
                    Else
                        s = s & "," & Qualifier & dt.Rows(r).Cells(c).Value.ToString() & Qualifier
                    End If

                End If

            Next
            oWrite.WriteLine(s)
            oWrite.Flush()

        Next

        'oWrite.Write(CSV.ToString())

        oWrite.Close()
        oWrite = Nothing

        'System.Diagnostics.Process.Start(xFile)
        GC.Collect()
        Dim X As Integer
        X = MsgBox("Report exported at: " & xFile & " !" & vbCrLf & "You want to open the file now?", vbQuestion + vbYesNo, "QCT")
        If X = 6 Then
            System.Diagnostics.Process.Start(xFile)
        End If

    End Sub

    Public Sub ExportToCsv2(ByVal dt As DataGridView, ByVal xFile As String)

        Dim oWrite As System.IO.StreamWriter
        oWrite = IO.File.CreateText(xFile)

        Dim CSV As StringBuilder = New StringBuilder()

        Dim j As Integer = 0
        Dim i As Integer = 1
        Dim k As Integer = 0
        Dim CSVHeader As StringBuilder = New StringBuilder()

        For j = 1 To 2

            CSVHeader.Clear()

            i = 1

            For k = 0 To dt.Columns.Count - 4 'siempre se elimina las ultimas 3 columnas!
                If j = 1 Then
                    If i = 1 Then
                        CSVHeader.Append(dt.Columns(k).Name.ToString())
                    Else
                        CSVHeader.Append("," & dt.Columns(k).Name.ToString())
                    End If
                Else
                    If i = 1 Then
                        CSVHeader.Append(dt.Columns(k).HeaderText.ToString())
                    Else
                        CSVHeader.Append("," & dt.Columns(k).HeaderText.ToString())
                    End If
                End If

                i += 1
            Next

            For Each c As DataGridViewColumn In dt.Columns


            Next

            oWrite.WriteLine(CSVHeader.ToString())

        Next

        'CSV.AppendLine(CSVHeader.ToString())
        'oWrite.WriteLine(CSVHeader.ToString())
        oWrite.Flush()

        For r As Integer = 0 To dt.Rows.Count - 2

            Dim CSVLine As StringBuilder = New StringBuilder()
            Dim s As String = ""

            For k = 0 To dt.Columns.Count - 4

                If k = 0 Then
                    If IsNothing(dt.Rows(r).Cells(k).Value) = True Then
                        s = s & ""
                    Else
                        s = s & dt.Rows(r).Cells(k).Value.ToString()
                    End If
                Else
                    If IsNothing(dt.Rows(r).Cells(k).Value) = True Then
                        s = s & "," & ""
                    Else
                        s = s & "," & dt.Rows(r).Cells(k).Value.ToString()
                    End If
                End If

            Next

            oWrite.WriteLine(s)
            oWrite.Flush()

        Next

        GC.Collect()

        oWrite.Close()
        oWrite = Nothing

    End Sub

    Public Function LoadFromCSV(ByVal filePath As String) As DataTable

        Dim dt As New DataTable

        'For each headerline As Integer In FileAttr.
        Try
            For Each headerLine In File.ReadLines(filePath).Take(1)

                For Each headerItem In headerLine.Split({","}, StringSplitOptions.RemoveEmptyEntries)
                    dt.Columns.Add(headerItem.Trim(), GetType(String))
                Next
            Next

            For Each line In File.ReadLines(filePath).Skip(1)
                dt.Rows.Add(line.Split(","))
            Next

        Catch ex As Exception



        End Try


        Return dt

    End Function

End Module
