Imports System.Text
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
                    s = s & Qualifier & dt.Rows(r).Cells(c).Value.ToString() & Qualifier
                Else
                    'CSVLine.Append("," & Qualifier & gridResults.Rows(r).Cells(c).Value.ToString() & Qualifier)
                    s = s & "," & Qualifier & dt.Rows(r).Cells(c).Value.ToString() & Qualifier
                End If

            Next
            oWrite.WriteLine(s)
            oWrite.Flush()
            'CSV.AppendLine(CSVLine.ToString())
            'CSVLine.Clear()
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

End Module
