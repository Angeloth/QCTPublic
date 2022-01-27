Imports System.Text
Imports System.IO
Imports xCel = Microsoft.Office.Interop.Excel
Imports oxCel = Microsoft.Office.Interop.Excel
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

    Public Sub ExportToCsv2(ByVal dt As DataGridView, ByVal xFile As String, ByVal limCols As Integer)

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

            '-1, ó -4
            For k = 0 To dt.Columns.Count - limCols 'siempre se elimina las ultimas 3 columnas!
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

            oWrite.WriteLine(CSVHeader.ToString())

        Next

        'CSV.AppendLine(CSVHeader.ToString())
        'oWrite.WriteLine(CSVHeader.ToString())
        oWrite.Flush()

        For r As Integer = 0 To dt.Rows.Count - 2

            Dim CSVLine As StringBuilder = New StringBuilder()
            Dim s As String = ""

            If dt.Rows(r).Visible = False Then Continue For 'lo invisible no lo descarga!

            For k = 0 To dt.Columns.Count - limCols '4

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

        Dim X As Integer
        X = MsgBox("Report exported at: " & xFile & " !" & vbCrLf & "You want to open the file now?", vbQuestion + vbYesNo, "QCT")
        If X = 6 Then
            System.Diagnostics.Process.Start(xFile)
        End If

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

            For Each line In File.ReadLines(filePath).Skip(2)
                dt.Rows.Add(line.Split(","))
            Next

        Catch ex As Exception



        End Try


        Return dt

    End Function

    Public Function ExportaExcelReport(ByVal elSetDa As DataSet, ByVal elFilNam As String, ByRef elError As String) As Boolean

        Dim elRes As Boolean = False

        Try

            Dim xlApp As New xCel.Application
            Dim xlWorkBook As xCel.Workbook
            Dim xlWorkSheet(elSetDa.Tables.Count) As xCel.Worksheet
            Dim misValue As Object = System.Reflection.Missing.Value

            Dim i As Integer
            Dim j As Integer
            Dim k As Integer
            'Dim w As Integer
            'Dim r As Integer
            Dim celIni As String
            Dim celFin As String
            Dim unNombre As String

            xlWorkBook = xlApp.Workbooks.Add(misValue)

            For i = 0 To elSetDa.Tables.Count - 1

                unNombre = elSetDa.Tables(i).ExtendedProperties.Item("TableCode")
                xlWorkSheet(i) = xlWorkBook.Sheets.Add()
                xlWorkSheet(i).Name = unNombre

                xlWorkSheet(i).Cells(1, 1).Value = elSetDa.Tables(i).ExtendedProperties.Item("TableName")
                xlWorkSheet(i).Cells(1, 1).Interior.Color = RGB(112, 173, 71)
                xlWorkSheet(i).Cells(1, 1).Font.Name = "Calibri"
                xlWorkSheet(i).Cells(1, 1).Font.Size = 14
                xlWorkSheet(i).Cells(1, 1).Font.Bold = True
                xlWorkSheet(i).Cells(1, 1).Font.Color = RGB(255, 255, 255)

                celIni = xlWorkSheet(i).Cells(1, 1).Address()
                celFin = xlWorkSheet(i).Cells(1, 5).Address()

                xlWorkSheet(i).Range(celIni & ":" & celFin).Merge()

                For j = 0 To elSetDa.Tables(i).Columns.Count - 1
                    xlWorkSheet(i).Cells(2, j + 1).Value = elSetDa.Tables(i).Columns(j).ColumnName
                    xlWorkSheet(i).Cells(3, j + 1).Value = elSetDa.Tables(i).Columns(j).ExtendedProperties.Item("HeaderText")
                Next

                celIni = xlWorkSheet(i).Cells(2, 1).Address()
                celFin = xlWorkSheet(i).Cells(2, j + 1).Address()

                xlWorkSheet(i).Range(celIni & ":" & celFin).Interior.Color = RGB(128, 128, 128)
                xlWorkSheet(i).Range(celIni & ":" & celFin).Font.Name = "Calibri"
                xlWorkSheet(i).Range(celIni & ":" & celFin).Font.Size = 10
                xlWorkSheet(i).Range(celIni & ":" & celFin).Font.Bold = True
                xlWorkSheet(i).Range(celIni & ":" & celFin).Font.Color = RGB(255, 255, 255)

                celIni = xlWorkSheet(i).Cells(3, 1).Address()
                celFin = xlWorkSheet(i).Cells(3, j + 1).Address()

                xlWorkSheet(i).Range(celIni & ":" & celFin).Interior.Color = RGB(68, 114, 196)
                xlWorkSheet(i).Range(celIni & ":" & celFin).Font.Name = "Calibri"
                xlWorkSheet(i).Range(celIni & ":" & celFin).Font.Size = 10
                xlWorkSheet(i).Range(celIni & ":" & celFin).Font.Bold = True
                xlWorkSheet(i).Range(celIni & ":" & celFin).Font.Color = RGB(255, 255, 255)

                For k = 0 To elSetDa.Tables(i).Rows.Count - 1

                    For j = 0 To elSetDa.Tables(i).Columns.Count - 1
                        xlWorkSheet(i).Cells(4 + k, j + 1).Value = elSetDa.Tables(i).Rows(k).Item(j).ToString()
                    Next

                Next

                xlWorkSheet(i).Cells.Columns.AutoFit()

            Next

            xlWorkBook.SaveAs(elFilNam, xCel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, xCel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            xlWorkBook.Close(True, misValue, misValue)
            xlApp.Quit()

            For i = 0 To elSetDa.Tables.Count - 1
                releaseObject(xlWorkSheet(i))
            Next

            releaseObject(xlWorkBook)
            releaseObject(xlApp)

            elRes = True

        Catch ex As Exception

            elError = ex.Message

        End Try

        Return elRes

    End Function

    Public Sub ImportaExcel(ByVal elArchivo As String, ByRef elSet As DataSet)

        'debería venir un match del objeto a matchear junto con sus hijos!
        'o venir la tabla con sus objetos y buscar
        'https://www.freecodespot.com/blog/csharp-import-excel/
        Dim xlApp As New oxCel.Application
        Dim xlWorkBook As oxCel.Workbook
        Dim xlWorkSheet(elSet.Tables.Count) As oxCel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value

        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim k As Long = 0
        Dim lastRow As Long
        Dim colName As String = ""
        Dim posCOl As Integer = 0
        'https://www.wallstreetmojo.com/vba-last-row/#:~:text=In%20VBA%20when%20we%20have,get%20to%20the%20last%20row.


        Try

            xlWorkBook = xlApp.Workbooks.Open(elArchivo)

            For i = 0 To xlWorkBook.Sheets.Count - 1

                Dim xSh As oxCel.Worksheet = xlWorkBook.Sheets(i)
                If elSet.Tables.IndexOf(xSh.Name) < 0 Then Continue For
                'SI lo encuentra, entonces lo pone!
                'como detectar el total de registros en un excel
                'el ultimo renglon!
                lastRow = xSh.UsedRange.Rows(xSh.UsedRange.Rows.Count).Row 'suponiendo que este funcione!

                For k = 4 To lastRow

                    elSet.Tables(xSh.Name).Rows.Add()

                    j = 1
                    Do While xSh.Cells(2, j).Value <> ""
                        colName = xSh.Cells(2, j).Value
                        If elSet.Tables(xSh.Name).Columns.IndexOf(colName) < 0 Then Continue Do
                        elSet.Tables(xSh.Name).Rows(elSet.Tables(xSh.Name).Rows.Count - 1).Item(colName) = xSh.Cells(k, j).Value
                    Loop

                Next

            Next


            xlWorkBook.Close()
            xlApp.Quit()

            For i = 0 To elSet.Tables.Count - 1
                releaseObject(xlWorkSheet(i))
            Next

            releaseObject(xlWorkBook)
            releaseObject(xlApp)

        Catch ex As Exception



        End Try


    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
            MessageBox.Show("Exception Occured while releasing object " + ex.ToString())
        Finally
            GC.Collect()
        End Try
    End Sub
End Module
