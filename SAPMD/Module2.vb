Imports Firebase.Database
Imports Firebase.Database.Query
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Net
Imports System.Threading
Module Module2

    'Modulo para procesar los pedidos del form1
    Public elQbusca As String
    Public buscaEnlower As String
    Public oNode As TreeNode
    Public euReka As Boolean
    Public TabColumnas As New DataTable
    Public niuTabcols As New DataTable

    Public Sub buSkaXNodo(ByVal elNOdin As TreeNode)

        Dim noDoekiz As TreeNode
        Dim i As Integer
        Dim nomenLow As String = elNOdin.Name.ToLower()
        Dim texenLow As String = elNOdin.Text.ToLower()

        If elNOdin.Name = elQbusca Or elNOdin.Text = elQbusca Or elNOdin.Name.Contains(elQbusca) = True Or elNOdin.Text.Contains(elQbusca) = True Then '
            'suponiendo que esta en la primera!!
            'Or elNOdin.Name.Contains(buscaEnlower) = True Or elNOdin.Text.Contains(buscaEnlower) = True
            euReka = True
            oNode = elNOdin
        End If

        If euReka = True Then Exit Sub

        If elNOdin.Nodes.Count > 0 Then

            noDoekiz = elNOdin.FirstNode

            'nomenLow = ""
            'texenLow = ""
            'Or noDoekiz.Name.Contains(buscaEnlower) = True Or noDoekiz.Text.Contains(buscaEnlower) = True
            If noDoekiz.Name = elQbusca Or noDoekiz.Text = elQbusca Or noDoekiz.Name.Contains(elQbusca) = True Or noDoekiz.Text.Contains(elQbusca) = True Then
                euReka = True
                oNode = noDoekiz
            End If

            If euReka = True Then Exit Sub

            Call buSkaXNodo(noDoekiz)

            For i = 1 To elNOdin.Nodes.Count - 1
                noDoekiz = noDoekiz.NextNode

                If noDoekiz.Name = elQbusca Or noDoekiz.Text = elQbusca Or noDoekiz.Name.Contains(elQbusca) = True Or noDoekiz.Text.Contains(elQbusca) = True Then
                    'se iguala al nodo objetivo y truena todo el ciclo!
                    'noDoekiz.Name.Contains(buscaEnlower) = True Or noDoekiz.Text.Contains(buscaEnlower) = True
                    euReka = True
                    oNode = noDoekiz
                End If

                If euReka = True Then Exit Sub

                Call buSkaXNodo(noDoekiz)
            Next

        End If

    End Sub


    Public Sub buSkaNOdo(ByVal elNOdin As TreeNode)

        Dim noDoekiz As TreeNode
        Dim i As Integer

        If elNOdin.Text = elQbusca Then '
            'suponiendo que esta en la primera!!
            euReka = True
            oNode = elNOdin
        End If

        If euReka = True Then Exit Sub

        If elNOdin.Nodes.Count > 0 Then

            noDoekiz = elNOdin.FirstNode

            If noDoekiz.Text = elQbusca Then
                euReka = True
                oNode = noDoekiz
            End If

            If euReka = True Then Exit Sub

            Call buSkaNOdo(noDoekiz)

            For i = 1 To elNOdin.Nodes.Count - 1
                noDoekiz = noDoekiz.NextNode

                If noDoekiz.Text = elQbusca Then
                    'se iguala al nodo objetivo y truena todo el ciclo!
                    euReka = True
                    oNode = noDoekiz
                End If

                If euReka = True Then Exit Sub

                Call buSkaNOdo(noDoekiz)
            Next

        End If

    End Sub

    Public Sub QuitaDuplisCats(ByRef elGrid As DataGridView, ByRef Resultado As String)

        Dim setDupli As New DataTable
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim Dup As Integer = 0
        Dim Bla As Integer = 0
        Dim sisTa As Boolean = False
        Dim enCuentra As DataRow
        Dim elValor As String = ""
        Dim unValor As String = ""
        Dim seVa As Boolean = False

        Call AsignaYavePrimariaATabla(setDupli, "Yave", True)

        i = 0
        Do While i < elGrid.Rows.Count - 1

            'si alguna columna tiene un campo vacío entonces se elimina el renglon completo!
            seVa = False
            elValor = ""

            For j = 0 To elGrid.Columns.Count - 2

                If IsNothing(elGrid.Rows(i).Cells(j).Value) = True Then
                    Bla = Bla + 1
                    elGrid.Rows.RemoveAt(i)
                    Continue Do
                    Exit For
                End If

                If elGrid.Rows(i).Cells(j).Value = "" Then
                    Bla = Bla + 1
                    elGrid.Rows.RemoveAt(i)
                    Continue Do
                    Exit For
                End If

                unValor = elGrid.Rows(i).Cells(j).Value
                unValor = unValor.Trim()
                elGrid.Rows(i).Cells(j).Value = unValor

                If elValor <> "" Then elValor = elValor & "#"

                elValor = elValor & unValor

            Next

            elValor = elValor.Trim()

            enCuentra = Nothing
            enCuentra = setDupli.Rows.Find(elValor)
            If IsNothing(enCuentra) = False Then
                sisTa = True
            Else
                sisTa = False
            End If

            If sisTa = True Then
                elGrid.Rows.RemoveAt(i)
                Dup = Dup + 1
                Continue Do
            End If

            setDupli.Rows.Add({elValor})

            i = i + 1

        Loop

        If Dup + Bla > 0 Then
            'MsgBox("Removed: " & Dup & " duplicates and " & Bla & " blanks!", vbInformation, "ContiShip")

            Resultado = "Removed: " & Dup & " duplicates and " & Bla & " blanks!"
        Else
            'MsgBox("No duplicates or blanks found!", vbInformation, "ContiShip")
            Resultado = "No duplicates or blanks found!"
        End If

        For i = 0 To elGrid.Rows.Count - 1
            elGrid.Rows(i).HeaderCell.Value = CStr(i + 1)
        Next

        setDupli.Rows.Clear()
        setDupli.Dispose()

    End Sub

    Public Sub QuitaDuplicadosDeGrid(ByRef elGrid As DataGridView, ByVal laColumnaYave As Integer, ByRef Resultado As String)

        Dim setDupli As New DataTable
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim Dup As Integer = 0
        Dim Bla As Integer = 0
        Dim sisTa As Boolean = False
        Dim enCuentra As DataRow
        Call AsignaYavePrimariaATabla(setDupli, "Yave", True)

        Dim elValor As String = ""

        'setDupli.Columns.Add()
        i = 0
        Do While i < elGrid.Rows.Count - 1

            If IsDBNull(elGrid.Rows(i).Cells(laColumnaYave).Value) = True Then
                Bla = Bla + 1
                elGrid.Rows.RemoveAt(i)
                Continue Do
            End If

            If elGrid.Rows(i).Cells(laColumnaYave).Value = "" Then
                Bla = Bla + 1
                elGrid.Rows.RemoveAt(i)
                Continue Do
            End If

            elValor = CStr(elGrid.Rows(i).Cells(laColumnaYave).Value)
            elValor = elValor.Trim()

            elGrid.Rows(i).Cells(laColumnaYave).Value = elValor

            enCuentra = Nothing
            enCuentra = setDupli.Rows.Find(elValor)
            If IsNothing(enCuentra) = False Then
                sisTa = True
            Else
                sisTa = False
            End If

            'sisTa = False
            'For j = 0 To setDupli.Rows.Count - 1
            '    If setDupli.Rows(j).Item(0) = elGrid.Rows(i).Cells(laColumnaYave).Value Then
            '        sisTa = True
            '        Exit For
            '    End If

            'Next

            If sisTa = True Then
                elGrid.Rows.RemoveAt(i)
                Dup = Dup + 1
                Continue Do
            End If

            setDupli.Rows.Add({elValor})

            i = i + 1

        Loop

        If Dup + Bla > 0 Then
            'MsgBox("Removed: " & Dup & " duplicates and " & Bla & " blanks!", vbInformation, "ContiShip")

            Resultado = "Removed: " & Dup & " duplicates and " & Bla & " blanks!"
        Else
            'MsgBox("No duplicates or blanks found!", vbInformation, "ContiShip")
            Resultado = "No duplicates or blanks found!"
        End If

        For i = 0 To elGrid.Rows.Count - 1
            elGrid.Rows(i).HeaderCell.Value = CStr(i + 1)
        Next

        setDupli.Rows.Clear()
        setDupli.Dispose()

    End Sub

    Public Sub QuitaDuplicadosMultiplesGrid(ByRef elGrid As DataGridView, ByVal ColumnasYave As String, ByRef Resultado As String)
        Dim setDupli As New DataTable
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim Dup As Integer = 0
        Dim Bla As Integer = 0
        Dim sisTa As Boolean = False
        setDupli.Columns.Add()
        Dim xObj As Object
        Dim Yav As String = ""
        xObj = Split(ColumnasYave, "#")

        i = 0
        Do While i < elGrid.Rows.Count - 1
            Yav = ""

            For j = 0 To UBound(xObj)
                If IsDBNull(elGrid.Rows(i).Cells(CInt(xObj(j))).Value) = True Then
                    Bla = Bla + 1
                    elGrid.Rows.RemoveAt(i)
                    Continue Do
                End If
            Next

            For j = 0 To UBound(xObj)
                If elGrid.Rows(i).Cells(CInt(xObj(j))).Value = "" Then
                    Bla = Bla + 1
                    elGrid.Rows.RemoveAt(i)
                    Continue Do
                End If
            Next

            Yav = ""
            For j = 0 To UBound(xObj)
                If j <> 0 Then Yav = Yav & "#"
                Yav = Yav & elGrid.Rows(i).Cells(CInt(xObj(j))).Value
            Next

            sisTa = False
            For j = 0 To setDupli.Rows.Count - 1
                If setDupli.Rows(j).Item(0) = Yav Then
                    sisTa = True
                    Exit For
                End If

            Next

            If sisTa = True Then
                elGrid.Rows.RemoveAt(i)
                Dup = Dup + 1
                Continue Do
            End If

            setDupli.Rows.Add({Yav})

            i = i + 1

        Loop

        If Dup + Bla > 0 Then
            'MsgBox("Removed: " & Dup & " duplicates and " & Bla & " blanks!", vbInformation, "ContiShip")

            Resultado = "Removed: " & Dup & " duplicates and " & Bla & " blanks!"
        Else
            'MsgBox("No duplicates or blanks found!", vbInformation, "ContiShip")
            Resultado = "No duplicates or blanks found!"
        End If

        For i = 0 To elGrid.Rows.Count - 1
            elGrid.Rows(i).HeaderCell.Value = CStr(i + 1)
        Next

        setDupli.Rows.Clear()
        setDupli.Dispose()

    End Sub
    Public Sub TomaInfoParaTemplates(ByVal elGrid As DataGridView, ByRef elDs As DataSet, ByVal laTabInd As Integer, ByVal ColYave As Integer, ByVal fromCol As Integer, ByVal toCol As Integer)

        Dim xYave As String = ""
        Dim tipoCell As String = ""
        Dim posDg As Integer = 0
        elDs.Tables(laTabInd).Rows.Clear()
        For i = 0 To elGrid.Rows.Count - 1

            xYave = elGrid.Rows(i).Cells(ColYave).Value
            elDs.Tables(laTabInd).Rows.Add({xYave, "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""})

            elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(1) = elGrid.Rows(i).Cells(1).Value
            If elGrid.Rows(i).Cells(2).Value = True Then
                elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(2) = "X"
            Else
                elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(2) = ""
            End If
            elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(3) = elGrid.Rows(i).HeaderCell.Value 'posicion!
            posDg = CInt(elGrid.Rows(i).HeaderCell.Value)
            elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(4) = LetraNumero.Tables(0).Rows(posDg - 1).Item(1)

            For j = fromCol To toCol 'elGrid.Columns.Count - 1

                Select Case elGrid.Rows(i).Cells(j).Value
                    Case Is = "No selection", "None", "N/A"
                        elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(j + 2) = ""
                    Case Else
                        elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(j + 2) = elGrid.Rows(i).Cells(j).Value
                End Select

            Next

        Next

    End Sub

    Public Sub TomaInfoPara1RowTemplate(ByVal elGrid As DataGridView, ByRef elDs As DataSet, ByVal laTabInd As Integer, ByVal ColYave As Integer, ByVal fromCol As Integer, ByVal toCol As Integer, ByVal elRenglon As Integer)
        Dim xYave As String = ""
        Dim posDg As Integer = 0

        elDs.Tables(laTabInd).Rows.Clear()

        'For i = 0 To elGrid.Rows.Count - 1
        posDg = elRenglon + 1

        xYave = elGrid.Rows(elRenglon).Cells(ColYave).Value
        elDs.Tables(laTabInd).Rows.Add({xYave, "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""})

        elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(1) = elGrid.Rows(elRenglon).Cells(1).Value
        If elGrid.Rows(elRenglon).Cells(2).Value = True Then
            elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(2) = "X"
        Else
            elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(2) = ""
            End If
        elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(3) = posDg ' elGrid.Rows(elRenglon).HeaderCell.Value 'posicion!
        'posDg = CInt(elGrid.Rows(elRenglon).HeaderCell.Value)

        elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(4) = LetraNumero.Tables(0).Rows(posDg - 1).Item(1)

            For j = fromCol To toCol 'elGrid.Columns.Count - 1

            Select Case elGrid.Rows(elRenglon).Cells(j).Value
                Case Is = "No selection", "None", "N/A"
                    elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(j + 2) = ""
                Case Else
                    elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(j + 2) = elGrid.Rows(elRenglon).Cells(j).Value
            End Select

        Next

        'Next
    End Sub

    Public Sub TomaInfoDeGridConCampoLlave(ByVal elGrid As DataGridView, ByRef elDs As DataSet, ByVal laTabInd As Integer, ByVal columnasYave As String, ByVal fromCol As Integer, ByVal codeTabla As String)

        Dim xObj As Object
        Dim laYav As String = ""

        xObj = Split(columnasYave, "#")
        Dim w As Integer = 0

        Dim posDg As Integer = 0
        Dim tipoCell As String = ""

        For i = 0 To elGrid.Rows.Count - 1
            laYav = codeTabla
            For j = 0 To UBound(xObj)
                'se supone que ya no deberia ocurrir errots!
                'If j <> 0 Then 
                laYav = laYav & "#"
                laYav = laYav & elGrid.Rows(i).Cells(CInt(xObj(j))).Value
            Next

            elDs.Tables(laTabInd).Rows.Add({laYav})

            'elGrid.Rows(i).HeaderCell.Value
            w = fromCol

            For j = 0 To 2 'elGrid.Columns.Count - 1

                If IsDBNull(elGrid.Rows(i).Cells(j).Value) = True Then
                    elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(w) = ""
                    w = w + 1
                    Continue For
                End If

                tipoCell = elGrid.Rows(i).Cells(j).GetType.ToString()

                Select Case tipoCell
                    Case Is = "System.Windows.Forms.DataGridViewTextBoxCell"
                        If elGrid.Rows(i).Cells(j).Value = "" Then
                            elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(w) = ""
                            w = w + 1
                            Continue For
                        End If

                        elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(w) = elGrid.Rows(i).Cells(j).Value

                    Case Is = "System.Windows.Forms.DataGridViewCheckBoxCell"
                        If elGrid.Rows(i).Cells(j).Value = True Then
                            elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(w) = "X"
                        Else
                            elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(w) = ""
                        End If

                End Select

                w = w + 1

                'If j = 0 Then
                'elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(j + 1) = CStr(elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(j + 1)).ToLower()
                'End If

            Next

            elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(w) = elGrid.Rows(i).HeaderCell.Value 'posicion!
            posDg = CInt(elGrid.Rows(i).HeaderCell.Value)
            elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(w + 1) = LetraNumero.Tables(0).Rows(posDg - 1).Item(1)

        Next


    End Sub

    Public Sub TomaInfoParaCats(ByVal tabOrig As DataTable, ByRef elGrid As DataGridView, ByRef tabNuevos As DataTable, ByRef tabRemoves As DataTable, ByRef tabUpds As DataTable)

        Dim i As Integer
        Dim j As Integer
        Dim z As Integer
        Dim laKey As String = ""
        Dim Descri As String = ""
        Dim enCuentra As DataRow
        Dim tabActual As DataTable
        tabNuevos.Rows.Clear()
        tabRemoves.Rows.Clear()
        tabUpds.Rows.Clear()

        tabActual = tabNuevos.Clone()

        For i = 0 To elGrid.Rows.Count - 1

            tabActual.Rows.Add({"xYave"})

            Descri = ""
            laKey = ""
            For j = 0 To elGrid.Columns.Count - 1

                tabActual.Rows(tabActual.Rows.Count - 1).Item(j + 1) = elGrid.Rows(i).Cells(j).Value

                If j = elGrid.Columns.Count - 1 Then
                    Descri = elGrid.Rows(i).Cells(j).Value
                    Continue For
                End If

                If laKey <> "" Then laKey = laKey & "#"

                laKey = laKey & elGrid.Rows(i).Cells(j).Value
            Next

            tabActual.Rows(tabActual.Rows.Count - 1).Item(0) = laKey

            'tambien es posible que en lo nuevo ya no existan registros de lo viejo y se tengan que eliminar!

        Next


        'aqui estamos comparando lo nuevo vs lo viejo
        For i = 0 To tabActual.Rows.Count - 1
            enCuentra = tabOrig.Rows.Find(tabActual.Rows(i).Item(0))
            If IsNothing(enCuentra) = True Then
                'NO estaba antes, es NUEVO registro!
                tabNuevos.Rows.Add({tabActual.Rows(i).Item(0)})
                For j = 0 To elGrid.Columns.Count - 1
                    tabNuevos.Rows(tabNuevos.Rows.Count - 1).Item(j + 1) = elGrid.Rows(i).Cells(j).Value
                Next
            Else
                z = tabOrig.Rows.IndexOf(enCuentra)
                If tabOrig.Rows(z).Item(tabOrig.Columns.Count - 1) <> tabActual.Rows(i).Item(tabActual.Columns.Count - 1) Then
                    'cambió la descripción!, se agrega a update!
                    tabUpds.Rows.Add({tabOrig.Rows(z).Item(1)}) 'la yave de FB
                    For j = 0 To elGrid.Columns.Count - 1
                        tabUpds.Rows(tabUpds.Rows.Count - 1).Item(j + 1) = elGrid.Rows(i).Cells(j).Value
                    Next
                End If

            End If

        Next

        For i = 0 To tabOrig.Rows.Count - 1
            enCuentra = tabActual.Rows.Find(tabOrig.Rows(i).Item(0))
            If IsNothing(enCuentra) = True Then
                'No esta, es registro para borrar!
                tabRemoves.Rows.Add({tabOrig.Rows(i).Item(1), tabOrig.Rows(i).Item(0)}) 'yave de FB primero porque se borra!
            Else
                'si estaba!, nada q hacer
            End If

        Next

        'comparar vs el grid!
        'tendríamos que tomar la tabla original y comparar vs los campos llave y las descripciones!





    End Sub

    Public Sub TomaInfoDeGrid(ByRef elGrid As DataGridView, ByRef elDs As DataSet, ByVal laTabInd As Integer)

        Dim i As Long
        Dim j As Integer
        Dim valX As String = ""
        Dim valY As String = ""
        For i = 0 To elGrid.Rows.Count - 1

            valX = CStr(elGrid.Rows(i).Cells(0).Value)

            If valX <> " " Then
                valX = CStr(elGrid.Rows(i).Cells(0).Value)
                valX = valX.Trim()
            End If

            elDs.Tables(laTabInd).Rows.Add({valX, "1"})

            For j = 1 To elGrid.Columns.Count - 1

                If IsNothing(elGrid.Rows(i).Cells(j).Value) = True Then
                    valY = ""
                Else
                    valY = elGrid.Rows(i).Cells(j).Value
                End If

                If valY <> " " Then
                    valY = valY.Trim()
                End If
                elDs.Tables(laTabInd).Rows(elDs.Tables(laTabInd).Rows.Count - 1).Item(j) = valY ' elGrid.Rows(i).Cells(j).Value

            Next

        Next

    End Sub

    Public Sub buSkaNuevoNodo(ByVal elNOdin As TreeNode)

        Dim noDoekiz As TreeNode
        Dim i As Integer

        If elNOdin.Name = elQbusca Then '
            'suponiendo que esta en la primera!!
            euReka = True
            oNode = elNOdin
        End If

        If euReka = True Then Exit Sub

        If elNOdin.Nodes.Count > 0 Then

            noDoekiz = elNOdin.FirstNode

            If noDoekiz.Name = elQbusca Then
                euReka = True
                oNode = noDoekiz
            End If

            If euReka = True Then Exit Sub

            Call buSkaNuevoNodo(noDoekiz)

            For i = 1 To elNOdin.Nodes.Count - 1
                noDoekiz = noDoekiz.NextNode

                If noDoekiz.Name = elQbusca Then
                    'se iguala al nodo objetivo y truena todo el ciclo!
                    euReka = True
                    oNode = noDoekiz
                End If

                If euReka = True Then Exit Sub

                Call buSkaNuevoNodo(noDoekiz)
            Next

        End If

    End Sub




End Module
