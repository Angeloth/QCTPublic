Public Class Form11

    'debe recibir la lista de módulos que se pueden seleccionar
    Public objetoCode As String
    Public objetoModule As String
    Public TablaCode As String
    Public elCampoCode As String
    Public elCampoType As String
    Public Modulo1 As String
    Public Modulo2 As String

    Public resCatCode As String
    Public resCatName As String
    Public resCatModule As String
    Public resCatMF As String
    Public resCatMC As String

    Private teniaCheckNumero As Boolean
    Private posiChekAnt As Integer

    Public toyLeyendo As Boolean
    Public CadResults As String
    Public huboExito As Boolean
    Public elEnfoque As String
    Public tablaTemp As New DataTable
    Public ultiCats As DataSet
    Private CatSelected As String
    Private MySource As New AutoCompleteStringCollection()
    Private FuenteFields As New AutoCompleteStringCollection()
    Private CatMaster As New DataTable
    Private FieldMaster As New DataTable
    Private EstoyAgregandoRows As Boolean
    Private Sub Form11_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        teniaCheckNumero = False

        CadResults = ""

        huboExito = False

        CatSelected = ""

        ToolStripTextBox1.Text = ""

        EstoyAgregandoRows = False

        DataGridView1.Columns.Clear()

        DataGridView1.ReadOnly = False

        DataGridView1.EnableHeadersVisualStyles = False

        DataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(64, 114, 196) ' System.Drawing.Color.FromArgb(228, 109, 10)

        DataGridView1.ColumnHeadersDefaultCellStyle.Font = New System.Drawing.Font("Calibri", 14, FontStyle.Bold)

        DataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White

        DataGridView1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders)

        DataGridView1.AdvancedCellBorderStyle.Left = DataGridViewAdvancedCellBorderStyle.None

        DataGridView1.AdvancedCellBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None

        DataGridView1.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None

        DataGridView1.AllowUserToAddRows = False

        DataGridView1.AllowUserToDeleteRows = False

        DataGridView1.AllowUserToOrderColumns = False

        Dim cHekKeyField As New DataGridViewCheckBoxColumn
        cHekKeyField.HeaderText = "Match field?"
        cHekKeyField.Name = "Match"
        DataGridView1.Columns.Add("FieldCode", "Field Name") '0
        DataGridView1.Columns.Add(cHekKeyField) '1
        DataGridView1.Columns.Add("MatchFieldCode", "Match Field Code") '2
        DataGridView1.Columns.Add("MatchFieldName", "Match Field Name") '3
        'datagridview1.columns.add("","")'4
        'si son numeros puedo seleccionar mas de 1, a seleccionar los 2 que aparezca un popup
        'indicando ANTES de guardar cual es mayor que ó menor que
        'X > ó >

        DataGridView1.Columns(0).ReadOnly = True
        DataGridView1.Columns(1).ReadOnly = False
        DataGridView1.Columns(2).ReadOnly = False
        DataGridView1.Columns(3).ReadOnly = True
        'solo hasta que marque 2 que ponga la opción de mach between

        DataGridView2.Columns.Clear()
        DataGridView2.ReadOnly = False

        DataGridView2.EnableHeadersVisualStyles = False

        DataGridView2.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(178, 34, 34) ' System.Drawing.Color.FromArgb(228, 109, 10)

        DataGridView2.ColumnHeadersDefaultCellStyle.Font = New System.Drawing.Font("Calibri", 14, FontStyle.Bold)

        DataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView2.ColumnHeadersDefaultCellStyle.ForeColor = Color.White

        DataGridView2.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders)



        DataGridView2.AdvancedCellBorderStyle.Left = DataGridViewAdvancedCellBorderStyle.None

        DataGridView2.AdvancedCellBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None

        DataGridView2.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None

        DataGridView2.AllowUserToAddRows = False

        DataGridView2.AllowUserToDeleteRows = False

        DataGridView2.AllowUserToOrderColumns = False

        'Al entrar debe vincular solo el catalogo global gb y el del que se este trabajando!

        'serán las únicas opciones

        Label1.Text = "Bind catalog for " & objetoModule & " > " & objetoCode & " > " & TablaCode & " > " & elCampoCode

        ReLlenaPredic()

        Me.CenterToScreen()

    End Sub

    Private Sub ReLlenaPredic()

        Dim xObj As Object = Nothing
        MySource.Clear()
        CatMaster.PrimaryKey = Nothing
        CatMaster.Clear()
        CatMaster.Rows.Clear()
        CatMaster.Columns.Clear()
        AsignaYavePrimariaATabla(CatMaster, "CatalogCode", True)
        CatMaster.Columns.Add("CatalogName", GetType(String))
        CatMaster.Columns.Add("ModuleCode", GetType(String))

        For i = 0 To ultiCats.Tables.Count - 1
            'ModuleCode
            If ultiCats.Tables(i).ExtendedProperties.Item("ModuleCode") = "gb" Or ultiCats.Tables(i).ExtendedProperties.Item("ModuleCode") = objetoModule Then
                'lo agregamos
                'CatalogName
                MySource.Add(CStr(ultiCats.Tables(i).ExtendedProperties.Item("CatalogName") & " " & ultiCats.Tables(i).ExtendedProperties.Item("CatalogCode")))
                CatMaster.Rows.Add({CStr(ultiCats.Tables(i).ExtendedProperties.Item("CatalogCode")), CStr(ultiCats.Tables(i).ExtendedProperties.Item("CatalogName")), CStr(ultiCats.Tables(i).ExtendedProperties.Item("ModuleCode"))})

            End If

        Next

        ToolStripTextBox1.AutoCompleteCustomSource = MySource
        ToolStripTextBox1.AutoCompleteMode = AutoCompleteMode.Suggest
        ToolStripTextBox1.AutoCompleteSource = AutoCompleteSource.CustomSource

        FieldMaster.PrimaryKey = Nothing
        FieldMaster.Clear()
        FieldMaster.Rows.Clear()
        FieldMaster.Columns.Clear()
        AsignaYavePrimariaATabla(FieldMaster, "FieldCode", False)
        FieldMaster.Columns.Add("FieldDescription", GetType(String))
        'Agregamos de base la Company Code
        FieldMaster.Rows.Add({"Company Code", "Company Code (general from mandant)"})

        'cargamos los campos de los templates
        Dim filterDt As DataTable = tablaTemp.Clone()
        Dim result() As DataRow = tablaTemp.Select("TableCode = '" & TablaCode & "'", "Position ASC")

        For Each row As DataRow In result
            filterDt.ImportRow(row)
        Next

        'ahora lo agregamos al set al
        FuenteFields.Clear()
        FuenteFields.Add("Company Code")
        For i = 0 To filterDt.Rows.Count - 1
            FieldMaster.Rows.Add({filterDt.Rows(i).Item(3), filterDt.Rows(i).Item(4)})
            FuenteFields.Add(CStr(filterDt.Rows(i).Item(3)))
        Next


    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Async Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        'Done

        If CatSelected = "" Then
            MsgBox("Please show one catalog first with some columns and records!!", vbCritical, "DQCT")
            Exit Sub
        End If

        If DataGridView1.Rows.Count = 0 Then
            MsgBox("Please show one catalog first with some columns and records!!", vbCritical, "DQCT")
            Exit Sub
        End If

        'OJO, verificar que exista y que NO se repita 
        Dim MatchConditions As String = ""
        Dim cuantosUno As Integer = 0
        Dim FieldCheck As String = ""
        Dim enCuentra As DataRow
        For i = 0 To DataGridView1.Rows.Count - 1

            If DataGridView1.Rows(i).Cells(1).Value = True Then
                cuantosUno = cuantosUno + 1
                If FieldCheck <> "" Then FieldCheck = FieldCheck & "="
                FieldCheck = FieldCheck & CStr(DataGridView1.Rows(i).Tag)
            End If

            If IsNothing(DataGridView1.Rows(i).Cells(2).Value) = False Then

                If DataGridView1.Rows(i).Cells(2).Value <> "" Then
                    'validamos que exista en la tabla
                    enCuentra = FieldMaster.Rows.Find(CStr(DataGridView1.Rows(i).Cells(2).Value))
                    If IsNothing(enCuentra) = True Then
                        'no esta!
                        MsgBox("The field code you entered on row " & CStr(i + 1) & ", on column 'Match Field Code' does not exists! on this template!, please verify!", vbCritical, "DQCT")
                        Exit Sub
                    Else
                        'lo concateno!
                        If MatchConditions <> "" Then MatchConditions = MatchConditions & "-"
                        'será de los campos dependientes al catálogo!
                        MatchConditions = MatchConditions & CStr(DataGridView1.Rows(i).Cells(2).Value) & "#" & CStr(DataGridView1.Rows(i).Tag)

                    End If
                End If

            End If

        Next

        If FieldCheck = "" Then
            MsgBox("Please check EXACTLY one field to match from column 'Match field?'", vbCritical, "DQCT")
            Exit Sub
        End If

        'ultiCats.Tables(CatSelected).Columns(CStr(DataGridView1.Rows(e.RowIndex).Tag)).ExtendedProperties.Item("isText")
        If cuantosUno = 2 Then
            'validación de mayorque ó menor que!
            'ultiCats.Tables(CatSelected).Columns(CStr(DataGridView1.Rows(e.RowIndex).Tag)).ExtendedProperties.Item("isText")
            'falta este form!

        End If

        'If cuantosUno <> 1 Then
        '    MsgBox("Please check EXACTLY one field to match from column 'Match field?'", vbCritical, "DQCT")
        '    Exit Sub
        'End If

        If MatchConditions = "" Then MatchConditions = "None"

        'Se agregan 5 valores al nodo de templates
        'CatMatchField
        'CatMatchConditions
        'CatalogModule
        'CatalogCode
        'CatalogName
        Dim elCatMod As String = ""
        Dim elCatCode As String = ""
        Dim elCatName As String = ""
        Dim elCatMF As String = ""
        Dim elCatMC As String = ""

        elCatMod = ultiCats.Tables(CatSelected).ExtendedProperties.Item("ModuleCode")
        elCatCode = ultiCats.Tables(CatSelected).ExtendedProperties.Item("CatalogCode")
        elCatName = ultiCats.Tables(CatSelected).ExtendedProperties.Item("CatalogName")
        elCatMF = FieldCheck
        elCatMC = MatchConditions

        resCatCode = elCatCode
        resCatName = elCatName
        resCatMC = elCatMC
        resCatMF = elCatMF
        resCatModule = elCatMod

        Dim Camino As String
        Select Case elEnfoque
            Case Is = "T"
                'Viene de los templates
                Camino = RaizFire
                Camino = Camino & "/" & "templates"
                Camino = Camino & "/" & objetoCode
                Camino = Camino & "/" & TablaCode
                Camino = Camino & "/" & elCampoCode

                Await HazPutEnFbSimple(Camino, "CatalogCode", elCatCode)
                Await HazPutEnFbSimple(Camino, "CatalogName", elCatName)
                Await HazPutEnFbSimple(Camino, "CatalogModule", elCatMod)
                Await HazPutEnFbSimple(Camino, "CatMatchField", elCatMF)
                Await HazPutEnFbSimple(Camino, "CatMatchConditions", elCatMC)

                MsgBox("Update done!", vbInformation, "DQCT")

                huboExito = True

                Me.Close()

            Case Is = "B"
                'Build, viene de construction rule
                'catalog route:gb#gb0001:FIELDMATCH:MATNR#A-BUKRS#B
                'Aqui SOLO le vas a regresar el valor!, NO se va a hacer el put
                CadResults = CatSelected & ":" & elCatMF & ":" & elCatMC

                huboExito = True

                Me.Close()


        End Select





    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        'boton de abrir catálogo
        If ToolStripTextBox1.Text = "" Then
            MsgBox("Please type some info to search", vbCritical, "DQCT")
            Exit Sub
        End If

        Dim miCatCode As String
        Dim xObj As Object = Nothing
        xObj = Split(ToolStripTextBox1.Text, " ")
        'la última posición tendrá el valor del catálogo
        miCatCode = xObj(UBound(xObj))

        Dim enCuentra As DataRow

        enCuentra = CatMaster.Rows.Find(miCatCode)

        If IsNothing(enCuentra) = True Then
            MsgBox("Catalog not found!!", vbCritical, "DQCT")
            Exit Sub
        End If

        Dim z As Integer
        z = CatMaster.Rows.IndexOf(enCuentra)

        ReloadCatFields(CatMaster.Rows(z).Item(0), CatMaster.Rows(z).Item(2))

    End Sub

    Private Sub ReloadCatFields(ByVal elCatCode As String, ByVal elModCode As String)

        EstoyAgregandoRows = True

        Dim nameCat As String = elModCode & "#" & elCatCode
        CatSelected = elModCode & "#" & elCatCode
        DataGridView1.Rows.Clear()

        DataGridView2.Columns.Clear()
        DataGridView2.Rows.Clear()

        'Dim elDatv As New DataTable
        'ultiCats.Tables(nameCat).DefaultView.Sort = "Position ASC"

        'y tmb desplegamos la información actual

        For i = 2 To ultiCats.Tables(nameCat).Columns.Count - 1

            DataGridView2.Columns.Add(ultiCats.Tables(nameCat).Columns(i).ColumnName, ultiCats.Tables(nameCat).Columns(i).ExtendedProperties.Item("FieldName"))

            If i = ultiCats.Tables(nameCat).Columns.Count - 1 Then Continue For

            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).HeaderCell.Value = CStr(i - 2 + 1)
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Tag = ultiCats.Tables(nameCat).Columns(i).ColumnName
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = ultiCats.Tables(nameCat).Columns(i).ExtendedProperties.Item("FieldName")
            'si es número, entoonces que pueda seleccionar ambos 

        Next

        DataGridView1.RowHeadersWidth = 70

        For i = 0 To ultiCats.Tables(nameCat).Rows.Count - 1

            DataGridView2.Rows.Add()
            DataGridView2.Rows(DataGridView2.Rows.Count - 1).HeaderCell.Value = CStr(i + 1)
            For j = 2 To ultiCats.Tables(nameCat).Columns.Count - 1
                DataGridView2.Rows(DataGridView2.Rows.Count - 1).Cells(j - 2).Value = ultiCats.Tables(nameCat).Rows(i).Item(j)
            Next

        Next

        DataGridView2.RowHeadersWidth = 70
        EstoyAgregandoRows = False

    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        'cerrar catálogo
        CatSelected = ""
        ToolStripTextBox1.Text = ""
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()

    End Sub

    Private Sub ToolStripTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ToolStripTextBox1.KeyDown
        If e.KeyCode = Keys.Return Then
            ToolStripButton1.PerformClick()
        End If
    End Sub

    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing

        Select Case DataGridView1.CurrentCell.ColumnIndex
            Case Is = 2
                'Field Code
                Dim autoText As TextBox = CType(e.Control, TextBox)
                If (autoText IsNot Nothing) Then
                    autoText.AutoCompleteCustomSource = FuenteFields
                    autoText.AutoCompleteMode = AutoCompleteMode.Suggest
                    autoText.AutoCompleteSource = AutoCompleteSource.CustomSource
                End If

        End Select

    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If EstoyAgregandoRows = True Then Exit Sub
        Dim enCuentra As DataRow
        Dim z As Integer
        Dim esteEsNumero As Boolean = False
        Select Case e.ColumnIndex
            Case Is = 1
                'cambió el valor de true o false
                If DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = True Then

                    If ultiCats.Tables(CatSelected).Columns(CStr(DataGridView1.Rows(e.RowIndex).Tag)).ExtendedProperties.Item("isText") = "Text" Then
                        esteEsNumero = False
                    Else
                        esteEsNumero = True
                    End If

                    If teniaCheckNumero = True Then
                        'puede marcar SOLO otro numero, si NO es otro numero entonces se desmarca!
                        If esteEsNumero = True Then
                            'se marca sin problema!
                            For i = 0 To DataGridView1.Rows.Count - 1

                                If e.RowIndex = i Or e.RowIndex = posiChekAnt Then Continue For 'excluyendo los anteriores menos ellos 2
                                DataGridView1.Rows(i).Cells(e.ColumnIndex).Value = False

                            Next

                        Else
                            'se desmarca el anterior
                            For i = 0 To DataGridView1.Rows.Count - 1

                                If e.RowIndex = i Then Continue For
                                DataGridView1.Rows(i).Cells(e.ColumnIndex).Value = False

                            Next

                            teniaCheckNumero = False
                            posiChekAnt = -1

                        End If
                    Else
                        If esteEsNumero = True Then

                            'se marca normal
                            For i = 0 To DataGridView1.Rows.Count - 1
                                If e.RowIndex = i Then Continue For
                                DataGridView1.Rows(i).Cells(e.ColumnIndex).Value = False
                            Next

                            posiChekAnt = e.RowIndex
                            teniaCheckNumero = True
                        Else
                            'es texto normal
                            For i = 0 To DataGridView1.Rows.Count - 1

                                If e.RowIndex = i Then Continue For
                                DataGridView1.Rows(i).Cells(e.ColumnIndex).Value = False

                            Next

                        End If

                    End If

                    DataGridView1.Rows(e.RowIndex).Cells(2).Value = ""
                    DataGridView1.Rows(e.RowIndex).Cells(2).ReadOnly = True

                Else
                    'habilida la celda contigua
                    DataGridView1.Rows(e.RowIndex).Cells(2).ReadOnly = False
                End If


            Case Is = 2
                enCuentra = FieldMaster.Rows.Find(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value)
                If IsNothing(enCuentra) = True Then
                    DataGridView1.Rows(e.RowIndex).Cells(3).Value = ""
                Else
                    z = FieldMaster.Rows.IndexOf(enCuentra)
                    DataGridView1.Rows(e.RowIndex).Cells(3).Value = CStr(FieldMaster.Rows(z).Item(1))
                End If

        End Select

    End Sub

    Private Sub DataGridView1_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles DataGridView1.CurrentCellDirtyStateChanged
        If DataGridView1.IsCurrentCellDirty Then
            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub Form11_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        If toyLeyendo = True Then
            LoadCatInis()
        End If

    End Sub

    Private Sub LoadCatInis()

        ReloadCatFields(resCatCode, resCatModule)
        'escribo los fields!

        Dim xObj As Object = Nothing
        Dim yObj As Object = Nothing
        'MATNR#
        Dim tabFix As New DataTable
        tabFix.Columns.Add("ObjField", GetType(String))
        tabFix.Columns.Add("CatField", GetType(String))

        If resCatMC <> "None" Then
            xObj = Split(resCatMC, "-")
            For i = 0 To UBound(xObj)
                yObj = Split(xObj(i), "#")
                tabFix.Rows.Add({yObj(0), yObj(1)})
            Next
        End If

        Dim zObj As Object = Nothing
        Dim conMult As Boolean = False
        If resCatMF.Contains("<") = True Or resCatMF.Contains(">") = True Then
            conMult = True
            If resCatMF.Contains("<") = True Then
                zObj = Split(resCatMF, "<")
            Else
                zObj = Split(resCatMF, ">")
            End If
        Else
            conMult = False
            zObj = resCatMF
        End If

        For i = 0 To DataGridView1.Rows.Count - 1

            DataGridView1.Rows(i).Cells(1).Value = False

            If conMult = True Then
                For w = 0 To UBound(zObj)
                    If CStr(DataGridView1.Rows(i).Tag) = zObj(w) Then
                        DataGridView1.Rows(i).Cells(1).Value = True
                    End If
                Next
            Else
                If CStr(DataGridView1.Rows(i).Tag) = resCatMF Then
                    DataGridView1.Rows(i).Cells(1).Value = True
                End If
            End If

            For j = 0 To tabFix.Rows.Count - 1

                If DataGridView1.Rows(i).Tag = tabFix.Rows(j).Item(1) Then

                    DataGridView1.Rows(i).Cells(2).Value = tabFix.Rows(j).Item(0)

                End If

            Next

        Next
    End Sub

End Class