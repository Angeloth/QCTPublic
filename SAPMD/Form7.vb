Public Class Form7

    Public allTemps As DataSet
    Public allDepes As DataSet
    Public TablaDatos As DataTable
    Public ConsObject As String
    Public ConsTable As String
    Public ConsField As String
    Public ConsModule As String
    Public BuildDs As DataSet
    Public MisCatas As DataSet
    Public posiTabla As Integer
    Private xSource As New AutoCompleteStringCollection()
    Private miCols As New DataSet
    Private reglaSelected As String
    Public maxCharLimit As Integer
    Private EstoyAgregandoRows As Boolean

    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Form para las reglas de construccion
        'Construction Rule, Debe tener su propia interfaz
        'Parent/Child dpendencie, debe tener su propia interfaz!
        'Nodo: construction/build/md40/md40-0001/ROUTE, construction/conditional/, construction/
        'va a tener POSTS por cada regla de construccion!
        'ROUTE-0001
        'ROUTE-0002
        'ROUTE-0003
        'ROUTE-0004
        EstoyAgregandoRows = False

        reglaSelected = ""

        ListView1.Items.Clear()
        ListView1.CheckBoxes = True
        ListView1.View = View.Details
        ListView1.Columns.Add("Rule", 90)
        ListView1.SmallImageList = ImageList1

        Label1.Text = "Construction logic for " & ConsObject & "/" & ConsTable & "/" & ConsField
        ToolStripLabel1.Text = "No selection"

        DataGridView1.ReadOnly = False

        DataGridView1.EnableHeadersVisualStyles = False

        DataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(178, 34, 34) ' System.Drawing.Color.FromArgb(228, 109, 10)
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

        DataGridView1.ColumnHeadersDefaultCellStyle.Font = New System.Drawing.Font("Calibri", 12, FontStyle.Bold)
        DataGridView1.Columns.Clear()

        Dim FillRule As New DataGridViewComboBoxColumn

        FillRule.Name = "Rule"
        FillRule.HeaderText = "Rule"

        FillRule.Items.Clear()
        FillRule.Items.Add("No selection")
        FillRule.Items.Add("A - From Catalog")
        FillRule.Items.Add("B - Free Text")
        FillRule.Items.Add("C - Running number")
        FillRule.Items.Add("D - Fixed Value")
        FillRule.Items.Add("E - External object") 'OJOO, puede ser de un objeto local ó externo!!, todos menos su mismo campo!!

        FillRule.DefaultCellStyle.BackColor = Color.White
        FillRule.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
        FillRule.DisplayStyleForCurrentCellOnly = True

        DataGridView1.Columns.Add("Consec", "#'s") '0
        DataGridView1.Columns.Add(FillRule) '1
        DataGridView1.Columns.Add("Value", "Value") '2
        DataGridView1.Columns.Add("CharLen", "Char Length") '3
        DataGridView1.Columns.Add("Signo", "Signo") '4

        For i = 0 To DataGridView1.Columns.Count - 1
            DataGridView1.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next

        DataGridView1.Columns(0).ReadOnly = True
        DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        'miCols
        Dim iAveprim(1) As DataColumn
        Dim kEys As New DataColumn()
        kEys.ColumnName = "Key"
        iAveprim(0) = kEys

        miCols.Tables.Clear()

        miCols.Tables.Add()
        miCols.Tables(0).Columns.Add(kEys)
        miCols.Tables(0).Columns.Add("Descrip")
        miCols.Tables(0).CaseSensitive = True
        miCols.Tables(0).PrimaryKey = iAveprim

        Dim xObj As Object = Nothing
        Dim mixNombre As String = ""
        Dim posMix As Integer = 0
        xSource.Clear()

        For j = 0 To MisCatas.Tables.Count - 1
            xObj = Split(MisCatas.Tables(j).TableName, "#")

            mixNombre = CStr(xObj(0)) & " " & CStr(MisCatas.Tables(j).Columns(0).ColumnName)
            posMix = xSource.IndexOf(mixNombre)
            If posMix >= 0 Then Continue For

            xSource.Add(mixNombre)
            miCols.Tables(0).Rows.Add({mixNombre, MisCatas.Tables(j).TableName})
            'gb#gb0007
        Next

        Me.CenterToScreen()

    End Sub
    Private Sub CargaListaDeReglas()

        ListView1.Items.Clear()

        Dim Yastaba As Boolean = False
        Dim xDs As New DataTable
        xDs.Columns.Add("Ruls")

        Dim filterDt As DataTable = BuildDs.Tables(posiTabla).Clone()
        Dim result() As DataRow = BuildDs.Tables(posiTabla).Select("TableCode = '" & ConsTable & "' AND FieldCode='" & ConsField & "'")

        For Each row As DataRow In result
            filterDt.ImportRow(row)
        Next

        For i = 0 To filterDt.Rows.Count - 1
            Yastaba = False
            For j = 0 To xDs.Rows.Count - 1
                If xDs.Rows(j).Item(0) = filterDt.Rows(i).Item(2) Then
                    Yastaba = True
                    Exit For
                End If
            Next
            If Yastaba = True Then Continue For
            xDs.Rows.Add({filterDt.Rows(i).Item(2)})
        Next

        For i = 0 To xDs.Rows.Count - 1
            ListView1.Items.Add(xDs.Rows(i).Item(0), "Rule " & CStr(i + 1), 0)
        Next

    End Sub

    Private Sub Form7_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        Call CargaListaDeReglas()

    End Sub

    Private Sub ListView1_ItemChecked(sender As Object, e As ItemCheckedEventArgs) Handles ListView1.ItemChecked


        DataGridView1.Rows.Clear()

        If e.Item.Checked = True Then

            For i = 0 To ListView1.Items.Count - 1
                If i <> e.Item.Index Then
                    ListView1.Items(i).Checked = False
                End If
            Next

            reglaSelected = e.Item.Name

            ToolStripLabel1.Text = e.Item.Name

            Dim filterDt As DataTable = BuildDs.Tables(posiTabla).Clone()
            Dim result() As DataRow = BuildDs.Tables(posiTabla).Select("TableCode = '" & ConsTable & "' AND FieldCode='" & ConsField & "' AND RuleKey='" & e.Item.Name & "'", "Consec ASC")
            For Each row As DataRow In result
                filterDt.ImportRow(row)
            Next


            'si es un catálogo, entonces lo buscamos en el set de catálogos!
            Dim xObj As Object = Nothing
            Dim posiTab As Integer = 0

            'Que guarde la ruta del catálogo: gb#gb0007, por ejemplo!

            EstoyAgregandoRows = True

            For i = 0 To filterDt.Rows.Count - 1
                DataGridView1.Rows.Add()
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = filterDt.Rows(i).Item(4)
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = filterDt.Rows(i).Item(5)


                Select Case filterDt.Rows(i).Item(5)
                    Case Is = "A - From Catalog"
                        'posiTab = MisCatas.Tables.IndexOf(CStr(filterDt.Rows(i).Item(6)))
                        'If posiTab >= 0 Then
                        '    xObj = Split(CStr(filterDt.Rows(i).Item(6)), "#")
                        '    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = CStr(xObj(0)) & " " & MisCatas.Tables(posiTab).Columns(0).ColumnName
                        'End If
                        If CStr(filterDt.Rows(i).Item(6)) = "" Then
                            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = "None"
                        Else
                            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = CStr(filterDt.Rows(i).Item(6))
                        End If
                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).ReadOnly = True

                    Case Is = "B - Free Text"
                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = filterDt.Rows(i).Item(6)
                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).ReadOnly = False

                    Case Is = "C - Running number"
                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = filterDt.Rows(i).Item(6)
                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).ReadOnly = False

                    Case Is = "D - Fixed Value"
                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = filterDt.Rows(i).Item(6)
                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).ReadOnly = False

                    Case Is = "E - External object"
                        If CStr(filterDt.Rows(i).Item(6)) = "" Then
                            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = "None"
                        Else
                            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = CStr(filterDt.Rows(i).Item(6))
                        End If
                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).ReadOnly = True

                End Select

                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = filterDt.Rows(i).Item(7)
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = filterDt.Rows(i).Item(8)

            Next

            EstoyAgregandoRows = False

        Else
            reglaSelected = ""
            ToolStripLabel1.Text = "Ready"
        End If

    End Sub

    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing

        Select Case DataGridView1.CurrentCell.ColumnIndex
            Case Is = 0


            Case Is = 1


            Case Is = 2 'value

                Dim autoText As TextBox = CType(e.Control, TextBox)
                If DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(1).Value = "A - From Catalog" Then
                    If (autoText IsNot Nothing) Then
                        autoText.AutoCompleteCustomSource = xSource
                        autoText.AutoCompleteMode = AutoCompleteMode.Suggest
                        autoText.AutoCompleteSource = AutoCompleteSource.CustomSource
                    End If
                Else
                    If (autoText IsNot Nothing) Then
                        autoText.AutoCompleteMode = AutoCompleteMode.None
                        autoText.AutoCompleteSource = AutoCompleteSource.CustomSource
                    End If
                End If

            Case Is = 3



            Case Is = 4



        End Select

    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click

        DataGridView1.Rows.Add()
        For i = 0 To DataGridView1.Rows.Count - 1
            DataGridView1.Rows(i).Cells(0).Value = CStr(i + 1)
        Next

        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = "No selection"
        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = ""
        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = ""
        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = ""

    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click

    End Sub

    Private Async Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click

        If reglaSelected = "" Then
            MsgBox("Please select a rule first!!", vbCritical, TitBox)
            Exit Sub
        End If

        If DataGridView1.Rows.Count = 0 Then
            MsgBox("Please add some rules first!", vbCritical, TitBox)
            Exit Sub
        End If


        Dim xDs As New DataSet
        xDs.Tables.Add()
        xDs.Tables(0).Columns.Add("Consec", GetType(Integer)) '0
        xDs.Tables(0).Columns.Add("Rule", GetType(String)) '1
        xDs.Tables(0).Columns.Add("Value", GetType(String)) '2
        xDs.Tables(0).Columns.Add("CharLen", GetType(String)) '3
        xDs.Tables(0).Columns.Add("Signo", GetType(String)) '4

        Dim posHash As Integer = 0
        Dim unEntero1 As Integer
        Dim unEntero2 As Integer = 0
        Dim laIzq As String = ""
        Dim laDer As String = ""
        'Dim enCuentra As DataRow
        Dim miLimit As Integer = 0
        Dim valorColumna As String = ""
        Dim posiColumna As Integer = 0

        For i = 0 To DataGridView1.Rows.Count - 1

            valorColumna = ""

            If DataGridView1.Rows(i).Cells(1).Value = "No selection" Then
                MsgBox("A rule is missing on the row '" & CStr(i + 1) & "', please select an option from the Rule column", vbCritical, TitBox)
                Exit Sub
            End If

            Select Case DataGridView1.Rows(i).Cells(1).Value
                Case Is = "A - From Catalog"
                    'verificamos que el renglon tenga dato en el dataset!
                    If DataGridView1.Rows(i).Cells(2).Value = "" Or DataGridView1.Rows(i).Cells(2).Value = "None" Then
                        MsgBox("Please make sure to set a catalog on row " & CStr(i), vbCritical, TitBox)
                        Exit Sub
                    End If

                    valorColumna = DataGridView1.Rows(i).Cells(2).Value



                Case Is = "B - Free Text"
                    'se pone vacío!
                    DataGridView1.Rows(i).Cells(2).Value = ""


                Case Is = "C - Running number"
                    'que tenga algo?
                    If Integer.TryParse(CStr(DataGridView1.Rows(i).Cells(3).Value), unEntero1) = True Then
                        '
                    Else
                        MsgBox("The max char length must of Running number rule must be integers!!, please check row:'" & CStr(i + 1), vbCritical, TitBox)
                        Exit Sub
                    End If



                Case Is = "D - Fixed Value"
                    'que tenga algo!!
                    If DataGridView1.Rows(i).Cells(2).Value = "" Then
                        MsgBox("Please type a Value for this fixed value on row: " & CStr(i + 1), vbCritical, TitBox)
                        Exit Sub
                    End If

                    valorColumna = DataGridView1.Rows(i).Cells(2).Value


                Case Is = "E - External object"
                    If DataGridView1.Rows(i).Cells(2).Value = "" Or DataGridView1.Rows(i).Cells(2).Value = "None" Then
                        MsgBox("Please make sure to set a external object traceability for row " & CStr(i), vbCritical, TitBox)
                        Exit Sub
                    End If
                    valorColumna = DataGridView1.Rows(i).Cells(2).Value

            End Select


            If DataGridView1.Rows(i).Cells(3).Value = "" Then
                MsgBox("Please type a max char length for the row '" & CStr(i + 1), vbCritical, TitBox)
                Exit Sub
            End If

            If CStr(DataGridView1.Rows(i).Cells(3).Value).Length > 5 Then
                MsgBox("The max char length should be at most 5 digits long'" & CStr(i + 1), vbCritical, TitBox)
                Exit Sub
            End If

            If CStr(DataGridView1.Rows(i).Cells(3).Value).Contains("#") = True Then

                posHash = CStr(DataGridView1.Rows(i).Cells(3).Value).IndexOf("#")

                laIzq = CStr(DataGridView1.Rows(i).Cells(3).Value).Substring(0, posHash)
                laDer = CStr(DataGridView1.Rows(i).Cells(3).Value).Substring(posHash + 1)

                If Integer.TryParse(laIzq, unEntero1) = True And Integer.TryParse(laDer, unEntero2) = True Then
                    'ok
                    miLimit = miLimit + CInt(laDer)
                Else
                    MsgBox("The left an right part of the max Char Length must be integers!!, please check row:'" & CStr(i + 1), vbCritical, TitBox)
                    Exit Sub
                End If

            Else

                If IsNumeric(DataGridView1.Rows(i).Cells(3).Value) = False Then
                    MsgBox("The max char length must be integers!!, please check row:'" & CStr(i + 1), vbCritical, TitBox)
                    Exit Sub
                End If

                If Integer.TryParse(CStr(DataGridView1.Rows(i).Cells(3).Value), unEntero1) = True Then
                    'ok!
                    miLimit = miLimit + unEntero1
                Else
                    MsgBox("The max char length must be an integer number, please check row '" & CStr(i + 1), vbCritical, TitBox)
                    Exit Sub
                End If
            End If

            If IsNothing(DataGridView1.Rows(i).Cells(4).Value) = False Then
                If DataGridView1.Rows(i).Cells(4).Value <> "" Then
                    miLimit = miLimit + CStr(DataGridView1.Rows(i).Cells(4).Value).Length
                End If
            End If

            'si hasta aqui todo Ok, lo agregamos!
            xDs.Tables(0).Rows.Add({CStr(DataGridView1.Rows(i).Cells(0).Value), CStr(DataGridView1.Rows(i).Cells(1).Value), valorColumna, CStr(DataGridView1.Rows(i).Cells(3).Value), CStr(DataGridView1.Rows(i).Cells(4).Value)})

        Next

        If miLimit > maxCharLimit Then
            MsgBox("The max char limit you defined for this entire field is " & maxCharLimit & ", and currently this rule has currently " & miLimit, vbCritical, TitBox)
            Exit Sub
        End If

        ToolStripLabel1.Text = "Wait..."
        ToolStrip1.Enabled = False
        ToolStrip2.Enabled = False

        Dim elCamino As String = ""
        elCamino = RaizFire
        elCamino = elCamino & "/construction/" & ConsObject & "/" & ConsTable & "/" & ConsField
        Dim resp As String = ""

        'primero lo borramos!
        'catalog route > gb:gb0001:FIELDMATCH:MATNR#A-BUKRS#B
        'catalog para valor between > gb:gb0001:FIELDMATCH>FIELDMATCH:MATNR#A-BUKRS#B

        resp = Await HazDeleteEnFbSimple(elCamino, reglaSelected)
        If resp = "Ok" Then
            resp = Await HazPostEnFireBaseConPathYColumnas(elCamino, xDs.Tables(0), reglaSelected, -1)

            MsgBox("Rule updated!", vbInformation, TitBox)

            Await ReCargaConstruction()

        End If

        ToolStripLabel1.Text = "Ready"

        ToolStrip1.Enabled = True
        ToolStrip2.Enabled = True


    End Sub

    Private Async Function ReCargaConstruction() As Task(Of String)

        Dim miSet As New DataSet
        miSet.Tables.Add("arbol")
        miSet.Tables(0).Columns.Add("address")
        miSet.Tables(0).Rows.Add({"construction"})
        'usaDataset.Tables(0).Rows.Clear()
        'usaDataset.Tables(0).Rows.Add({"construction"})
        BuildDs.Tables.Clear()
        BuildDs = Await PullUrlWs(miSet, "construction")

        For i = 0 To BuildDs.Tables.Count - 1
            If BuildDs.Tables(i).TableName = ConsObject Then
                posiTabla = i
                Exit For
            End If
        Next

        Return "Ok"

    End Function

    Private Sub DataGridView1_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles DataGridView1.RowsRemoved
        For i = 0 To DataGridView1.Rows.Count - 1
            DataGridView1.Rows(i).Cells(0).Value = CStr(i + 1)
        Next
    End Sub

    Private Async Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        If reglaSelected = "" Then
            MsgBox("Please select first one rule to delete!", vbCritical, TitBox)
            Exit Sub
        End If

        Dim x As Integer = 0

        x = MsgBox("ATTENTION!!" & vbCrLf & "Are you sure you want to delete this rule?, this action can not be undone!", vbQuestion + vbYesNo, TitBox)

        If x <> 6 Then Exit Sub

        Dim elCamino As String = ""
        elCamino = RaizFire
        elCamino = elCamino & "/construction/" & ConsObject & "/" & ConsTable & "/" & ConsField

        Dim elResp As String = ""
        elResp = Await HazDeleteEnFbSimple(elCamino, reglaSelected)

        If elResp = "Ok" Then

            Await ReCargaConstruction()

            Call CargaListaDeReglas()

            reglaSelected = ""

            DataGridView1.Rows.Clear()

        End If


    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick



        Select Case e.ColumnIndex

            Case Is = 1


            Case Is = 2

                'valor!
                Select Case DataGridView1.Rows(e.RowIndex).Cells(1).Value

                    Case Is = "A - From Catalog"
                        'abrimos el form11
                        Dim miTag As Object = DataGridView1.Rows(e.RowIndex).Cells(2).Value
                        If miTag = "None" Then
                            'va directo a load

                            Form11.toyLeyendo = False
                            Form11.resCatCode = ""
                            Form11.resCatModule = ""
                            Form11.resCatMC = ""
                            Form11.resCatMF = ""

                        Else
                            'carga condiciones iniciales
                            Dim xObj As Object = Nothing
                            xObj = Split(miTag, ":")

                            If UBound(xObj) <> 3 Then
                                Form11.toyLeyendo = False
                                Form11.resCatCode = ""
                                Form11.resCatModule = ""
                                Form11.resCatMC = ""
                                Form11.resCatMF = ""
                            Else
                                Form11.toyLeyendo = True
                                Form11.resCatModule = CStr(xObj(0))
                                Form11.resCatCode = CStr(xObj(1))
                                Form11.resCatMF = CStr(xObj(2))
                                Form11.resCatMC = CStr(xObj(3))
                            End If

                        End If

                        Form11.tablaTemp = TablaDatos ' tempDs.Tables(tablaNombre)
                        Form11.objetoModule = ConsModule 'moduloSelek
                        Form11.objetoCode = ConsObject ' objetoSelek
                        Form11.TablaCode = ConsTable ' tableSelek
                        Form11.elCampoCode = ConsField ' CStr(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                        Form11.ultiCats = MisCatas
                        Form11.elEnfoque = "B" 'de construccion
                        Form11.ShowDialog()

                        If Form11.huboExito = False Then Exit Sub

                        DataGridView1.Rows(e.RowIndex).Cells(2).Value = Form11.resCatModule & ":" & Form11.resCatCode & ":" & Form11.resCatMF & ":" & Form11.resCatMC

                    Case Is = "E - External object"

                        Form3.elEnfoque = "B" 'Build!
                        Form3.xtraDs = allTemps
                        Form3.yTraDs = allDepes ' depeDs
                        Form3.depeTemplate = ConsObject ' elNode.Parent.Name
                        Form3.depeTabla = ConsTable ' tableSelek ' elNode.Name

                        Form3.resDepFieldCode = ""
                        Form3.resDepFieldName = ""
                        Form3.resConTempCode = ""
                        Form3.resConTempName = ""
                        Form3.resConTempModule = ""
                        Form3.resConTableCode = ""
                        Form3.resConTableName = ""
                        Form3.resConFieldCode = ""
                        Form3.resConFieldName = ""

                        Form3.resConType = ""
                        Form3.resConRule = ""
                        Form3.resConVal = ""
                        Form3.resMachFields = ""

                        Form3.depeCampo = ConsField ' DataGridView1.Rows(e.RowIndex).Cells(0).Value

                        'dep/objeto/tabla
                        Form3.elPath = "construction/" & ConsObject & "/" & ConsTable

                        Form3.huboExito = False

                        Form3.ShowDialog()

                        If Form3.huboExito = False Then Exit Sub

                        'se pega la estructura!
                        DataGridView1.Rows(e.RowIndex).Cells(2).Value = Form3.resConTempModule & ":" & Form3.resConTempCode & ":" & Form3.resConTableCode & ":" & Form3.resConFieldCode & ":" & Form3.resMachFields & ":" & Form3.resConRule



                End Select


        End Select

    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged

        If EstoyAgregandoRows = True Then Exit Sub

        Select Case e.ColumnIndex

            Case Is = 0


            Case Is = 1

                Select Case DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                    Case Is = "A - From Catalog"
                        DataGridView1.Rows(e.RowIndex).Cells(2).Value = "None"
                        DataGridView1.Rows(e.RowIndex).Cells(2).ReadOnly = True

                    Case Is = "E - External object"
                        DataGridView1.Rows(e.RowIndex).Cells(2).Value = "None"
                        DataGridView1.Rows(e.RowIndex).Cells(2).ReadOnly = True

                    Case Else
                        DataGridView1.Rows(e.RowIndex).Cells(2).ReadOnly = False
                        DataGridView1.Rows(e.RowIndex).Cells(2).Value = ""

                End Select


            Case Is = 2



            Case Is = 3



        End Select

    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click

        If IsNothing(DataGridView1.CurrentCell) = True Then Exit Sub

        DataGridView1.Rows.RemoveAt(DataGridView1.CurrentCell.RowIndex)

        For i = 0 To DataGridView1.Rows.Count - 1
            DataGridView1.Rows(i).Cells(0).Value = CStr(i + 1)
        Next

    End Sub
End Class