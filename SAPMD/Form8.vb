Public Class Form8
    Public queObjetoEs As String

    Public xTempDs As DataSet

    Public relaDs As DataSet

    Private tablaLocal As New DataTable

    Private relaLocal As New DataTable

    Private FuenteCamposChild As New AutoCompleteStringCollection

    Private FuenteCamposParent As New AutoCompleteStringCollection

    Private TablaChilds As New DataTable

    Private TablaParents As New DataTable

    Private tableSelekted As String = ""

    Private tableNameSelekted As String = ""

    Private parentTable As String = ""

    Private parentTableName As String = ""

    Private reglaSelekted As String = ""


    Private Sub Form8_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Form para Agregar parent/child interno
        'Cuando sea de una mismo objeto, validar que una combinacion de campos exista entre todas las tablas!
        'Cuando la combinacion de campos es solo 1, es similar a la dependencia de objetos singular!
        ListView1.Items.Clear()

        ListView1.View = View.Details

        ListView1.CheckBoxes = True

        ListView1.SmallImageList = ImageList1

        ListView1.Columns.Clear()

        ListView1.Columns.Add("#", 50, HorizontalAlignment.Center)

        ListView1.Columns.Add("Rule Name", 70, HorizontalAlignment.Center)

        ListView1.Columns.Add("No. of Fields", 60, HorizontalAlignment.Center)

        Call RecargaBinds()

        ToolStripLabel1.Text = "Ready"

        DataGridView1.ReadOnly = False



        DataGridView1.EnableHeadersVisualStyles = False



        DataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 64, 114, 196) ' System.Drawing.Color.FromArgb(228, 109, 10)

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



        DataGridView1.Columns.Add("Row", "Row") '0

        DataGridView1.Columns.Add("ChildFieldCode", "Child Field Code") '1

        DataGridView1.Columns.Add("Child Field Name", "Child Field Name") '2

        DataGridView1.Columns.Add("ParentFieldCode", "Parent Field Code") '3

        DataGridView1.Columns.Add("ParentFieldName", "Parent Field Name") '4

        DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns(0).ReadOnly = True

        DataGridView1.Columns(1).ReadOnly = True

        DataGridView1.Columns(2).ReadOnly = True

        DataGridView1.Columns(3).ReadOnly = True

        DataGridView1.Columns(4).ReadOnly = True



        If tablaLocal.Rows.Count = 0 Then

            MsgBox("No table objects found!!", vbInformation, "SAP MD")

            Me.Close()

        End If

        tableSelekted = ""

        tableNameSelekted = ""

        parentTable = ""

        parentTableName = ""

        reglaSelekted = ""

        ToolStripComboBox2.Enabled = False
        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = False
        ToolStripButton5.Enabled = False
        ToolStripButton6.Enabled = False

        Label1.Text = "No table selected"

        Call FillComboConTables()

        Me.CenterToScreen()

    End Sub

    Private Sub RecargaBinds()

        If IsNothing(tablaLocal) = False Then
            tablaLocal.Columns.Clear()

            tablaLocal.Rows.Clear()
        End If

        If IsNothing(relaLocal) = False Then
            relaLocal.Columns.Clear()

            relaLocal.Rows.Clear()

        End If


        Dim i As Integer = 0

        Dim xObj As Object = Nothing

        For i = 0 To xTempDs.Tables.Count - 1

            xObj = Split(xTempDs.Tables(i).TableName, "#")

            If CStr(xObj(0)) = queObjetoEs Then

                tablaLocal = xTempDs.Tables(i).DefaultView.ToTable()

                Exit For

            End If

        Next

        Dim poSi As Integer = 0

        poSi = relaDs.Tables.IndexOf(queObjetoEs)

        If poSi >= 0 Then

            relaLocal = relaDs.Tables(poSi)

        End If

    End Sub

    Private Sub FillComboConTables()


        ToolStripComboBox1.Items.Clear()

        ToolStripComboBox2.Items.Clear()

        ToolStripComboBox1.Items.Add("Child table")

        For i = 0 To tablaLocal.Rows.Count - 1

            If ToolStripComboBox1.Items.IndexOf(CStr(tablaLocal.Rows(i).Item(1))) >= 0 Then Continue For

            ToolStripComboBox1.Items.Add(CStr(tablaLocal.Rows(i).Item(1)))

        Next

        ToolStripComboBox1.SelectedIndex = 0

    End Sub

    Private Sub ToolStripComboBox1_Click(sender As Object, e As EventArgs) Handles ToolStripComboBox1.Click

    End Sub

    Private Sub ToolStripComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        If ToolStripComboBox1.SelectedIndex = 0 Then

            'se deshabilita todo!

            ListView1.Items.Clear()

            DataGridView1.Rows.Clear()

            ToolStripComboBox2.Items.Clear()

            ToolStripComboBox2.Items.Add("Parent table")

            DataGridView1.Columns(1).ReadOnly = True

            DataGridView1.Columns(3).ReadOnly = True

            tableSelekted = ""
            tableNameSelekted = ""
            parentTable = ""
            parentTableName = ""
            reglaSelekted = ""

        Else

            FuenteCamposChild.Clear()

            If IsNothing(TablaChilds) = False Then
                TablaChilds.PrimaryKey = Nothing

                TablaChilds.Columns.Clear()

                TablaChilds.Rows.Clear()
            End If

            If IsNothing(TablaParents) = False Then
                TablaParents.PrimaryKey = Nothing

                TablaParents.Columns.Clear()

                TablaParents.Rows.Clear()
            End If


            Dim iAveprim(1) As DataColumn

            Dim kEys As New DataColumn()

            kEys.ColumnName = "FieldCode"

            iAveprim(0) = kEys

            TablaChilds.Columns.Add(kEys)

            TablaChilds.Columns.Add("NombreField")

            TablaChilds.PrimaryKey = iAveprim

            Dim tablaX As String = ""

            Dim filterDt As DataTable

            tablaX = ToolStripComboBox1.Items(ToolStripComboBox1.SelectedIndex)

            tableSelekted = tablaX

            filterDt = tablaLocal.Clone()

            Dim result() As DataRow = tablaLocal.Select("TableCode = '" & tablaX & "'")

            For Each row As DataRow In result

                filterDt.ImportRow(row)

            Next

            For i = 0 To filterDt.Rows.Count - 1

                FuenteCamposChild.Add(filterDt.Rows(i).Item(3))

                TablaChilds.Rows.Add({filterDt.Rows(i).Item(3), filterDt.Rows(i).Item(4)})

                tableNameSelekted = CStr(filterDt.Rows(i).Item(2))

            Next

            DataGridView1.Rows.Clear()
            ToolStripComboBox2.Enabled = False
            ToolStripButton3.Enabled = False
            ToolStripButton4.Enabled = False
            ToolStripButton5.Enabled = False
            ToolStripButton6.Enabled = False

            Label1.Text = "Child table " & tableSelekted & " / " & tableNameSelekted

            DataGridView1.Columns(1).ReadOnly = False

            Call ReCargaListaDeListView(ToolStripComboBox1.Items(ToolStripComboBox1.SelectedIndex))

            Call RecargaComboPadres()

        End If

    End Sub

    Private Sub RecargaComboPadres()

        ToolStripComboBox2.Items.Clear()

        ToolStripComboBox2.Items.Add("Parent table")

        For i = 0 To tablaLocal.Rows.Count - 1

            If CStr(tablaLocal.Rows(i).Item(1)) = tableSelekted Then Continue For ' no puedes seleccionar de la misma tabla hija!

            If ToolStripComboBox2.Items.IndexOf(CStr(tablaLocal.Rows(i).Item(1))) >= 0 Then Continue For

            ToolStripComboBox2.Items.Add(CStr(tablaLocal.Rows(i).Item(1)))

        Next

        ToolStripComboBox2.SelectedIndex = 0

    End Sub

    Private Sub ReCargaListaDeListView(ByVal laTablaCode As String)

        ListView1.Items.Clear()

        If relaLocal.Columns.Count = 0 Then Exit Sub

        'se selecciona el codigo de la tabla que tiene la info!

        Dim filterDt As DataTable = relaLocal.Clone()

        Dim result() As DataRow = relaLocal.Select("ChildTableCode = '" & laTablaCode & "'")

        For Each row As DataRow In result

            filterDt.ImportRow(row)

        Next

        'Dim uniqueRules = filterDt.DefaultView.ToTable(True, "RuleCode") 'poner el nombre de la columna!!

        'aqui se debe seleccionar las reglas!!

        ListView1.Tag = laTablaCode
        Dim k As Integer = 0
        For i = 0 To filterDt.Rows.Count - 1

            If ListView1.Items.IndexOfKey(CStr(filterDt.Rows(i).Item(1))) >= 0 Then Continue For
            k = k + 1
            ListView1.Items.Add(CStr(filterDt.Rows(i).Item(1)), CStr(k), 0) 'posicion de la columna donde esta el nombre de la regla!
            ListView1.Items(ListView1.Items.Count - 1).SubItems.Add(CStr(filterDt.Rows(i).Item(1)))
            ListView1.Items(ListView1.Items.Count - 1).SubItems.Add(CStr(filterDt.Rows(i).Item(3)))
        Next

    End Sub

    Private Sub ListView1_ItemChecked(sender As Object, e As ItemCheckedEventArgs) Handles ListView1.ItemChecked
        If e.Item.Checked = False Then

            DataGridView1.Rows.Clear()

            reglaSelekted = ""

            parentTable = ""

            parentTableName = ""

            ToolStripComboBox2.Enabled = False
            ToolStripButton3.Enabled = False
            ToolStripButton4.Enabled = False
            ToolStripButton5.Enabled = False
            ToolStripButton6.Enabled = False

            ToolStripLabel2.Text = "No parent selected"

            Exit Sub

        End If

        For i = 0 To ListView1.Items.Count - 1

            If i <> e.Item.Index Then

                ListView1.Items(i).Checked = False

            End If

        Next

        reglaSelekted = e.Item.Name
        parentTable = ""
        parentTableName = ""
        Call MuestraRelaciones(ListView1.Tag, e.Item.Name)

    End Sub

    Private Sub MuestraRelaciones(ByVal laTablaCode As String, ByVal laReglaYave As String)

        DataGridView1.Rows.Clear()

        'filtramos de la tabla de relaciones global!

        Dim filterDt As DataTable = relaLocal.Clone()

        Dim result() As DataRow = relaLocal.Select("ChildTableCode = '" & laTablaCode & "' AND RuleCode='" & laReglaYave & "'") 'AND NumberOfColumns > 0

        For Each row As DataRow In result

            filterDt.ImportRow(row)

        Next

        For i = 0 To filterDt.Rows.Count - 1

            'ParentTableCode
            parentTable = filterDt.Rows(i).Item("ParentTableCode")
            If filterDt.Rows(i).Item("ChildField") = "" Then Continue For
            If filterDt.Rows(i).Item("ParentField") = "" Then Continue For
            DataGridView1.Rows.Add({CStr(i + 1), filterDt.Rows(i).Item("ChildField"), filterDt.Rows(i).Item("ChildFieldName"), filterDt.Rows(i).Item("ParentField"), filterDt.Rows(i).Item("ParentFieldName")})

        Next

        If parentTable = "" Then
            ToolStripButton3.Enabled = False
            ToolStripButton4.Enabled = False
            ToolStripButton5.Enabled = False

            ToolStripComboBox2.SelectedIndex = 0
            ToolStripComboBox2.Enabled = True
            ToolStripButton6.Enabled = True

        Else
            ToolStripButton3.Enabled = True
            ToolStripButton4.Enabled = True
            ToolStripButton5.Enabled = True

            ToolStripComboBox2.Enabled = False
            ToolStripButton6.Enabled = False

            ToolStripComboBox2.SelectedItem = CStr(parentTable)

        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        Select Case DataGridView1.CurrentCell.ColumnIndex

            Case Is = 1

                Dim autoText As TextBox = CType(e.Control, TextBox)

                If (autoText IsNot Nothing) Then

                    autoText.AutoCompleteCustomSource = FuenteCamposChild

                    autoText.AutoCompleteMode = AutoCompleteMode.Suggest

                    autoText.AutoCompleteSource = AutoCompleteSource.CustomSource

                End If

            Case Is = 3

                Dim autoText As TextBox = CType(e.Control, TextBox)

                If (autoText IsNot Nothing) Then

                    autoText.AutoCompleteCustomSource = FuenteCamposParent

                    autoText.AutoCompleteMode = AutoCompleteMode.Suggest

                    autoText.AutoCompleteSource = AutoCompleteSource.CustomSource

                End If

        End Select
    End Sub

    Private Sub ToolStripComboBox2_Click(sender As Object, e As EventArgs) Handles ToolStripComboBox2.Click

    End Sub

    Private Sub ToolStripComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ToolStripComboBox2.SelectedIndexChanged

        FuenteCamposParent.Clear()

        If ToolStripComboBox2.SelectedIndex = 0 Then

            DataGridView1.Columns(3).ReadOnly = True

            Exit Sub

        End If

        TablaParents.PrimaryKey = Nothing

        TablaParents.Columns.Clear()

        TablaParents.Rows.Clear()



        Dim iAveprim(1) As DataColumn

        Dim kEys As New DataColumn()

        kEys.ColumnName = "FieldCode"

        iAveprim(0) = kEys

        TablaParents.Columns.Add(kEys)

        TablaParents.Columns.Add("NombreField")

        TablaParents.PrimaryKey = iAveprim

        Dim tablaX As String = ""

        Dim filterDt As DataTable

        tablaX = ToolStripComboBox2.Items(ToolStripComboBox2.SelectedIndex)

        parentTable = tablaX

        filterDt = tablaLocal.Clone()

        Dim result() As DataRow = tablaLocal.Select("TableCode = '" & tablaX & "'")

        For Each row As DataRow In result

            filterDt.ImportRow(row)

        Next

        For i = 0 To filterDt.Rows.Count - 1

            parentTableName = CStr(filterDt.Rows(i).Item(2))

            FuenteCamposParent.Add(filterDt.Rows(i).Item(3))

            TablaParents.Rows.Add({filterDt.Rows(i).Item(3), filterDt.Rows(i).Item(4)})

        Next

        ToolStripLabel2.Text = parentTableName

        DataGridView1.Columns(3).ReadOnly = False

    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        Dim enCuentra As DataRow

        Dim z As Integer

        Select Case e.ColumnIndex

            Case Is = 1

                enCuentra = TablaChilds.Rows.Find(CStr(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value))

                If IsNothing(enCuentra) = False Then

                    z = TablaChilds.Rows.IndexOf(enCuentra)

                    DataGridView1.Rows(e.RowIndex).Cells(2).Value = TablaChilds.Rows(z).Item(1)

                Else

                    DataGridView1.Rows(e.RowIndex).Cells(2).Value = ""

                End If

            Case Is = 3

                enCuentra = TablaParents.Rows.Find(CStr(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value))

                If IsNothing(enCuentra) = False Then

                    z = TablaParents.Rows.IndexOf(enCuentra)

                    DataGridView1.Rows(e.RowIndex).Cells(4).Value = TablaParents.Rows(z).Item(1)

                Else

                    DataGridView1.Rows(e.RowIndex).Cells(4).Value = ""

                End If

        End Select
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        DataGridView1.Rows.Add()

        For i = 0 To DataGridView1.Rows.Count - 1

            DataGridView1.Rows(i).Cells(0).Value = CStr(i + 1)

        Next
    End Sub

    Private Sub DataGridView1_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles DataGridView1.CurrentCellDirtyStateChanged
        If DataGridView1.IsCurrentCellDirty Then

            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)

        End If
    End Sub

    Private Async Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        'add new rule!

        If ToolStripComboBox1.SelectedIndex < 0 Then
            MsgBox("Please select a child table first!!", vbCritical, "DQCT")
            Exit Sub
        End If

        If tableSelekted = "" Then
            MsgBox("Please select a child table first!!", vbCritical, "DQCT")
            Exit Sub
        End If

        Dim niuRul As String = ""

        niuRul = InputBox("New relation for " & tableSelekted, "Please type the name of this new relation", "")

        If niuRul = "" Then
            MsgBox("Please type a name for this new rule!!", vbCritical, "DQCT")
            Exit Sub
        End If

        If niuRul.Length <= 4 Then
            MsgBox("Please type name longer than 4 characters!!", vbCritical, "DQCT")
            Exit Sub
        End If

        If niuRul.Contains(".") = True Then
            MsgBox("The rule name must NOT contain special characters, please type another one!", vbCritical, "DQCT")
            Exit Sub
        End If

        'verificar que NO exista previamente el mismo nombre de esta tabla
        If ListView1.Items.IndexOfKey(niuRul) >= 0 Then
            MsgBox("This name for this relation already exists!!, please type another one!", vbCritical, "DQCT")
            Exit Sub
        End If

        Dim misDatos As New DataTable

        misDatos.Columns.Add("Yave")
        misDatos.Columns.Add("NumberOfColumns")
        misDatos.Columns.Add("ParentTable")
        misDatos.Columns.Add("RuleName")

        Dim resP As String = ""
        Dim elCamino As String = ""
        elCamino = RaizFire
        elCamino = elCamino & "/relations"
        elCamino = elCamino & "/" & queObjetoEs
        elCamino = elCamino & "/" & tableSelekted

        misDatos.Rows.Add({niuRul, 0, "", niuRul})

        resP = Await HazPutEnFireBasePathYColumnas(elCamino, misDatos, 0)

        Await ReCargaRelaciones()

        RecargaBinds()

        Call ReCargaListaDeListView(tableSelekted)

        MsgBox(resP, vbInformation, "DQCT")
    End Sub


    Private Async Function ReCargaRelaciones() As Task(Of String)

        Dim miSet As New DataSet
        miSet.Tables.Add()
        miSet.Tables(0).Columns.Add()
        miSet.Tables(0).Rows.Clear()
        miSet.Tables(0).Rows.Add({"relations"})

        relaDs.Tables.Clear()
        relaDs = Await PullUrlWs(miSet, "relations")

        Return "ok"

    End Function

    Private Async Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click

        If reglaSelekted = "" Then
            MsgBox("Please select a relation key first!!", vbCritical, "DQCT")
            Exit Sub
        End If


        If ToolStripComboBox2.SelectedIndex = 0 Then
            MsgBox("Please select a parent table first!!", vbCritical, "DQCT")
            Exit Sub
        End If

        'If ToolStripComboBox2.selectedindex0
        If parentTable = "" Then
            MsgBox("Please select a parent table first!!", vbCritical, "DQCT")
            Exit Sub
        End If

        'hacemos un put!
        Dim laResp As String = ""
        Dim elCamino As String = ""
        elCamino = RaizFire
        elCamino = elCamino & "/relations"
        elCamino = elCamino & "/" & queObjetoEs
        elCamino = elCamino & "/" & tableSelekted
        elCamino = elCamino & "/" & reglaSelekted
        laResp = Await HazPutEnFbSimple(elCamino, "ParentTable", parentTable)

        If laResp = "Ok" Then
            'hacemos reload!
            Await ReCargaRelaciones()
            RecargaBinds()

            'habilitamos el combo!
            ToolStripButton3.Enabled = True
            ToolStripButton4.Enabled = True
            ToolStripButton5.Enabled = True

            ToolStripComboBox2.Enabled = False
            ToolStripButton6.Enabled = False

            'ToolStripComboBox2.SelectedItem = CStr(parentTable)

        Else
            MsgBox(laResp, vbCritical, "DQCT")
        End If

    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click

        If reglaSelekted = "" Then
            MsgBox("Please select a relation key first!!", vbCritical, "DQCT")
            Exit Sub
        End If


        If ToolStripComboBox2.SelectedIndex = 0 Then
            MsgBox("Please select a parent table first!!", vbCritical, "DQCT")
            Exit Sub
        End If

        'If ToolStripComboBox2.selectedindex0
        If parentTable = "" Then
            MsgBox("Please select a parent table first!!", vbCritical, "DQCT")
            Exit Sub
        End If

        If IsNothing(DataGridView1.CurrentCell) = True Then
            MsgBox("Please select a row first!!", vbCritical, "DQCT")
            Exit Sub
        End If

        If DataGridView1.CurrentCell.RowIndex < 0 Then
            MsgBox("Please select a row first!!", vbCritical, "DQCT")
            Exit Sub
        End If

        DataGridView1.Rows.RemoveAt(DataGridView1.CurrentCell.RowIndex)

        For i = 0 To DataGridView1.Rows.Count - 1
            DataGridView1.Rows(i).Cells(0).Value = CStr(i + 1)
        Next

    End Sub

    Private Async Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click

        If ToolStripComboBox1.SelectedIndex < 0 Then
            MsgBox("Please select a child table first!!", vbCritical, "DQCT")
            Exit Sub
        End If

        If tableSelekted = "" Then
            MsgBox("Please select a child table first!!", vbCritical, "DQCT")
            Exit Sub
        End If

        If reglaSelekted = "" Then
            MsgBox("Please select a relation rule first!!", vbCritical, "DQCT")
            Exit Sub
        End If

        Dim x As Integer = 0
        x = MsgBox("ATTENTION!" & vbCrLf & "Are you sure you want to delete this rule relationship?, this action can not be undone!!", vbExclamation + vbYesNo, "DQCT")

        If x <> 6 Then Exit Sub

        Dim elCamino As String = ""
        Dim laResp As String = ""
        elCamino = RaizFire
        elCamino = elCamino & "/relations"
        elCamino = elCamino & "/" & queObjetoEs
        elCamino = elCamino & "/" & tableSelekted

        laResp = Await HazDeleteEnFbSimple(elCamino, reglaSelekted)

        If laResp = "Ok" Then

            'hacemos reload!
            Await ReCargaRelaciones()

            RecargaBinds()

            Call ReCargaListaDeListView(tableSelekted)

            DataGridView1.Rows.Clear()

            reglaSelekted = ""

            parentTable = ""

            parentTableName = ""

            ToolStripComboBox2.Enabled = False
            ToolStripButton3.Enabled = False
            ToolStripButton4.Enabled = False
            ToolStripButton5.Enabled = False
            ToolStripButton6.Enabled = False

            ToolStripLabel2.Text = "No parent selected"

            MsgBox("Rule gone!", vbInformation, "DQCT")

        Else
            MsgBox(laResp, vbInformation, "DQCT")
        End If

    End Sub

    Private Async Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click

        'validar los campos!!
        If tableSelekted = "" Then
            MsgBox("Please select a child table first!!", vbCritical, "DQCT")
            Exit Sub
        End If

        If reglaSelekted = "" Then
            MsgBox("Please select a rule first!!", vbCritical, "DQCT")
            Exit Sub
        End If

        If parentTable = "" Then
            MsgBox("Please select a parent table first!!", vbCritical, "DQCT")
            Exit Sub
        End If

        If DataGridView1.Rows.Count = 0 Then
            MsgBox("Please add at least 1 row relation!!", vbCritical, "DQCT")
            Exit Sub
        End If

        'OJO NO puede haber repeticiones de campos en los renglones!
        'y los campos SI deben existir en las fuentes de fuentechilds y fuenteparents

        Dim miTab As New DataTable
        miTab.Columns.Add("Consec") '0
        miTab.Columns.Add("ChildField") '1
        miTab.Columns.Add("ChildFieldName") '2
        miTab.Columns.Add("ParentField") '3
        miTab.Columns.Add("ParentFieldName") '4

        Dim xTab As New DataTable
        Dim iAveprim(1) As DataColumn
        Dim kEys As New DataColumn()

        kEys.ColumnName = "ChildField"

        iAveprim(0) = kEys

        xTab.Columns.Add(kEys)

        xTab.PrimaryKey = iAveprim


        Dim yTab As New DataTable
        Dim YavePrim(1) As DataColumn
        Dim Yaves As New DataColumn()

        Yaves.ColumnName = "ParentField"

        YavePrim(0) = Yaves

        yTab.Columns.Add(Yaves)

        yTab.PrimaryKey = YavePrim

        Dim z As Integer = 0
        Dim enCuentra As DataRow
        For i = 0 To DataGridView1.Rows.Count - 1
            'validamos que haya algo en la columnas de los códigos!

            If IsNothing(DataGridView1.Rows(i).Cells(1).Value) = True Then
                MsgBox("The row " & CStr(i + 1) & " is missing the 'Child Field Code', please fill it accordingly!!", vbCritical, "DQCT")
                Exit Sub
            End If

            If IsNothing(DataGridView1.Rows(i).Cells(3).Value) = True Then
                MsgBox("The row " & CStr(i + 1) & " is missing the 'Parent Field Code', please fill it accordingly!!", vbCritical, "DQCT")
                Exit Sub
            End If

            If DataGridView1.Rows(i).Cells(1).Value = "" Then
                MsgBox("The row " & CStr(i + 1) & " is missing the 'Child Field Code', please fill it accordingly!!", vbCritical, "DQCT")
                Exit Sub
            End If

            If DataGridView1.Rows(i).Cells(3).Value = "" Then
                MsgBox("The row " & CStr(i + 1) & " is missing the 'Parent Field Code', please fill it accordingly!!", vbCritical, "DQCT")
                Exit Sub
            End If

            'verificar que exista!

            enCuentra = TablaChilds.Rows.Find(CStr(DataGridView1.Rows(i).Cells(1).Value))
            If IsNothing(enCuentra) = True Then
                MsgBox("The row " & CStr(i + 1) & " does not contain a valid Child Field Code, please verify!!!", vbCritical, "DQCT")
                Exit Sub
            End If

            enCuentra = TablaParents.Rows.Find(CStr(DataGridView1.Rows(i).Cells(3).Value))
            If IsNothing(enCuentra) = True Then
                MsgBox("The row " & CStr(i + 1) & " does not contain a valid Parent Field Code, please verify!!!", vbCritical, "DQCT")
                Exit Sub
            End If

            enCuentra = xTab.Rows.Find(CStr(DataGridView1.Rows(i).Cells(1).Value))
            If IsNothing(enCuentra) = True Then
                'lo agrego, esta OK
                xTab.Rows.Add({CStr(DataGridView1.Rows(i).Cells(1).Value)})
            Else
                z = xTab.Rows.IndexOf(enCuentra)
                MsgBox("The row " & CStr(i + 1) & " has a duplicate record of the Child Field Code on the row: " & CStr(z + 1) & ", please verify!", vbCritical, "DQCT")
                Exit Sub
            End If

            enCuentra = yTab.Rows.Find(CStr(DataGridView1.Rows(i).Cells(3).Value))
            If IsNothing(enCuentra) = True Then
                yTab.Rows.Add({CStr(DataGridView1.Rows(i).Cells(3).Value)})
            Else
                MsgBox("The row " & CStr(i + 1) & " has a duplicate record of the Parent Field Code on the row " & CStr(z + 1) & ", please verify!", vbCritical, "DQCT")
                Exit Sub
            End If

            miTab.Rows.Add({CStr(i + 1), CStr(DataGridView1.Rows(i).Cells(1).Value), CStr(DataGridView1.Rows(i).Cells(2).Value), CStr(DataGridView1.Rows(i).Cells(3).Value), CStr(DataGridView1.Rows(i).Cells(4).Value)})

        Next

        Dim unResp As String = ""

        Dim elCamino As String = ""
        elCamino = RaizFire
        'OJO, primero debemos borrar o que tenia antes!
        elCamino = elCamino & "/relations"
        elCamino = elCamino & "/" & queObjetoEs
        elCamino = elCamino & "/" & tableSelekted
        elCamino = elCamino & "/" & reglaSelekted

        unResp = Await HazDeleteEnFbSimple(elCamino, "logic")
        If unResp = "Ok" Then

            'colocamos el numero de columnas!!

            elCamino = RaizFire
            elCamino = elCamino & "/relations"
            elCamino = elCamino & "/" & queObjetoEs
            elCamino = elCamino & "/" & tableSelekted
            elCamino = elCamino & "/" & reglaSelekted

            'colocamos el numero de columnas!!
            unResp = Await HazPutEnFbSimple(elCamino, "NumberOfColumns", CStr(miTab.Rows.Count))

            unResp = Await HazPostEnFireBaseConPathYColumnas(elCamino, miTab, "logic", -1)

            'hacemos reload!
            Await ReCargaRelaciones()

            RecargaBinds()

            MsgBox("Relation updated!!", vbInformation, "DQCT")

        End If


    End Sub

End Class