Public Class Form9


    Private FuenteCamposChild As New AutoCompleteStringCollection

    Private FuenteCamposParent As New AutoCompleteStringCollection

    Public tabHija As DataTable

    Public tabPapa As DataTable

    Private TablaHijos As New DataTable

    Private TablaPapas As New DataTable

    Public depeField As String

    Public respCadena As String = ""

    Private Sub Form9_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Form para agregar reglas interdependencias del template mismo!

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

        DataGridView1.Columns.Add("DepFieldCode", "Dep Field Code") '1

        DataGridView1.Columns.Add("DepFieldName", "Dep Field Name") '2

        DataGridView1.Columns.Add("MatchFieldCode", "Match Field Code") '3

        DataGridView1.Columns.Add("MatchFieldName", "Match Field Name") '4

        DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns(0).ReadOnly = True

        DataGridView1.Columns(1).ReadOnly = False

        DataGridView1.Columns(2).ReadOnly = True

        DataGridView1.Columns(3).ReadOnly = False

        DataGridView1.Columns(4).ReadOnly = True


        If IsNothing(TablaHijos) = False Then
            TablaHijos.PrimaryKey = Nothing

            TablaHijos.Columns.Clear()

            TablaHijos.Rows.Clear()
        End If


        If IsNothing(TablaPapas) = False Then
            TablaPapas.PrimaryKey = Nothing

            TablaPapas.Columns.Clear()

            TablaPapas.Rows.Clear()
        End If

        Dim iAveprim(1) As DataColumn

        Dim kEys As New DataColumn()

        kEys.ColumnName = "FieldCode"

        iAveprim(0) = kEys

        TablaHijos.Columns.Add(kEys)

        TablaHijos.Columns.Add("NombreField")

        TablaHijos.PrimaryKey = iAveprim


        FuenteCamposChild.Clear()

        For i = 0 To tabHija.Rows.Count - 1

            If depeField = CStr(tabHija.Rows(i).Item(3)) Then Continue For

            FuenteCamposChild.Add(tabHija.Rows(i).Item(3))

            TablaHijos.Rows.Add({tabHija.Rows(i).Item(3), tabHija.Rows(i).Item(4)})

        Next


        Dim YavePrim(1) As DataColumn

        Dim Yaves As New DataColumn()

        Yaves.ColumnName = "FieldCode"

        YavePrim(0) = Yaves

        TablaPapas.Columns.Add(Yaves)

        TablaPapas.Columns.Add("NombreField")

        TablaPapas.PrimaryKey = YavePrim

        For i = 0 To tabPapa.Rows.Count - 1
            FuenteCamposParent.Add(tabPapa.Rows(i).Item(3))
            TablaPapas.Rows.Add({tabPapa.Rows(i).Item(3), tabPapa.Rows(i).Item(4)})
        Next

        respCadena = "None"

        Me.CenterToScreen()

    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click

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

        Dim conStruye As String = ""

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

            enCuentra = TablaHijos.Rows.Find(CStr(DataGridView1.Rows(i).Cells(1).Value))
            If IsNothing(enCuentra) = True Then
                MsgBox("The row " & CStr(i + 1) & " does not contain a valid Child Field Code, please verify!!!", vbCritical, "DQCT")
                Exit Sub
            End If

            enCuentra = TablaPapas.Rows.Find(CStr(DataGridView1.Rows(i).Cells(3).Value))
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

            If i <> 0 Then conStruye = conStruye & "-"

            conStruye = conStruye & CStr(DataGridView1.Rows(i).Cells(1).Value) & "#" & CStr(DataGridView1.Rows(i).Cells(3).Value)

        Next

        If conStruye = "" Then conStruye = "None"

        respCadena = conStruye

        Me.Close()

    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        DataGridView1.Rows.Add()

        For i = 0 To DataGridView1.Rows.Count - 1

            DataGridView1.Rows(i).Cells(0).Value = CStr(i + 1)

        Next
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
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

    Private Sub DataGridView1_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles DataGridView1.CurrentCellDirtyStateChanged
        If DataGridView1.IsCurrentCellDirty Then

            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)

        End If
    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        Dim enCuentra As DataRow

        Dim z As Integer

        Select Case e.ColumnIndex

            Case Is = 1

                enCuentra = TablaHijos.Rows.Find(CStr(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value))

                If IsNothing(enCuentra) = False Then

                    z = TablaHijos.Rows.IndexOf(enCuentra)

                    DataGridView1.Rows(e.RowIndex).Cells(2).Value = TablaHijos.Rows(z).Item(1)

                Else

                    DataGridView1.Rows(e.RowIndex).Cells(2).Value = ""

                End If

            Case Is = 3

                enCuentra = TablaPapas.Rows.Find(CStr(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value))

                If IsNothing(enCuentra) = False Then

                    z = TablaPapas.Rows.IndexOf(enCuentra)

                    DataGridView1.Rows(e.RowIndex).Cells(4).Value = TablaPapas.Rows(z).Item(1)

                Else

                    DataGridView1.Rows(e.RowIndex).Cells(4).Value = ""

                End If

        End Select
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
End Class