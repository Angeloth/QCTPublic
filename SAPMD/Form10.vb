Public Class Form10

    Public huboCambio As Boolean
    Public moduObjeto As String
    Public tablaCodigo As String

    'va a tener la estructura:

    'Columna 1: FieldCode

    'Columna 2: FieldName

    'Columna 3: Description

    'Columna 4: Position

    'Columna 5: isText
    Private Sub Form10_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Form para agregar reglas de condicion para campos
        DataGridView1.Columns.Clear()
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

        Dim FillRule As New DataGridViewComboBoxColumn

        FillRule.Name = "FieldType"

        FillRule.HeaderText = "Field Type"

        FillRule.Items.Clear()

        FillRule.Items.Add("Text") '0

        FillRule.Items.Add("Number") '1

        FillRule.DefaultCellStyle.BackColor = Color.White

        FillRule.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox

        FillRule.DisplayStyleForCurrentCellOnly = True

        DataGridView1.Columns.Add("FieldCode", "Field Code") '0

        DataGridView1.Columns.Add("FieldName", "Field Name") '1

        DataGridView1.Columns.Add("Description", "Description") '2

        DataGridView1.Columns.Add(FillRule) '3 numero ó texto

        DataGridView1.Columns(0).ReadOnly = True

        DataGridView1.Columns(1).ReadOnly = False

        DataGridView1.Columns(2).ReadOnly = False

        DataGridView1.Columns(3).ReadOnly = False

        'OJO, siempre mandar una primer columna de campo inicial,

        'Y otra de campo de descripción!

        'NumericUpDown1.Value = TabColumnas.Rows.Count



        Me.CenterToScreen()

        huboCambio = False

    End Sub

    Private Sub CargaRens()

        DataGridView1.Rows.Clear()

        'OJO, acomodar el tabcolumnas pos posicion ascendente!

        TabColumnas.DefaultView.Sort = "Position ASC"

        For i = 0 To TabColumnas.Rows.Count - 1

            DataGridView1.Rows.Add()

            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = CStr(TabColumnas.Rows(i).Item(0))

            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = CStr(TabColumnas.Rows(i).Item(1))

            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = CStr(TabColumnas.Rows(i).Item(2))

            Select Case CStr(TabColumnas.Rows(i).Item(4))

                Case Is = "Text"

                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = CStr(TabColumnas.Rows(i).Item(4))

                Case Is = "Number"

                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = CStr(TabColumnas.Rows(i).Item(4))

            End Select

        Next

        'siempre debe tener mínimo 1 campo llave y el campo de descripción!

        'el campo de descripción final debe generarse automáticamente al copiar / guardar la info!

        'si cambia el nombre de la columna, ó la posición, SIN afectar la cantidad de campos/columnas NO HAY PEX

        'Si elimina una columna, o cambia la cantidad de columnas, OJO, si esto pasa, debe desvincular de todos los templates existentes este catálogo y volverlo a vincular

    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        'add row

        PonMiFieldCode()

    End Sub

    Private Sub PonMiFieldCode()

        Dim tuTab As New DataTable
        AsignaYavePrimariaATabla(tuTab, "Yave", True)
        'buscamos en todos los renglones el valor de su field Code, si YA existe lo omitimos!
        'solo debes asignar un field code de la lista que NO se haya usado antes!, FIN
        Dim i As Integer = 0
        For i = 0 To DataGridView1.Rows.Count - 1
            tuTab.Rows.Add({DataGridView1.Rows(i).Cells(0).Value})
        Next

        DataGridView1.Rows.Add()
        Dim elRen As Integer = DataGridView1.Rows.Count - 1
        Dim enCuentra As DataRow
        Dim anoTher As DataRow
        Dim z As Integer
        Dim yaTengoField As Boolean = False

        i = 1
        Do While yaTengoField = False

            enCuentra = LetraNumero.Tables(0).Rows.Find(i)
            If IsNothing(enCuentra) = False Then
                z = LetraNumero.Tables(0).Rows.IndexOf(enCuentra)
                anoTher = tuTab.Rows.Find(CStr(LetraNumero.Tables(0).Rows(z).Item(1)))
                If IsNothing(anoTher) = True Then
                    'NO encontro el valor, se puede usars!!
                    DataGridView1.Rows(elRen).Cells(0).Value = CStr(LetraNumero.Tables(0).Rows(z).Item(1))
                    yaTengoField = True
                    Exit Do
                Else
                    'Encontró el valor, entonces le seguimos buscando!

                End If

            End If

            i = i + 1
        Loop

        tuTab.Dispose()

    End Sub

    Private Sub ReEscribeFieldCodes()

        Dim enCuentra As DataRow

        Dim z As Integer

        For i = 0 To DataGridView1.Rows.Count - 1

            enCuentra = LetraNumero.Tables(0).Rows.Find(i)

            If IsNothing(enCuentra) = True Then Continue For

            z = LetraNumero.Tables(0).Rows.IndexOf(enCuentra)

            DataGridView1.Rows(i).Cells(0).Value = CStr(LetraNumero.Tables(0).Rows(z).Item(1))

        Next

    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click

        If IsNothing(DataGridView1.CurrentCell) = True Then Exit Sub

        If DataGridView1.Rows.Count <= 1 Then

            MsgBox("There should be at least 1 field per catalog!!", vbInformation, "DQCT")

            Exit Sub

        End If

        DataGridView1.Rows.RemoveAt(DataGridView1.CurrentCell.RowIndex)

    End Sub

    Private Async Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click

        If DataGridView1.Rows.Count = 0 Then
            MsgBox("There should be at least one column on the catalog!!", vbCritical, "DQCT")
            Exit Sub
        End If

        'Columna 1: FieldCode

        'Columna 2: FieldName

        'Columna 3: Description

        'Columna 4: Position

        'Columna 5: isText
        niuTabcols.Rows.Clear()

        Dim descri As String = ""
        For i = 0 To DataGridView1.Rows.Count - 1

            If IsNothing(DataGridView1.Rows(i).Cells(0).Value) = True Then
                MsgBox("The row " & CStr(i + 1) & " must have the Field Code defined!!, please review!", vbCritical, "DQCT")
                Exit Sub
            End If

            If DataGridView1.Rows(i).Cells(0).Value = "" Then
                MsgBox("The row " & CStr(i + 1) & " must have the Field Code defined!!, please review!", vbCritical, "DQCT")
                Exit Sub
            End If

            If IsNothing(DataGridView1.Rows(i).Cells(1).Value) = True Then
                MsgBox("The row " & CStr(i + 1) & " is missing the Field Name!, please define a name for this row", vbCritical, "DQCT")
                Exit Sub
            End If

            If DataGridView1.Rows(i).Cells(1).Value = "" Then
                MsgBox("The row " & CStr(i + 1) & " is missing the Field Name!, please define a name for this row", vbCritical, "DQCT")
                Exit Sub
            End If

            If IsNothing(DataGridView1.Rows(i).Cells(2).Value) = True Then
                descri = ""
            Else
                If DataGridView1.Rows(i).Cells(2).Value = "" Then
                    descri = ""
                Else
                    descri = CStr(DataGridView1.Rows(i).Cells(2).Value)
                End If
            End If

            If IsNothing(DataGridView1.Rows(i).Cells(3).Value) = True Then
                MsgBox("The row " & CStr(i + 1) & " is missing the Field Type!, please define a field type for this row", vbCritical, "DQCT")
                Exit Sub
            End If

            If DataGridView1.Rows(i).Cells(3).Value = "" Then
                MsgBox("The row " & CStr(i + 1) & " is missing the Field Type!, please define a field type for this row", vbCritical, "DQCT")
                Exit Sub
            End If

            niuTabcols.Rows.Add({DataGridView1.Rows(i).Cells(0).Value, DataGridView1.Rows(i).Cells(1).Value, DataGridView1.Rows(i).Cells(2).Value, CInt(i + 1), DataGridView1.Rows(i).Cells(3).Value})

        Next

        huboCambio = False
        'comparamos vs lo que teniamos al cargarts,
        'Si hay Más Ó menos campos, se debe avisar del cambio!
        If TabColumnas.Rows.Count = 0 Then
            'es la primera vez!!
            'no hay problema, pasa sin ver!
            huboCambio = True

        Else
            'NO es la primera vez!, se evalua si hay una cantidad de campos diferente!
            'todo vs todo, solo verificar si los field codes son los mismos y los tipos siguen siendo los mismos, de lo contrario
            'se avisa!
            Dim cambioMenor As Integer = 0
            Dim numCambios As Integer = 0
            Dim fieldExis As Boolean = False
            Dim sameTipe As Boolean = False
            Dim posSame As Boolean = False
            Dim NameSame As Boolean = False
            Dim descSame As Boolean = False

            For i = 0 To TabColumnas.Rows.Count - 1

                fieldExis = False
                sameTipe = False
                posSame = False
                NameSame = False
                descSame = False

                For j = 0 To niuTabcols.Rows.Count - 1

                    If TabColumnas.Rows(i).Item(0) = niuTabcols.Rows(j).Item(0) Then
                        'SI existe!
                        fieldExis = True

                        If TabColumnas.Rows(i).Item(4) = niuTabcols.Rows(j).Item(4) Then
                            'es el mismo tipo!,
                            'NO hay pex!
                            sameTipe = True
                        Else
                            'cambió  de tipo!!
                        End If

                        If TabColumnas.Rows(i).Item(1) = niuTabcols.Rows(j).Item(1) Then
                            'es el mismo tipo!,
                            'NO hay pex!
                            NameSame = True

                        End If

                        If TabColumnas.Rows(i).Item(2) = niuTabcols.Rows(j).Item(2) Then
                            'es el mismo tipo!,
                            'NO hay pex!
                            descSame = True

                        End If


                        If CInt(TabColumnas.Rows(i).Item(3)) = CInt(niuTabcols.Rows(j).Item(3)) Then
                            posSame = True
                        Else
                            posSame = False
                        End If

                        Exit For

                    End If

                Next

                If fieldExis = False Or sameTipe = False Then
                    numCambios = numCambios + 1
                End If

                If posSame = False Or NameSame = False Or descSame = False Then
                    cambioMenor = cambioMenor + 1
                End If

            Next



            For i = 0 To niuTabcols.Rows.Count - 1

                fieldExis = False
                sameTipe = False
                posSame = False
                NameSame = False
                descSame = False

                For j = 0 To TabColumnas.Rows.Count - 1

                    If niuTabcols.Rows(i).Item(0) = TabColumnas.Rows(j).Item(0) Then
                        'SI existe!
                        fieldExis = True

                        If niuTabcols.Rows(i).Item(4) = TabColumnas.Rows(j).Item(4) Then
                            'es el mismo tipo!,
                            'NO hay pex!
                            sameTipe = True
                        Else
                            'cambió  de tipo!!
                        End If


                        If niuTabcols.Rows(i).Item(1) = TabColumnas.Rows(j).Item(1) Then
                            'es el mismo tipo!,
                            'NO hay pex!
                            NameSame = True

                        End If

                        If niuTabcols.Rows(i).Item(2) = TabColumnas.Rows(j).Item(2) Then
                            'es el mismo tipo!,
                            'NO hay pex!
                            descSame = True

                        End If

                        If CInt(niuTabcols.Rows(i).Item(3)) = CInt(TabColumnas.Rows(j).Item(3)) Then
                            posSame = True
                        Else
                            posSame = False
                        End If


                        Exit For

                    End If

                Next

                If fieldExis = False Or sameTipe = False Then
                    numCambios = numCambios + 1
                End If

                If posSame = False Or NameSame = False Or descSame = False Then
                    cambioMenor = cambioMenor + 1
                End If

            Next

            If numCambios > 0 Then
                'se avisa!
                Dim X As Integer = 0
                X = MsgBox("Multiple changes detected on catalog!, If you change the structure of this catalog you must re-define the records and re-bind affected templates!!. Are you sure you want to continue??", vbQuestion + vbYesNo, "DQCT")
                If X <> 6 Then Exit Sub
                huboCambio = True
            End If

            If cambioMenor > 0 Then
                huboCambio = True
            End If

        End If


        If huboCambio = True Then
            'hacemos aqui el cambio en firebase!
            'niuTabcols
            'agregamos la última columna de la descripción combinada!

            niuTabcols.Rows.Add({"ZZ", "Combined description", "This is the final description of all columns combination", niuTabcols.Rows.Count + 1, "Text"})

            Dim elCamino As String
            elCamino = RaizFire
            elCamino = elCamino & "/" & "catpro"
            elCamino = elCamino & "/" & moduObjeto
            elCamino = elCamino & "/" & tablaCodigo
            Await HazDeleteEnFbSimple(elCamino, "data")
            Await HazPostEnFireBaseConPathYColumnas(elCamino, niuTabcols, "data", -1)

        End If

        'siempre va a hacer reload?
        'Al salir, primero debe re-escribir la nueva estructura en firebase
        'Luego hacer reload de la info con el nuevo data table, 
        'mostrar la estructura en el form1, 
        'jalar los registros sobrevivientes y mostralos, al guardar los cambios, se hace re-write de los registros con los campos
        'bindeados sobrevivientes!

        Me.Close()

    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        'move field up
        If IsNothing(DataGridView1.CurrentCell) = True Then
            MsgBox("Please select a row with data to move up or down!", vbCritical, "DQCT")
            Exit Sub
        End If

        Dim posXold As Integer = DataGridView1.CurrentCell.RowIndex

        If posXold - 1 < 0 Then
            MsgBox("This field is already on top!!", vbExclamation, "MD")
            Exit Sub
        End If

        Dim xRow As New DataGridViewRow
        xRow = DataGridView1.Rows(posXold).Clone()

        For i = 0 To xRow.Cells.Count - 1
            xRow.Cells(i).Value = DataGridView1.Rows(posXold).Cells(i).Value
        Next

        DataGridView1.Rows.InsertRange(posXold - 1, xRow)

        DataGridView1.Rows.RemoveAt(posXold + 1)
    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        'move field down
        If IsNothing(DataGridView1.CurrentCell) = True Then
            MsgBox("Please select a row with data to move up or down!", vbCritical, "DQCT")
            Exit Sub
        End If

        Dim posXold As Integer = DataGridView1.CurrentCell.RowIndex
        If posXold + 1 > DataGridView1.Rows.Count - 1 Then
            MsgBox("This field is already on bottom!!", vbExclamation, "DQCT")
            Exit Sub
        End If

        Dim xRow As New DataGridViewRow
        xRow = DataGridView1.Rows(posXold).Clone()

        For i = 0 To xRow.Cells.Count - 1
            xRow.Cells(i).Value = DataGridView1.Rows(posXold).Cells(i).Value
        Next

        DataGridView1.Rows.InsertRange(posXold + 2, xRow)

        DataGridView1.Rows.RemoveAt(posXold)

    End Sub

    Private Sub Form10_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        CargaRens()
    End Sub
End Class