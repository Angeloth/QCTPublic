Public Class Form2
    Public huboExito As Boolean
    Public elTitulo As String
    Public keyName As String
    Public tabName As String
    Public elCamino As String

    Public keyValue As String
    Public tabValue As String
    Public queOpcion As Integer

    Public pathValue As String
    Public pathLabel As String
    Private MySource As New AutoCompleteStringCollection()
    Public xtraDs As New DataSet
    Public yTraDs As New DataSet
    Public ModuDs As New DataSet
    Private filtroDs As New DataSet

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        huboExito = False

        Dim xObj As Object = Nothing
        Dim yObj As Object = Nothing
        Dim siTuve As Boolean = False
        Dim Posi As Integer = 0
        Dim posDep As Integer = 0

        Select Case queOpcion
            Case Is = 1, 4
                Label1.Text = elTitulo
                TextBox1.Text = elCamino
                Label4.Text = keyName
                Label3.Text = tabName
                Label2.Text = pathLabel
                TextBox2.Text = keyValue
                TextBox3.Text = tabValue
                TextBox1.Enabled = False
                TextBox2.Enabled = False



            Case Is = 2
                'Add template de dependencias!
                Label1.Text = "Add template dependencie"
                Label2.Text = "Template code:"
                Label4.Text = "Template name:"
                Label3.Text = "Module:"

                Dim iAveprim(1) As DataColumn
                Dim kEys As New DataColumn()
                kEys.ColumnName = "ID"
                iAveprim(0) = kEys

                filtroDs.Tables.Clear()
                filtroDs.Tables.Add("filtro")
                filtroDs.Tables(0).PrimaryKey = Nothing
                filtroDs.Tables(0).Columns.Add(kEys) '0
                filtroDs.Tables(0).Columns.Add("Nombre") '1
                filtroDs.Tables(0).Columns.Add("Modulo") '2

                filtroDs.Tables(0).PrimaryKey = iAveprim

                'se agregan SOLO los que no se han agregado anteriormente!
                MySource.Clear()

                For i = 0 To xtraDs.Tables.Count - 1

                    siTuve = False

                    xObj = Nothing
                    xObj = Split(xtraDs.Tables(i).TableName, "#")

                    For j = 0 To yTraDs.Tables.Count - 1
                        yObj = Nothing
                        yObj = Split(yTraDs.Tables(j).TableName, "#")
                        If CStr(xObj(0)) = CStr(yObj(0)) Then
                            siTuve = True
                            Exit For
                        End If
                    Next
                    If siTuve = True Then Continue For
                    MySource.Add(CStr(xObj(0)))
                    filtroDs.Tables(0).Rows.Add({CStr(xObj(0)), CStr(xObj(1)), CStr(xObj(2))})
                Next

                TextBox1.AutoCompleteCustomSource = MySource
                TextBox1.AutoCompleteMode = AutoCompleteMode.Suggest
                TextBox1.AutoCompleteSource = AutoCompleteSource.CustomSource

                TextBox1.Text = ""
                TextBox1.Enabled = True
                TextBox2.Enabled = False
                TextBox3.Enabled = False

            Case Is = 3
                'dependencias agregar una tabla de un template!
                Label1.Text = "Add dependant table"
                Label2.Text = "Template code:"
                Label4.Text = "Table code:"
                Label3.Text = "Table name:"

                Dim iAveprim(1) As DataColumn
                Dim kEys As New DataColumn()
                kEys.ColumnName = "TableCode"
                iAveprim(0) = kEys

                filtroDs.Tables.Clear()
                filtroDs.Tables.Add("filtro")
                filtroDs.Tables(0).PrimaryKey = Nothing
                filtroDs.Tables(0).Columns.Add(kEys) '0
                filtroDs.Tables(0).Columns.Add("Nombre") '1

                filtroDs.Tables(0).PrimaryKey = iAveprim

                MySource.Clear()

                Dim posX As Integer = 0

                Posi = -1
                siTuve = False

                For i = 0 To xtraDs.Tables.Count - 1
                    xObj = Nothing
                    xObj = Split(xtraDs.Tables(i).TableName, "#")
                    If xObj(0) = pathValue Then
                        Posi = i
                        Exit For
                    End If
                Next

                posDep = -1

                For i = 0 To yTraDs.Tables.Count - 1
                    yObj = Nothing
                    yObj = Split(yTraDs.Tables(i).TableName, "#")
                    If pathValue = CStr(yObj(0)) Then
                        posDep = i
                        Exit For
                    End If
                Next

                If Posi >= 0 And posDep >= 0 Then

                    For i = 0 To xtraDs.Tables(Posi).Rows.Count - 1
                        siTuve = False
                        For w = 0 To yTraDs.Tables(posDep).Rows.Count - 1
                            If yTraDs.Tables(posDep).Rows(w).Item(0) = xtraDs.Tables(Posi).Rows(i).Item(1) Then
                                siTuve = True
                                'ya estaba agregada!
                                Exit For
                            End If
                        Next

                        If siTuve = True Then Continue For

                        posX = MySource.IndexOf(xtraDs.Tables(Posi).Rows(i).Item(1))
                        If posX >= 0 Then Continue For

                        MySource.Add(xtraDs.Tables(Posi).Rows(i).Item(1))
                        filtroDs.Tables(0).Rows.Add({CStr(xtraDs.Tables(Posi).Rows(i).Item(1)), CStr(xtraDs.Tables(Posi).Rows(i).Item(2))})

                    Next

                Else
                    MsgBox("No table founds!", vbCritical, TitBox)
                    Me.Close()
                End If

                TextBox1.Text = pathValue
                TextBox1.Enabled = False

                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox3.Enabled = False
                TextBox2.Enabled = True
                'If siTuve = True Then Continue For

                TextBox2.AutoCompleteCustomSource = MySource
                TextBox2.AutoCompleteMode = AutoCompleteMode.Suggest
                TextBox2.AutoCompleteSource = AutoCompleteSource.CustomSource


            Case Is = 5 'caso de add template para templates

                Label2.Text = pathLabel

                Label1.Text = elTitulo
                TextBox1.Text = elCamino
                Label4.Text = keyName
                Label3.Text = tabName

                TextBox2.Text = keyValue
                TextBox3.Text = tabValue
                TextBox1.Enabled = True
                TextBox2.Enabled = True

                MySource.Clear()
                For i = 0 To ModuDs.Tables(0).Rows.Count - 1
                    MySource.Add(CStr(ModuDs.Tables(0).Rows(i).Item(0)).ToUpper())
                Next

                TextBox1.AutoCompleteCustomSource = MySource
                TextBox1.AutoCompleteMode = AutoCompleteMode.Suggest
                TextBox1.AutoCompleteSource = AutoCompleteSource.CustomSource

                TextBox2.AutoCompleteSource = AutoCompleteSource.None
                TextBox2.AutoCompleteMode = AutoCompleteMode.None

                TextBox3.AutoCompleteSource = AutoCompleteSource.None
                TextBox3.AutoCompleteMode = AutoCompleteMode.None


            Case Is = 6
                Label1.Text = elTitulo
                TextBox1.Text = elCamino
                Label4.Text = keyName
                Label3.Text = tabName
                Label2.Text = pathLabel
                TextBox2.Text = keyValue
                TextBox3.Text = tabValue
                TextBox1.Enabled = False
                TextBox2.Enabled = False


        End Select

        Me.CenterToScreen()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Select Case queOpcion

            Case Is = 1, 4
                If TextBox3.Text = "" Then
                    MsgBox("Please provide a name for this new node!", vbCritical, TitBox)
                    Exit Sub
                End If

                If TextBox3.Text.Length <= 2 Then
                    MsgBox("Please provide name longer than 2 characters!!", vbCritical, TitBox)
                    Exit Sub
                End If

                'Aqui falta ver que NO se repita!



            Case Is = 5

                If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
                    MsgBox("All fields are required!!", vbCritical, TitBox)
                    Exit Sub
                End If

                Dim enCuentra As DataRow
                enCuentra = ModuDs.Tables(0).Rows.Find(TextBox1.Text.ToLower())
                If IsNothing(enCuentra) = True Then
                    MsgBox("This is not a valid SAP module!!, please review!!", vbCritical, TitBox)
                    Exit Sub
                End If

                enCuentra = ModuPermit.Tables(0).Rows.Find(TextBox1.Text.ToUpper())
                If IsNothing(enCuentra) = True Then
                    MsgBox("Sorry you are not allowed to add templates on the selected module", vbCritical, TitBox)
                    Exit Sub
                End If

            Case Is = 2
                If TextBox1.Text = "" Then
                    MsgBox("Please type a valid template code!", vbCritical, TitBox)
                    Exit Sub
                End If

                Dim enCuentra As DataRow

                enCuentra = filtroDs.Tables(0).Rows.Find(TextBox1.Text)
                If IsNothing(enCuentra) = True Then
                    MsgBox("Please enter a valid template code!", vbCritical, TitBox)
                    Exit Sub
                End If

                enCuentra = ModuPermit.Tables(0).Rows.Find(TextBox3.Text.ToUpper())
                If IsNothing(enCuentra) = True Then
                    MsgBox("Sorry you are not allowed to add templates on the selected module", vbCritical, TitBox)
                    Exit Sub
                End If


            Case Is = 3

                If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
                    MsgBox("All fields are required!!", vbCritical, TitBox)
                    Exit Sub
                End If


            Case Is = 6
                If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
                    MsgBox("All fields are required!!", vbCritical, TitBox)
                    Exit Sub
                End If


        End Select


        pathValue = TextBox1.Text

        keyValue = TextBox2.Text

        tabValue = TextBox3.Text


        huboExito = True

        Me.Close()

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown

        Dim enCuentra As DataRow

        Select Case queOpcion
            Case Is = 2

                If e.KeyCode = Keys.Return Then

                    enCuentra = filtroDs.Tables(0).Rows.Find(TextBox1.Text)
                    If IsNothing(enCuentra) = True Then
                        MsgBox("Please enter a valid template code!", vbCritical, TitBox)
                        Exit Sub
                    End If

                    Dim z As Integer = 0
                    z = filtroDs.Tables(0).Rows.IndexOf(enCuentra)

                    TextBox2.Text = CStr(filtroDs.Tables(0).Rows(z).Item(1))
                    TextBox3.Text = CStr(filtroDs.Tables(0).Rows(z).Item(2))
                End If


        End Select

    End Sub


    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        Dim enCuentra As DataRow

        Select Case queOpcion
            Case Is = 3
                If e.KeyCode = Keys.Return Then

                    enCuentra = filtroDs.Tables(0).Rows.Find(TextBox2.Text)
                    If IsNothing(enCuentra) = True Then
                        MsgBox("Please enter a valid table code!", vbCritical, TitBox)
                        Exit Sub
                    End If

                    Dim z As Integer = 0
                    z = filtroDs.Tables(0).Rows.IndexOf(enCuentra)
                    TextBox3.Text = CStr(filtroDs.Tables(0).Rows(z).Item(1))
                End If


        End Select

    End Sub
End Class