Public Class Form3

    Private depFieldOk As Boolean
    Private conFieldOk As Boolean

    Public elEnfoque As String

    Public elPath As String
    Public elDepds As New DataSet
    Public xtraDs As New DataSet
    Public yTraDs As New DataSet
    Public huboExito As Boolean
    Public depeTabla As String
    Public depeTemplate As String
    Public depeCampo As String

    Private filtroDt As DataTable 'tabla dependiente

    Private matchDt As New DataTable

    Private filtroDep As DataTable

    Private PosiX As Integer = 0
    Private PosiY As Integer = 0

    Private MySourceDepField As New AutoCompleteStringCollection()
    Private MySourceTempDs As New AutoCompleteStringCollection()
    Private DatDepField As New DataSet
    Private DatConTempDs As New DataSet
    Private DsBox6 As New DataSet
    Private DsBox8 As New DataSet

    Public resDepFieldCode As String
    Public resDepFieldName As String
    Public resConTempCode As String
    Public resConTempName As String
    Public resConTempModule As String
    Public resConTableCode As String
    Public resConTableName As String
    Public resConFieldCode As String
    Public resConFieldName As String

    Public resConType As String
    Public resConRule As String
    Public resConVal As String
    Public resMachFields As String
    Public resScope As String

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        huboExito = False
        depFieldOk = False
        conFieldOk = False

        Label1.Text = "Add dependant field"
        Label2.Text = elPath

        TextBox1.Enabled = True

        TextBox1.Text = "" 'field code
        TextBox2.Text = "" 'filed name

        TextBox3.Text = "" 'template code
        TextBox4.Text = "" 'template name
        TextBox5.Text = "" 'module
        TextBox6.Text = "" 'table code
        TextBox7.Text = "" 'table name
        TextBox8.Text = "" 'field code
        TextBox9.Text = "" 'field name
        TextBox10.Text = ""
        TextBox11.Text = ""

        Label13.Text = "Not set"
        ComboBox1.Items.Clear()


        ComboBox2.Items.Clear()
        ComboBox2.Items.Add("To Condition") 'para condicionar el valor
        ComboBox2.Items.Add("To Apply") 'para aplicar la evaluación!
        ComboBox2.SelectedIndex = 0
        ComboBox2.Enabled = False

        TextBox2.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False
        TextBox9.Enabled = False
        ComboBox1.Enabled = False
        TextBox10.Enabled = False
        TextBox11.Enabled = False

        TextBox3.Enabled = True

        Button2.Visible = False
        Button3.Visible = False
        Button4.Visible = False
        Button5.Visible = False

        'filtramos
        Dim xObj As Object = Nothing
        PosiX = -1
        For i = 0 To xtraDs.Tables.Count - 1
            xObj = Split(xtraDs.Tables(i).TableName, "#")
            If xObj(0) = depeTemplate Then
                PosiX = i
                Exit For
            End If
            xObj = Nothing
        Next

        PosiY = -1
        For i = 0 To yTraDs.Tables.Count - 1
            xObj = Split(yTraDs.Tables(i).TableName, "#")
            If xObj(0) = depeTemplate Then
                PosiY = i
                Exit For
            End If
            xObj = Nothing
        Next

        If PosiX >= 0 Then 'And PosiY >= 0'esto si estaba!

            Dim iAveprim(1) As DataColumn
            Dim kEys As New DataColumn()
            kEys.ColumnName = "FieldCode"
            iAveprim(0) = kEys

            DatDepField.Tables.Clear()
            DatDepField.Tables.Add("Dependant")
            DatDepField.Tables(0).Columns.Add(kEys)
            DatDepField.Tables(0).Columns.Add("Name")
            DatDepField.Tables(0).PrimaryKey = iAveprim


            Dim Yaveprim(1) As DataColumn
            Dim LLave As New DataColumn()
            LLave.ColumnName = "FieldCode"
            Yaveprim(0) = LLave

            DatConTempDs.Tables.Clear()
            DatConTempDs.Tables.Add("CondiTemp")
            DatConTempDs.Tables(0).Columns.Add(LLave)
            DatConTempDs.Tables(0).Columns.Add("Name")
            DatConTempDs.Tables(0).Columns.Add("Module")
            DatConTempDs.Tables(0).PrimaryKey = Yaveprim

            'filtroDt.Clear()
            filtroDt = xtraDs.Tables(PosiX).Clone
            Dim result() As DataRow = xtraDs.Tables(PosiX).Select("TableCode = '" & depeTabla & "'")

            For Each row As DataRow In result
                filtroDt.ImportRow(row)
            Next

            'OJO, esto si estaba, depende de posiY
            If PosiY >= 0 Then
                filtroDep = yTraDs.Tables(PosiY).Clone
                Dim risult() As DataRow = yTraDs.Tables(PosiY).Select("DepTableCode = '" & depeTabla & "'")
                For Each raw As DataRow In risult
                    filtroDep.ImportRow(raw)
                Next
            Else
                filtroDep = yTraDs.Tables(0).Clone
            End If

            Dim yaTaba As Boolean = False

            MySourceDepField.Clear()
            For i = 0 To filtroDt.Rows.Count - 1
                'yaTaba = False
                'For j = 0 To filtroDep.Rows.Count - 1
                '    If filtroDt.Rows(i).Item(3) = filtroDep.Rows(j).Item(2) Then
                '        yaTaba = True
                '        Exit For
                '    End If
                'Next

                'If yaTaba = True Then Continue For
                MySourceDepField.Add(CStr(filtroDt.Rows(i).Item(3)))

                DatDepField.Tables(0).Rows.Add({filtroDt.Rows(i).Item(3), filtroDt.Rows(i).Item(4)})
            Next

            'se agregan todos los templates MENOS a si mismo!!
            MySourceTempDs.Clear()
            For i = 0 To xtraDs.Tables.Count - 1
                xObj = Nothing
                xObj = Split(xtraDs.Tables(i).TableName, "#")
                'If CStr(xObj(0)) = depeTemplate Then Continue For'esta linea SI estaba!
                MySourceTempDs.Add(CStr(xObj(0)))
                DatConTempDs.Tables(0).Rows.Add({CStr(xObj(0)), CStr(xObj(1)), CStr(xObj(2))})
            Next

            TextBox1.AutoCompleteCustomSource = MySourceDepField
            TextBox1.AutoCompleteMode = AutoCompleteMode.Suggest
            TextBox1.AutoCompleteSource = AutoCompleteSource.CustomSource

            TextBox3.AutoCompleteCustomSource = MySourceTempDs
            TextBox3.AutoCompleteMode = AutoCompleteMode.Suggest
            TextBox3.AutoCompleteSource = AutoCompleteSource.CustomSource


        Else
            MsgBox("Tables not found!!", vbCritical, TitBox)
            Me.Close()
            Exit Sub
        End If

        Me.CenterToScreen()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If ComboBox2.SelectedIndex = 0 Then

            If ComboBox1.SelectedIndex <> 1 And ComboBox1.SelectedIndex <> 2 Then
                If TextBox10.Text = "" Then
                    MsgBox("Please define a Conditional Value to Match a condition rule!!", vbCritical, TitBox)
                    Exit Sub
                End If
            End If

        End If


        If depFieldOk = True And conFieldOk = True Then
            'tenemos todo para meterlo de una vez a FB!
            resDepFieldCode = TextBox1.Text
            resDepFieldName = TextBox2.Text
            resConTempCode = TextBox3.Text
            resConTempName = TextBox4.Text
            resConTempModule = TextBox5.Text
            resConTableCode = TextBox6.Text
            resConTableName = TextBox7.Text
            resConFieldCode = TextBox8.Text
            resConFieldName = TextBox9.Text

            resConType = Label13.Text
            resConRule = ComboBox1.Items(ComboBox1.SelectedIndex)
            resConVal = TextBox10.Text
            resMachFields = TextBox11.Text

            resScope = ComboBox2.Items(ComboBox2.SelectedIndex)

            huboExito = True

            Me.Close()

        Else
            MsgBox("Please complete all fields!", vbCritical, TitBox)
        End If


    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Return Then
            Dim enCuentra As DataRow
            enCuentra = DatDepField.Tables(0).Rows.Find(TextBox1.Text)
            If IsNothing(enCuentra) = True Then
                MsgBox("Please enter a valid field code!", vbCritical, TitBox)
                Exit Sub
            End If

            Dim z As Integer = 0
            z = DatDepField.Tables(0).Rows.IndexOf(enCuentra)
            TextBox2.Text = CStr(DatDepField.Tables(0).Rows(z).Item(1))

            depFieldOk = True

            If depeCampo <> "" Then
                TextBox1.Enabled = False
            End If

        End If

    End Sub


    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Return Then
            Dim enCuentra As DataRow
            enCuentra = DatConTempDs.Tables(0).Rows.Find(TextBox3.Text)
            If IsNothing(enCuentra) = True Then
                MsgBox("Please enter a valid template code!", vbCritical, TitBox)
                Exit Sub
            End If

            Dim z As Integer = 0
            z = DatConTempDs.Tables(0).Rows.IndexOf(enCuentra)
            TextBox4.Text = CStr(DatConTempDs.Tables(0).Rows(z).Item(1))
            TextBox5.Text = CStr(DatConTempDs.Tables(0).Rows(z).Item(2))
            TextBox3.Enabled = False
            Call LlenaBox6(TextBox3.Text)

        End If

    End Sub

    Private Sub LlenaBox6(ByVal queTemplate As String)

        Dim MiLocalSource As New AutoCompleteStringCollection()

        Dim iAveprim(1) As DataColumn
        Dim kEys As New DataColumn()
        kEys.ColumnName = "TableCode"
        iAveprim(0) = kEys

        DsBox6.Tables.Clear()
        DsBox6.Tables.Add()
        DsBox6.Tables(0).Columns.Add(kEys)
        DsBox6.Tables(0).Columns.Add("Name")
        DsBox6.Tables(0).PrimaryKey = iAveprim

        Dim xObj As Object = Nothing
        Dim Pos As Integer = -1
        For i = 0 To xtraDs.Tables.Count - 1
            xObj = Split(xtraDs.Tables(i).TableName, "#")
            If xObj(0) = queTemplate Then
                Pos = i
                Exit For
            End If
        Next

        If Pos >= 0 Then
            Dim laYe As Integer = -1
            MiLocalSource.Clear()

            For i = 0 To xtraDs.Tables(Pos).Rows.Count - 1
                laYe = MiLocalSource.IndexOf(xtraDs.Tables(Pos).Rows(i).Item(1))
                If laYe >= 0 Then Continue For
                MiLocalSource.Add(xtraDs.Tables(Pos).Rows(i).Item(1))
                DsBox6.Tables(0).Rows.Add({xtraDs.Tables(Pos).Rows(i).Item(1), xtraDs.Tables(Pos).Rows(i).Item(2)})
            Next

            TextBox6.AutoCompleteCustomSource = MiLocalSource
            TextBox6.AutoCompleteMode = AutoCompleteMode.Suggest
            TextBox6.AutoCompleteSource = AutoCompleteSource.CustomSource

            TextBox6.Text = ""
            TextBox6.Enabled = True

            Button3.Visible = True
            Button4.Visible = False
        Else
            MsgBox("", vbCritical, TitBox)
        End If

    End Sub

    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown

        If e.KeyCode = Keys.Return Then
            Dim enCuentra As DataRow
            enCuentra = DsBox6.Tables(0).Rows.Find(TextBox6.Text)
            If IsNothing(enCuentra) = True Then
                MsgBox("Please enter a valid template code!", vbCritical, TitBox)
                Exit Sub
            End If

            Dim z As Integer = 0
            z = DsBox6.Tables(0).Rows.IndexOf(enCuentra)
            TextBox7.Text = CStr(DsBox6.Tables(0).Rows(z).Item(1))
            TextBox6.Enabled = False
            Call LlenaBox8(TextBox3.Text, TextBox6.Text)
        End If

    End Sub

    Private Sub LlenaBox8(ByVal queTemplate As String, ByVal queTabla As String)

        Dim MiLocalSource As New AutoCompleteStringCollection()

        Dim iAveprim(1) As DataColumn
        Dim kEys As New DataColumn()
        kEys.ColumnName = "FieldCode"
        iAveprim(0) = kEys

        DsBox8.Tables.Clear()
        DsBox8.Tables.Add()
        DsBox8.Tables(0).Columns.Add(kEys)
        DsBox8.Tables(0).Columns.Add("Name")
        DsBox8.Tables(0).PrimaryKey = iAveprim

        Dim xObj As Object = Nothing
        Dim Pos As Integer = -1
        For i = 0 To xtraDs.Tables.Count - 1
            xObj = Split(xtraDs.Tables(i).TableName, "#")
            If xObj(0) = queTemplate Then
                Pos = i
                Exit For
            End If
        Next

        'depetemplate
        'depeTabla
        'depeCampo
        Dim misMo As Boolean = False

        If depeTemplate = TextBox3.Text And depeTabla = TextBox6.Text Then misMo = True

        If Pos >= 0 Then
            Dim filterDT As DataTable = xtraDs.Tables(Pos).Clone()

            Dim result() As DataRow = xtraDs.Tables(Pos).Select("TableCode = '" & queTabla & "'")

            For Each row As DataRow In result
                filterDT.ImportRow(row)
            Next

            matchDt.Clear()
            matchDt = filterDT.Copy()

            'SOLO si es interno, entonces nunca debe ser el mismo campo!!

            For i = 0 To filterDT.Rows.Count - 1
                If MiLocalSource.IndexOf(filterDT.Rows(i).Item(3)) >= 0 Then Continue For
                If misMo = True Then
                    If filterDT.Rows(i).Item(3) = TextBox1.Text Then Continue For
                End If
                MiLocalSource.Add(filterDT.Rows(i).Item(3))
                DsBox8.Tables(0).Rows.Add({filterDT.Rows(i).Item(3), filterDT.Rows(i).Item(4)})
            Next

            TextBox8.AutoCompleteCustomSource = MiLocalSource
            TextBox8.AutoCompleteMode = AutoCompleteMode.Suggest
            TextBox8.AutoCompleteSource = AutoCompleteSource.CustomSource

            TextBox8.Text = ""
            TextBox8.Enabled = True

            Button3.Visible = False
            Button4.Visible = True

            If misMo = True Then
                'esta en el mismo objeto!, el matchfields no aplica!
                TextBox11.Text = "None"
                Button5.Visible = False
            Else
                'esta en un diferente objeto, el matchfields SI aplica!
                TextBox11.Text = "None"
                Button5.Visible = True
            End If

        Else
            MsgBox("No table found!", vbCritical, TitBox)

        End If

    End Sub

    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown

        If e.KeyCode = Keys.Return Then
            Dim enCuentra As DataRow
            enCuentra = DsBox8.Tables(0).Rows.Find(TextBox8.Text)
            If IsNothing(enCuentra) = True Then
                MsgBox("Please enter a valid template code!", vbCritical, TitBox)
                Exit Sub
            End If

            Dim z As Integer = 0
            z = DsBox8.Tables(0).Rows.IndexOf(enCuentra)
            TextBox9.Text = CStr(DsBox8.Tables(0).Rows(z).Item(1))

            conFieldOk = True

            If elEnfoque = "B" Then 'si es de construccion, entonces SIEMPRE será apply!
                ComboBox2.SelectedIndex = 1 '
                ComboBox2.Enabled = False
            Else
                ComboBox2.Enabled = True
            End If

            Call LlenaCombos()

        End If

    End Sub
    Private Sub LlenaCombos()

        'la unica diferencia es que si viene de Scope de Construction, debe ponerse SIEMPRE regla Apply
        'Y deshabilitar el textbox10!

        ComboBox1.Items.Clear()
        If conFieldOk = False Then
            ComboBox1.Enabled = False
            Exit Sub
        End If

        If depeTemplate = TextBox3.Text Then

            If depeTabla = TextBox6.Text Then
                Label13.Text = "Local" 'misma hoja!

                If ComboBox2.SelectedIndex = 0 Then
                    'condiciona
                    ComboBox1.Items.Add("OR")
                    ComboBox1.Items.Add("NULL")
                    ComboBox1.Items.Add("NOTNULL")
                    ComboBox1.Items.Add("MATCHS")
                    ComboBox1.Items.Add("STARTWITH")
                    ComboBox1.Items.Add("ENDWITH")
                    ComboBox1.Items.Add("CONTAINS")
                    ComboBox1.Items.Add("EXCEPT")
                Else
                    'aplica regla!
                    ComboBox1.Items.Add("MATCHS")
                    ComboBox1.Items.Add("STARTWITH")
                    ComboBox1.Items.Add("ENDWITH")
                    ComboBox1.Items.Add("CONTAINS")

                    'Si el valor en VALUE es VACIO, entonces va a aplicar la regla vs la columna condicional

                End If

                'que termine o inicia con el valor de la columna condicional!
                'que contenga el valor de lo escrito en la columna condicional!
                'ó que inicie o termine con el valor escrito en el textbox!

            Else

                Label13.Text = "Internal" 'mismo objeto, diferente hoja

                If ComboBox2.SelectedIndex = 0 Then
                    ComboBox1.Items.Add("OR")
                    ComboBox1.Items.Add("NULL")
                    ComboBox1.Items.Add("NOTNULL")
                    ComboBox1.Items.Add("MATCHS")
                    ComboBox1.Items.Add("STARTWITH")
                    ComboBox1.Items.Add("ENDWITH")
                    ComboBox1.Items.Add("CONTAINS")
                    ComboBox1.Items.Add("EXCEPT")
                Else
                    ComboBox1.Items.Add("MATCHS")
                    ComboBox1.Items.Add("STARTWITH")
                    ComboBox1.Items.Add("ENDWITH")
                    ComboBox1.Items.Add("CONTAINS")
                End If

            End If

        Else
            Label13.Text = "External"

            If ComboBox2.SelectedIndex = 0 Then
                ComboBox1.Items.Add("OR")
                ComboBox1.Items.Add("NULL")
                ComboBox1.Items.Add("NOTNULL")
                ComboBox1.Items.Add("MATCHS")
                ComboBox1.Items.Add("STARTWITH")
                ComboBox1.Items.Add("ENDWITH")
                ComboBox1.Items.Add("CONTAINS")
                ComboBox1.Items.Add("EXCEPT")
            Else
                ComboBox1.Items.Add("MATCHS")
                ComboBox1.Items.Add("STARTWITH")
                ComboBox1.Items.Add("ENDWITH")
                ComboBox1.Items.Add("CONTAINS")
            End If

        End If

        'dependencias entre hojas
        If elEnfoque = "B" Then
            ComboBox1.SelectedIndex = 0
            ComboBox1.Enabled = False 'La regla SIEMPRE va a ser MATCHS! para construccion!
        Else

            ComboBox1.Enabled = True
            ComboBox1.SelectedIndex = 0

        End If



    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox3.Enabled = True
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox6.Enabled = False
        Button3.Visible = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        TextBox6.Enabled = True
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox8.Enabled = False
        Button4.Visible = False
        Button3.Visible = True

        TextBox11.Text = "None"
        Button5.Visible = False

        ComboBox2.Enabled = False

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        'depFieldOk = False
        depFieldOk = False
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        conFieldOk = False
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        If elEnfoque = "B" Then
            TextBox10.Text = ""
            TextBox10.Enabled = False
        Else
            'Si se aplica la regla entonces debe tener al menos 1 renglon al encontrarlo
            'Si condiciona
            Select Case CStr(ComboBox1.Items(ComboBox1.SelectedIndex))
                Case Is = "NULL", "NOTNULL"
                    TextBox10.Text = ""
                    TextBox10.Enabled = False

                Case Else

                    TextBox10.Text = ""
                    If ComboBox2.SelectedIndex = 1 Then
                        'Es de aplicación, se deshabilita el text10
                        TextBox10.Enabled = False
                    Else
                        TextBox10.Enabled = True
                    End If

                    'TextBox10.Enabled = True

            End Select

        End If

    End Sub

    Private Sub Form3_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If depeCampo <> "" Then
            TextBox1.Focus()
            TextBox1.Text = depeCampo
            SendKeys.Send("{ENTER}")
            '

        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Form9.depeField = depeCampo
        Form9.tabHija = filtroDt.Copy()
        Form9.tabPapa = matchDt.Copy()
        Form9.ShowDialog()

        TextBox11.Text = Form9.respCadena

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        Call LlenaCombos()

    End Sub
End Class