Imports System.Net.Mail
Public Class Form4

    Private setModul As New DataSet
    Private setMix As New DataSet
    Private setUsers As New DataSet
    Private kienTaba As String
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Apellido,Email,FirstTime,Modulo,Nombre,Pass,Role
        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("Viewer")
        ComboBox1.Items.Add("Editor")
        ComboBox1.Items.Add("Administrator")

        ListView1.Columns.Clear()
        ListView1.Items.Clear()
        ListView1.CheckBoxes = True
        ListView1.View = View.Details
        ListView1.SmallImageList = ImageList1

        ListView1.Columns.Add("Email", 100)
        ListView1.Columns.Add("Name", 50)
        ListView1.Columns.Add("LastName", 100)

        ListView2.Columns.Clear()
        ListView2.Items.Clear()
        ListView2.CheckBoxes = True
        ListView2.View = View.Details
        ListView2.SmallImageList = ImageList1

        ListView2.Columns.Add("Module Code", 70)
        ListView2.Columns.Add("Module Name", 120)

        setMix.Tables.Clear()
        setMix.Tables.Add()
        setMix.Tables(0).Columns.Add()

        kienTaba = ""

        Call ReloadUserList()
        Call ReloadModules()

        ComboBox1.SelectedIndex = 0

        Me.CenterToScreen()

    End Sub

    Private Async Sub ReloadUserList()

        setUsers.Clear()

        setMix.Tables(0).Rows.Clear()
        setMix.Tables(0).Rows.Add({"users"}) '0
        setMix.Tables(0).Rows.Add({"list"}) '1

        setUsers = Await PullUrlWs(setMix, "users")

        ListView1.Items.Clear()

        For i = 0 To setUsers.Tables(0).Rows.Count - 1
            ListView1.Items.Add(setUsers.Tables(0).Rows(i).Item(0), setUsers.Tables(0).Rows(i).Item(0), 0)
            ListView1.Items(ListView1.Items.Count - 1).SubItems.Add(setUsers.Tables(0).Rows(i).Item(1))
            ListView1.Items(ListView1.Items.Count - 1).SubItems.Add(setUsers.Tables(0).Rows(i).Item(2))
        Next

        For i = 0 To ListView2.Items.Count - 1
            ListView2.Items(i).Checked = False
        Next

    End Sub

    Private Async Sub ReloadModules()

        setModul.Clear()
        setMix.Tables(0).Rows.Clear()
        setMix.Tables(0).Rows.Add({"modules"}) '0
        setModul = Await PullUrlWs(setMix, "modules")

        ListView2.Items.Clear()

        For i = 0 To setModul.Tables(0).Rows.Count - 1
            ListView2.Items.Add(setModul.Tables(0).Rows(i).Item(0), CStr(setModul.Tables(0).Rows(i).Item(0)).ToUpper(), 1)
            ListView2.Items(ListView2.Items.Count - 1).SubItems.Add(setModul.Tables(0).Rows(i).Item(1))
        Next

    End Sub

    Private Sub ListView1_ItemChecked(sender As Object, e As ItemCheckedEventArgs) Handles ListView1.ItemChecked

        If e.Item.Checked = True Then
            For i = 0 To ListView1.Items.Count - 1
                If i <> e.Item.Index Then
                    ListView1.Items(i).Checked = False
                End If
            Next
            Dim z As Integer = 0
            Dim enCuentra As DataRow
            Dim xObj As Object
            kienTaba = e.Item.Text
            If kienTaba = "" Then Exit Sub

            enCuentra = setUsers.Tables(0).Rows.Find(kienTaba)

            If IsNothing(enCuentra) = True Then
                MsgBox("User not found!", vbCritical, "SAP MD")
            Else

                z = setUsers.Tables(0).Rows.IndexOf(enCuentra)

                TextBox1.Text = CStr(setUsers.Tables(0).Rows(z).Item(0))
                TextBox2.Text = CStr(setUsers.Tables(0).Rows(z).Item(1))
                TextBox3.Text = CStr(setUsers.Tables(0).Rows(z).Item(2))

                If setUsers.Tables(0).Rows(z).Item(3) = "X" Then
                    CheckBox1.Checked = True
                Else
                    CheckBox1.Checked = False
                End If

                Select Case CStr(setUsers.Tables(0).Rows(z).Item(5))
                    Case Is = "Viewer"
                        ComboBox1.SelectedIndex = 0

                    Case Is = "Editor"
                        ComboBox1.SelectedIndex = 1

                    Case Is = "Admin"
                        ComboBox1.SelectedIndex = 2

                End Select

                xObj = Split(setUsers.Tables(0).Rows(z).Item(6), "#")

                For j = 0 To ListView2.Items.Count - 1
                    ListView2.Items(j).Checked = False
                    For i = 0 To UBound(xObj)
                        If ListView2.Items(j).Text = xObj(i) Then
                            ListView2.Items(j).Checked = True
                        End If
                    Next
                Next


            End If

        Else
            kienTaba = ""
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            CheckBox1.Checked = False
        End If

    End Sub

    Private Sub ToolStripButton1_CheckedChanged(sender As Object, e As EventArgs) Handles ToolStripButton1.CheckedChanged

        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        CheckBox1.Checked = False

        For i = 0 To ListView1.Items.Count - 1
            ListView1.Items(i).Checked = False
        Next

        If ToolStripButton1.Checked = True Then
            ToolStripLabel1.Text = "Add user"
            ListView1.Enabled = False
            TextBox1.Enabled = True
            CheckBox1.Checked = True
            CheckBox1.Enabled = False
            ToolStripLabel1.ForeColor = Color.DarkGreen
        Else
            ToolStripLabel1.Text = "Edit user"
            ListView1.Enabled = True
            TextBox1.Enabled = False
            CheckBox1.Checked = False
            CheckBox1.Enabled = True
            ToolStripLabel1.ForeColor = Color.FromArgb(220, 20, 60)
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        Select Case ComboBox1.SelectedIndex
            Case Is = 0
                ListView2.Enabled = False

            Case Is = 1
                ListView2.Enabled = True

            Case Is = 2
                ListView2.Enabled = False

        End Select

    End Sub

    Private Async Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        'save1

        If TextBox1.Text = "" Then
            MsgBox("Please enter a valid Email address!", vbCritical, "SAP MD")
            Exit Sub
        End If

        Err.Clear()
        Try
            Dim xMail As New MailAddress(TextBox1.Text)
        Catch ex As Exception
            MsgBox("Please enter a valid Email address!", vbCritical, "SAP MD")
            Exit Sub
        End Try

        If TextBox2.Text = "" Then
            MsgBox("Please enter a Name for this user", vbCritical, "SAP MD")
            Exit Sub
        End If

        If TextBox3.Text = "" Then
            MsgBox("Please enter a Last name for this user", vbCritical, "SAP MD")
            Exit Sub
        End If

        If ComboBox1.SelectedIndex < 0 Then
            MsgBox("Please select a role for this user!!", vbCritical, "SAP MD")
            Exit Sub
        End If

        Dim cadModulos As String = "NA"
        If ComboBox1.SelectedIndex = 1 Then
            cadModulos = ""
            Dim k As Integer = 0
            For i = 0 To ListView2.Items.Count - 1
                If ListView2.Items(i).Checked = True Then
                    If k <> 0 Then cadModulos = cadModulos & "#"
                    cadModulos = cadModulos & ListView2.Items(i).Text
                    k = k + 1
                End If
            Next

            If cadModulos = "" Then
                MsgBox("Please check some modules for this Editor role!", vbCritical, "SAP MD")
                Exit Sub
            End If

        End If

        Dim suPass As String = ""
        Dim z As Integer = 0
        Dim cadChek As String = " "
        Dim elRol As String = ""

        If CheckBox1.Checked = True Then
            cadChek = "X"
            suPass = getSHA1Hash("NextCloud2021")
        End If

        Select Case ComboBox1.SelectedIndex
            Case Is = 0
                elRol = "Viewer"

            Case Is = 1
                elRol = "Editor"

            Case Is = 2
                elRol = "Admin"

        End Select

        Dim enCuentra As DataRow
        Dim xSet As New DataTable
        Dim elRet As String = ""

        xSet = setUsers.Tables(0).Clone()

        Dim elCamino As String = RaizFire
        elCamino = elCamino & "/users" ' & elNode.Parent.Parent.Name 'compania

        If ToolStripButton1.Checked = True Then
            'Add user
            enCuentra = setUsers.Tables(0).Rows.Find(TextBox1.Text)
            If IsNothing(enCuentra) = False Then
                'ya estaba!!
                MsgBox("This user already exists!!, if you want to edit it, please click on Edit User and check it from the list on the left!", vbCritical, "SAP MD")
                Exit Sub
            End If

            xSet.Rows.Add({TextBox1.Text, TextBox2.Text, TextBox3.Text, cadChek, suPass, elRol, cadModulos, ""})

            elRet = Await HazPostEnFireBaseConPathYColumnas(elCamino, xSet, "list", 7)
            MsgBox(elRet, vbInformation, "SAP MD")

            Call ReloadUserList()

        Else
            'update user!
            'un poco diferente!, se debe actualizar!
            enCuentra = setUsers.Tables(0).Rows.Find(TextBox1.Text)
            If IsNothing(enCuentra) = True Then
                MsgBox("User not found!!", vbCritical, "SAP MD")
                Exit Sub
            End If

            elCamino = elCamino & "/list"

            z = setUsers.Tables(0).Rows.IndexOf(enCuentra)

            If CheckBox1.Checked = True Then
                suPass = getSHA1Hash("NextCloud2021") 'password reset!
            Else
                suPass = setUsers.Tables(0).Rows(z).Item(4) 'mismo password!
            End If

            xSet.Rows.Add({TextBox1.Text, TextBox2.Text, TextBox3.Text, cadChek, suPass, elRol, cadModulos, CStr(setUsers.Tables(0).Rows(z).Item(7))})

            elRet = Await HazPutEnFireBasePathYColumnas(elCamino, xSet, 7)

            MsgBox(elRet, vbInformation, "SAP MD")
            Call ReloadUserList()

        End If

    End Sub

    Private Async Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        'eliminar usuario!
        Dim elUsr As String = ""

        For i = 0 To ListView1.Items.Count - 1
            If ListView1.Items(i).Checked = True Then
                elUsr = ListView1.Items(i).Name 'correo!
                Exit For
            End If
        Next

        If elUsr = "" Then
            MsgBox("Please check a user to delete first!!", vbCritical, "SAP MP")
            Exit Sub
        End If

        Dim enCuentra As DataRow

        enCuentra = setUsers.Tables(0).Rows.Find(elUsr)

        If IsNothing(enCuentra) = True Then
            'no se encuentra el usuario!
            MsgBox("User not found!!", vbCritical, "SAP MD")
            Exit Sub
        End If

        Dim z As Integer = -1
        z = setUsers.Tables(0).Rows.IndexOf(enCuentra)

        If setUsers.Tables(0).Rows(z).Item(5) = "Admin" Then
            'quiere borrar un admin!!
            'se verifica!
            Dim filterDT As DataTable = setUsers.Tables(0).Clone()

            Dim result() As DataRow = setUsers.Tables(0).Select("Role = 'Admin'")

            For Each row As DataRow In result
                filterDT.ImportRow(row)
            Next

            If filterDT.Rows.Count <= 2 Then
                MsgBox("You can't delete this Admin, there must be at least 2 admins in the user list!!", vbCritical, "SAP MD")
                Exit Sub
            End If

        End If

        'pasa sin ver!
        Dim elRet As String = ""
        Dim pathDel As String = ""
        pathDel = RaizFire
        pathDel = pathDel & "/users/list"

        elRet = Await HazDeleteEnFbSimple(pathDel, CStr(setUsers.Tables(0).Rows(z).Item(7)))

        If elRet = "Ok" Then
            MsgBox("User gone!", vbInformation, "SAP MD")
            Call ReloadUserList()
        Else
            MsgBox(elRet, vbInformation, "SAP MD")
        End If

    End Sub

End Class