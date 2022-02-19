Public Class LoginForm1
    Private setUsuarios As New DataSet
    Private setMix As New DataSet
    Private xSource As New AutoCompleteStringCollection()

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See https://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click

        'validar el login del usuario!!

        If UsernameTextBox.Text = "" Or PasswordTextBox.Text = "" Then
            MsgBox("Please fill you Email and Password to continue!!", vbCritical, TitBox)
            Exit Sub
        End If

        Dim enCuentra As DataRow
        enCuentra = setUsuarios.Tables(0).Rows.Find(UsernameTextBox.Text)

        If IsNothing(enCuentra) = True Then
            MsgBox("This user don't exists!, please verify!!", vbCritical, TitBox)
            Exit Sub
        End If

        Dim z As Integer = -1
        z = setUsuarios.Tables(0).Rows.IndexOf(enCuentra)

        If getSHA1Hash(PasswordTextBox.Text) <> setUsuarios.Tables(0).Rows(z).Item(4) Then
            MsgBox("Invalid password!!", vbCritical, TitBox)
            Exit Sub
        End If

        'antes deee!!
        'please 

        If setUsuarios.Tables(0).Rows(z).Item(3) = "X" Then
            'First Time!
            Form6.LlaveUsuario = setUsuarios.Tables(0).Rows(z).Item(7)
            Form6.huboExito = False
            Form6.ShowDialog()
        Else
            Form6.huboExito = True
        End If

        If Form6.huboExito = False Then
            MsgBox("Please provide a new password for your session!!", vbCritical, TitBox)
            Exit Sub
        End If

        UsuarioCorreo = setUsuarios.Tables(0).Rows(z).Item(0)
        UsuarioNombre = setUsuarios.Tables(0).Rows(z).Item(1)
        UsuarioApellido = setUsuarios.Tables(0).Rows(z).Item(2)
        UsuarioRole = setUsuarios.Tables(0).Rows(z).Item(5)
        UsuarioModulos = setUsuarios.Tables(0).Rows(z).Item(6) 'modulos

        RoleUsuario = setUsuarios.Tables(0).Rows(z).Item(5)

        Dim iAveprim(1) As DataColumn
        Dim kEys As New DataColumn()
        kEys.ColumnName = "Unike"
        iAveprim(0) = kEys

        ModuPermit.Tables.Clear()
        ModuPermit.Tables.Add()
        ModuPermit.Tables(0).Columns.Add(kEys)
        ModuPermit.Tables(0).Columns.Add("aLower")
        ModuPermit.Tables(0).PrimaryKey = iAveprim

        If RoleUsuario = "Editor" Then

            Dim xObj As Object = Nothing
            xObj = Split(setUsuarios.Tables(0).Rows(z).Item(6), "#")

            For i = 0 To UBound(xObj)
                ModuPermit.Tables(0).Rows.Add({xObj(i), CStr(xObj(i)).ToLower()})
            Next

        End If


        enCuentra = My.Settings.xUsers.Rows.Find(UsernameTextBox.Text)
        If IsNothing(enCuentra) = True Then
            My.Settings.xUsers.Rows.Add({UsernameTextBox.Text})
            My.Settings.Save()
        End If

        Form5.Show()
        Me.Close()

    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Private Async Sub LoginForm1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.CenterToScreen()

        UsuarioCorreo = ""
        UsuarioNombre = ""
        UsuarioApellido = ""
        UsuarioRole = ""
        UsuarioModulos = "" 'modulos

        setUsuarios.Tables.Clear()

        setMix.Tables.Clear()
        setMix.Tables.Add()
        setMix.Tables(0).Columns.Add()

        setMix.Tables(0).Rows.Clear()
        setMix.Tables(0).Rows.Add({"users"}) '0
        setMix.Tables(0).Rows.Add({"list"}) '1

        setUsuarios = Await PullUrlWs(setMix, "users")

        xSource.Clear()

        If IsNothing(My.Settings.xUsers) = True Then
            My.Settings.xUsers = New DataTable
        End If

        If My.Settings.xUsers.Columns.Count = 0 Then
            AsignaYavePrimariaATabla(My.Settings.xUsers, "Usuarios", True)
        End If

        For i = 0 To My.Settings.xUsers.Rows.Count - 1
            xSource.Add(My.Settings.xUsers.Rows(i).Item(0))
        Next

        If xSource.Count = 0 Then Exit Sub

        UsernameTextBox.AutoCompleteCustomSource = xSource
        UsernameTextBox.AutoCompleteMode = AutoCompleteMode.Suggest
        UsernameTextBox.AutoCompleteSource = AutoCompleteSource.CustomSource

    End Sub

    Private Async Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click



    End Sub
End Class
