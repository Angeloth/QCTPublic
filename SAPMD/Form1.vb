Imports System.Data
Public Class Form1
    Public writeDs As New DataSet
    Private usaDataset As New DataSet
    Private catDs As New DataSet
    Private depeDs As New DataSet
    Private modsDs As New DataSet
    Private tempDs As New DataSet
    Private recDs As New DataSet
    Private arbolDs As New DataSet
    Private ModuloDs As New DataSet
    Private CatSimple As New DataSet
    Private ConstruDs As New DataSet

    Private relaUniqRul As New DataTable

    Private internRecords As New DataSet
    Private relaTionDs As New DataSet
    Private internRelation As New DataSet

    Private BuildBobDs As New DataSet
    Private TableBuild As New DataTable

    Private filtRelas As New DataTable
    Private filtCat As New DataTable
    Private filtDepe As New DataTable
    Private MySource As New AutoCompleteStringCollection()
    Private FuenteCatalogos As New AutoCompleteStringCollection()
    Private NodoNameActual As String
    Private CategSelected As Integer
    Private elNode As TreeNode
    Private estoyAgregandoRows As Boolean
    Private objetoSelek As String
    Private tableSelek As String
    Private moduloSelek As String
    Private objetoName As String
    Private tablaNombre As String
    Private compaSelekted As String

    Private ValidaDt As New DataTable

    Private multiDepe As New DataSet
    Private RecDt As New DataTable

    Private xDepeDs As New DataSet 'este se usa momentaneamente para validar!
    Private rowEstaba As Long
    Private puSSyCat As Integer
    Private PussyTemp As Integer
    Private cataNombre As String

    Private toyTrabajando As Boolean

    Private pathFinduse As String
    Private editDs As New DataSet 'set para ver los datos si se puede editar o no!
    Private puedoEditar As Boolean
    Private lastNodeEdit As String
    Private inUseDt As New DataTable


    Private DoneEvaluacion As Boolean

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        DoneEvaluacion = False

        'Para el status de objeto en uso
        'Que se postee el objeto, su nombre, la hoja, su nombre, y el nodo de inuse:X ó vacío
        'Y que se postee la hora, timestamp de utc
        'que pasa si cambio de nodo ó me salgo?
        'debo trazar que nodo estaba viendo, 
        'Poner un timer que luego de cierto tiempo de inactividad se cierre solo!!
        'Esto aplica para nodos de catalogos, de templates y de records!!!
        'Aun asi, poner una regla que luego de cierto tiempo de inactividad pueda entrar!!
        'Y verificar que el que lo este usando no sea YO MISMO!!
        'Al seleccionar un nodo debe hacer 2 cosas,
        'OJO, si es Viewer o Admin NO aplica nada!!
        'PRIMERO checar que NO se este usando por alguien mas
        'SEGUNDO, si NO se esta usando por nadie mas, se postea que se esta usando por X usuario
        'TERCERO, si SI se esta usando por alguien mas, se debe bloquear la edición,
        toyTrabajando = False
        pathFinduse = ""
        puedoEditar = False
        lastNodeEdit = ""
        ToolStripLabel2.Text = ""
        editDs.Clear()
        Label3.Text = ""

        inUseDt.Columns.Clear()
        inUseDt.Rows.Clear()

        inUseDt.Columns.Add("LastUsed")
        inUseDt.Columns.Add("TableName")
        inUseDt.Columns.Add("User")
        inUseDt.Columns.Add("UserName")
        inUseDt.Columns.Add("inUse")
        inUseDt.ExtendedProperties.Add("inEdit", False)
        inUseDt.ExtendedProperties.Add("Key", "")
        inUseDt.ExtendedProperties.Add("lastPath", "")

        inUseDt.Rows.Clear()
        inUseDt.Rows.Add({"", "", "", "", ""}) 'vacío!

        If RoleUsuario = "Viewer" Or RoleUsuario = "Admin" Then

            ToolStripButton1.Enabled = False
            ToolStripButton4.Enabled = False
            ToolStripButton5.Enabled = False
            ToolStripButton6.Enabled = False
            ToolStripButton7.Enabled = False
            ToolStripButton8.Enabled = False
            ToolStripButton10.Enabled = False
            ToolStripButton11.Enabled = False
            ToolStripButton14.Enabled = False
            ToolStripButton15.Enabled = False
            ToolStripButton16.Enabled = False

        Else
            ToolStripButton1.Enabled = True
            ToolStripButton4.Enabled = True
            ToolStripButton5.Enabled = True
            ToolStripButton6.Enabled = True
            ToolStripButton7.Enabled = True
            ToolStripButton8.Enabled = True
            ToolStripButton10.Enabled = True
            ToolStripButton11.Enabled = True
            ToolStripButton14.Enabled = True
            ToolStripButton15.Enabled = True
            ToolStripButton16.Enabled = True

            Module5.Start(Date.Now.AddMinutes(10))
            AddHandler Module5.Elapsed, AddressOf HandleElapsed
            AddHandler Module5.TimerComplete, AddressOf HandleComplete
            Label3.ForeColor = Color.Black
            'ToolStripLabel21.Font = New System.Drawing.Font("Calibri", 10, FontStyle.Bold)
            'ToolStripLabel21.ForeColor = Color.Black

        End If


        ToolStripComboBox1.Items.Clear()
        ToolStripComboBox1.Items.Add("Select") '0
        ToolStripComboBox1.Items.Add("Catalogs") '1
        ToolStripComboBox1.Items.Add("Dependencies") '2
        ToolStripComboBox1.Items.Add("Records") '3
        ToolStripComboBox1.Items.Add("Templates") '4
        ToolStripComboBox1.Items.Add("All dependencies") '5

        arbolDs.Tables.Clear()
        arbolDs.Tables.Add("Tree")

        'ToolStripButton10.Enabled = False
        'ToolStripButton11.Enabled = False

        Dim iAveprim(1) As DataColumn
        Dim kEys As New DataColumn()
        kEys.ColumnName = "Key"
        iAveprim(0) = kEys

        arbolDs.Tables(0).Columns.Add(kEys) 'Llave> tm#md32#md32-0001#LAND1'0
        arbolDs.Tables(0).Columns.Add("DepModule") '1
        arbolDs.Tables(0).Columns.Add("DepObject") '2
        arbolDs.Tables(0).Columns.Add("DepTable") '3
        arbolDs.Tables(0).Columns.Add("DepField") '4

        arbolDs.Tables(0).Columns.Add("DepModuleName") '5
        arbolDs.Tables(0).Columns.Add("DepObjectName") '6
        arbolDs.Tables(0).Columns.Add("DepTableName") '7
        arbolDs.Tables(0).Columns.Add("DepFieldName") '8

        arbolDs.Tables(0).Columns.Add("ConKey") '9'tm#md32#md32-0001#LAND1'0

        arbolDs.Tables(0).Columns.Add("ConModule") '10
        arbolDs.Tables(0).Columns.Add("ConObject") '11
        arbolDs.Tables(0).Columns.Add("ConTable") '12
        arbolDs.Tables(0).Columns.Add("ConField") '13

        arbolDs.Tables(0).Columns.Add("ConModuleName") '14
        arbolDs.Tables(0).Columns.Add("ConObjectName") '15
        arbolDs.Tables(0).Columns.Add("ConTableName") '16
        arbolDs.Tables(0).Columns.Add("ConFieldName") '17

        arbolDs.Tables(0).PrimaryKey = iAveprim

        usaDataset.Tables.Clear()
        usaDataset.Tables.Add("arbol")
        usaDataset.Tables(0).Columns.Add("address")

        writeDs.Tables.Clear()
        writeDs.Tables.Add("Levels")
        writeDs.Tables(0).Columns.Add("Deep") '0

        writeDs.Tables.Add("UpdNewRecords") '1
        writeDs.Tables.Add("DeletedRecords") '2
        writeDs.Tables.Add("CurrentRecords") '3
        writeDs.Tables.Add("UpdatedRecords") '4
        'writeDs.Tables(1).Columns.Add("")

        Dim YavePrim(1) As DataColumn
        Dim Llaves As New DataColumn()
        Llaves.ColumnName = "CatalogCode"
        YavePrim(0) = Llaves

        CatSimple.Tables.Clear()
        CatSimple.Tables.Add("CatSimple")
        CatSimple.Tables(0).Columns.Add(Llaves) 'catalogCode
        CatSimple.Tables(0).Columns.Add("Module") 'Module
        CatSimple.Tables(0).Columns.Add("CatalogName") 'Name
        CatSimple.Tables(0).PrimaryKey = YavePrim

        TreeView1.ImageList = ImageList1


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

        Call SetDoubleBuffered(TableLayoutPanel1)
        Call SetDoubleBuffered(DataGridView1)
        Call SetDoubleBuffered(TreeView1)
        Call SetDoubleBuffered(ToolStrip1)
        Call SetDoubleBuffered(ToolStrip2)
        Call SetDoubleBuffered(ToolStrip3)
        Call SetDoubleBuffered(SplitContainer1)

        Call CreaDBLetraNumero()
        'ToolStripProgressBar1.Visible = False
        Call ReloadModules()

        Call LoadTablaCatalogos()

        Label1.Text = "No object selected"
        Label2.Text = ""

        '1 a 1 puede ser con CSV, muchos a muchos es con Interop Excel o con OleDb
        'Mucho a muchos export es con Interop Excel NO hay mas!

        'Upload from excel
        'https://www.freecodespot.com/blog/csharp-import-excel/
        'https://www.c-sharpcorner.com/UploadFile/ankurmee/import-data-from-excel-to-datagridview-in-C-Sharp/
        'https://www.csharp-console-examples.com/winform/c-program-to-read-csv-file-and-display-data-in-datagridview/
        'https://stackoverflow.com/questions/16293643/import-excel-to-datagridview


        'ftp read write:
        'Application.StartupPath
        'https://www.codeguru.com/visual-basic/ftp-and-vb-net/
        'http://vbcity.com/blogs/xtab/archive/2016/04/13/how-to-upload-and-download-files-with-ftp-from-a-vb-net-application.aspx
        'https://newbedev.com/upload-file-to-ftp-site-using-vb-net
        'https://www.codeproject.com/Questions/1213460/How-do-I-download-a-folder-from-ftp-VB-NET
        'https://stackoverflow.com/questions/46363829/download-all-files-and-sub-directories-from-ftp-folder-in-vb-net

        'https://stackoverflow.com/questions/52030168/publish-winform-on-ftp-with-subsequent-updates

        'improve writting:
        'https://firebase.google.com/docs/database/rest/save-data

        'https://forum.asana.com/t/excel-vba-api-token-authentication/49943
        'https://stackoverflow.com/questions/22149169/how-to-pass-authentication-credentials-in-vba
        'https://firebase.google.com/docs/database/rest/auth#node.js
        'https://firebase.google.com/docs/auth/admin/verify-id-tokens#retrieve_id_tokens_on_clients
        'https://developers.google.com/identity/protocols/oauth2

        'https://firebase.google.com/docs/database/web/structure-data

        'Tree Slow:
        'https://forums.asp.net/t/1783420.aspx?ASP+NET+TreeView+Working+Too+Slow
        'https://stackoverflow.com/questions/25804276/slow-loading-of-treeview
        'https://bytes.com/topic/visual-basic-net/answers/570932-loading-treeview-dynamically-very-slow
        'https://social.msdn.microsoft.com/Forums/en-US/4211833f-9dd0-4675-927a-5f8999e0e066/treeview-numbering-slow?forum=vbgeneral

        'Cargar en Datagrid con virtualization
        'https://docs.microsoft.com/en-us/dotnet/desktop/winforms/controls/implementing-virtual-mode-wf-datagridview-control?view=netframeworkdesktop-4.8&redirectedfrom=MSDN

        'Que tienen que Ver:
        '1- Firebase, estructura
        '2- Catalogos, como se llenan, como interactuan con el template
        '3- Templates- como se llenan
        '4- Dependencias, como se agregan
        '5- Registros, como se suben con la App, como se descargan con el template

        rowEstaba = -1
        puSSyCat = -1
        cataNombre = ""

    End Sub

    Private Async Sub ToolStripComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged

        Dim seCarga As Boolean = False
        Dim suReg As String = ""
        rowEstaba = -1

        DoneEvaluacion = False

        Select Case ToolStripComboBox1.SelectedIndex
            Case Is = 0
                'se eliminar el tree, se deshabilita todo

                If RoleUsuario = "Editor" Then
                    'ToolStripButton10.Enabled = False
                    'ToolStripButton11.Enabled = False
                End If

                CategSelected = 0
                elNode = Nothing
                TreeView1.Nodes.Clear()
                NodoNameActual = ""

            Case Is = 1 'catalogos

                If RoleUsuario = "Editor" Then
                    'ToolStripButton10.Enabled = False
                    'ToolStripButton11.Enabled = False
                End If

                CategSelected = 1
                If catDs.Tables.Count = 0 Then seCarga = True

            Case Is = 2 'dependencias

                If RoleUsuario = "Editor" Then
                    'ToolStripButton10.Enabled = False
                    'ToolStripButton11.Enabled = False
                End If

                CategSelected = 2
                If depeDs.Tables.Count = 0 Then seCarga = True

            Case Is = 3 'records
                If RoleUsuario = "Editor" Then
                    'ToolStripButton10.Enabled = True
                    'ToolStripButton11.Enabled = True
                End If

                CategSelected = 3
                If recDs.Tables.Count = 0 Then seCarga = True

            Case Is = 4 'Templates
                If RoleUsuario = "Editor" Then
                    'ToolStripButton10.Enabled = False
                    'ToolStripButton11.Enabled = False
                End If

                CategSelected = 4
                If tempDs.Tables.Count = 0 Or depeDs.Tables.Count = 0 Then seCarga = True

            Case Is = 5
                If RoleUsuario = "Editor" Then
                    'ToolStripButton10.Enabled = False
                    'ToolStripButton11.Enabled = False
                End If
                CategSelected = 5
                If depeDs.Tables.Count = 0 Then seCarga = True

        End Select

        If seCarga = True Then
            suReg = Await CargaOpcion(ToolStripComboBox1.SelectedIndex)
        Else
            suReg = "ok"
        End If

        If suReg = "ok" Then
            Call CargaHeadersDg(ToolStripComboBox1.SelectedIndex)
            Call MuestraTree(ToolStripComboBox1.SelectedIndex)
        End If

    End Sub

    Private Sub LoadTablaCatalogos()

        niuTabcols.Columns.Clear()
        niuTabcols.Columns.Add("FieldCode", GetType(String)) '0
        niuTabcols.Columns.Add("FieldName", GetType(String)) '1
        niuTabcols.Columns.Add("Description", GetType(String)) '2
        niuTabcols.Columns.Add("Position", GetType(Integer)) '3
        niuTabcols.Columns.Add("isText", GetType(String)) '4

        TabColumnas.Columns.Clear()
        TabColumnas.Columns.Add("FieldCode", GetType(String)) '0
        TabColumnas.Columns.Add("FieldName", GetType(String)) '1
        TabColumnas.Columns.Add("Description", GetType(String)) '2
        TabColumnas.Columns.Add("Position", GetType(Integer)) '3
        TabColumnas.Columns.Add("isText", GetType(String)) '4

    End Sub

    Private Async Sub ReloadModules()

        ModuloDs.Clear()
        usaDataset.Tables(0).Rows.Clear()
        usaDataset.Tables(0).Rows.Add({"modules"}) '0
        ModuloDs = Await PullUrlWs(usaDataset, "modules")

    End Sub



    Private Sub MuestraTree(ByVal xOpcion As Integer)

        Dim xObj As Object = Nothing
        Dim posi As Integer = 0
        Dim pos1 As Integer = 0
        Dim pos2 As Integer = 0
        Dim pos3 As Integer = 0
        Dim xYave As String = ""

        Select Case xOpcion

            Case Is = 0


            Case Is = 1

                MySource.Clear()

                writeDs.Tables(1).Rows.Clear()
                writeDs.Tables(2).Rows.Clear()
                writeDs.Tables(3).Rows.Clear()
                writeDs.Tables(4).Rows.Clear()

                writeDs.Tables(1).Columns.Clear() 'new
                writeDs.Tables(2).Columns.Clear() 'deleted
                writeDs.Tables(3).PrimaryKey = Nothing
                writeDs.Tables(3).Columns.Clear() 'current
                writeDs.Tables(4).Columns.Clear() 'deleted


                Dim iAveprim(1) As DataColumn
                Dim kEys As New DataColumn()
                kEys.ColumnName = "ID"
                iAveprim(0) = kEys

                'agregar campo Llave a los catalogos!
                'si lo borro, 
                'este solo debe tener registros nuevos
                writeDs.Tables(1).Columns.Add("Key")
                writeDs.Tables(1).Columns.Add("Description")
                writeDs.Tables(1).Columns.Add("FireKey")

                writeDs.Tables(2).Columns.Add("Key")
                writeDs.Tables(2).Columns.Add("Description")
                writeDs.Tables(2).Columns.Add("FireKey")

                writeDs.Tables(3).Columns.Add(kEys)
                writeDs.Tables(3).Columns.Add("Description")
                writeDs.Tables(3).Columns.Add("FireKey")
                writeDs.Tables(3).CaseSensitive = True
                writeDs.Tables(3).PrimaryKey = iAveprim

                writeDs.Tables(4).Columns.Add("Key") 'updated!
                writeDs.Tables(4).Columns.Add("Description")
                writeDs.Tables(4).Columns.Add("FireKey")

                'se busca vs lo nuevo, campo llave, si existe, se compara su descripcion, si sigue igual NO se agrega
                TreeView1.SuspendLayout()
                TreeView1.BeginUpdate()

                TreeView1.Nodes.Clear()

                'Cursor.Current = Cursors.WaitCursor

                'ToolStripProgressBar2.Value = 0
                'ToolStripProgressBar2.Maximum = catDs.Tables.Count
                'ToolStripLabel2.Text = "Building..."
                'ToolStripProgressBar2.Visible = True
                'ToolStripLabel2.Visible = True

                Label1.Text = "No object selected"
                Label2.Text = ""

                'improve update

                For i = 0 To catDs.Tables.Count - 1
                    'ToolStripProgressBar2.Value = i + 1
                    xObj = Nothing
                    xObj = Split(catDs.Tables(i).TableName, "#")

                    posi = TreeView1.Nodes.IndexOfKey(CStr(xObj(0)))
                    If posi < 0 Then
                        TreeView1.Nodes.Add(CStr(xObj(0)), CStr(xObj(0)) & " / " & catDs.Tables(i).ExtendedProperties.Item("ModuleName"), 1, 1)
                        MySource.Add(CStr(xObj(0)))
                        posi = TreeView1.Nodes.Count - 1
                    End If

                    TreeView1.Nodes(posi).Nodes.Add(CStr(xObj(1)), catDs.Tables(i).ExtendedProperties.Item("CatalogName"), 1, 1)

                    pos1 = TreeView1.Nodes(posi).Nodes.Count - 1

                    MySource.Add(CStr(xObj(1)))
                    MySource.Add(catDs.Tables(i).ExtendedProperties.Item("CatalogName"))

                Next

                'For i = 0 To catDs.Tables.Count - 1
                '    ToolStripProgressBar2.Value = i + 1

                '    xObj = Nothing
                '    xObj = Split(catDs.Tables(i).TableName, "#")

                '    posi = TreeView1.Nodes.IndexOfKey(CStr(xObj(0)))
                '    If posi < 0 Then
                '        TreeView1.Nodes.Add(CStr(xObj(0)), CStr(xObj(0)), 1, 1)
                '        MySource.Add(CStr(xObj(0)))
                '        posi = TreeView1.Nodes.Count - 1
                '    End If

                '    TreeView1.Nodes(posi).Nodes.Add(CStr(xObj(1)), catDs.Tables(i).Columns(0).ColumnName, 1, 1)

                '    pos1 = TreeView1.Nodes(posi).Nodes.Count - 1

                '    MySource.Add(CStr(xObj(1)))
                '    MySource.Add(catDs.Tables(i).Columns(0).ColumnName)

                '    'Dim subArbol(catDs.Tables(i).Rows.Count - 1) As TreeNode

                '    'For j = 0 To catDs.Tables(i).Rows.Count - 1
                '    '    subArbol(j) = New TreeNode
                '    '    subArbol(j).Name = CStr(catDs.Tables(i).Rows(j).Item(0))
                '    '    subArbol(j).Text = CStr(catDs.Tables(i).Rows(j).Item(0) & " / " & CStr(catDs.Tables(i).Rows(j).Item(1)))
                '    '    MySource.Add(CStr(catDs.Tables(i).Rows(j).Item(0)))
                '    '    MySource.Add(CStr(catDs.Tables(i).Rows(j).Item(1)))
                '    'Next

                '    'TreeView1.Nodes(posi).Nodes(pos1).Nodes.AddRange(subArbol)

                'Next

                'ToolStripLabel2.Text = "Ready"
                'ToolStripProgressBar2.Value = 0
                'ToolStripProgressBar2.Visible = False
                'ToolStripLabel2.Visible = False

                'Cursor.Current = Cursors.Default

                TreeView1.EndUpdate()
                TreeView1.ResumeLayout()

                ToolStripTextBox1.AutoCompleteCustomSource = MySource
                ToolStripTextBox1.AutoCompleteMode = AutoCompleteMode.Suggest
                ToolStripTextBox1.AutoCompleteSource = AutoCompleteSource.CustomSource

            Case Is = 2

                Label1.Text = "No object selected"
                Label2.Text = ""

                MySource.Clear()

                writeDs.Tables(1).Rows.Clear()
                writeDs.Tables(2).Rows.Clear()
                writeDs.Tables(3).Rows.Clear()

                'aqui NO hay llave primaria!
                writeDs.Tables(1).Columns.Clear() 'new
                writeDs.Tables(2).Columns.Clear() 'deleted
                writeDs.Tables(3).PrimaryKey = Nothing
                writeDs.Tables(3).Columns.Clear() 'current

                Dim iAveprim(1) As DataColumn
                Dim kEys As New DataColumn()
                kEys.ColumnName = "Unike"
                iAveprim(0) = kEys

                writeDs.Tables(1).Columns.Add("DepFieldCode") '0
                writeDs.Tables(1).Columns.Add("DepFieldName") '1
                writeDs.Tables(1).Columns.Add("ConFieldCode") '2
                writeDs.Tables(1).Columns.Add("ConFieldName") '3
                writeDs.Tables(1).Columns.Add("ConFieldModule") '4
                writeDs.Tables(1).Columns.Add("ConFieldObject") '5
                writeDs.Tables(1).Columns.Add("ConFieldTableCode") '6
                writeDs.Tables(1).Columns.Add("ConFieldTableName") '7


                writeDs.Tables(2).Columns.Add("DepFieldCode") '0
                writeDs.Tables(2).Columns.Add("DepFieldName") '1
                writeDs.Tables(2).Columns.Add("ConFieldCode") '2
                writeDs.Tables(2).Columns.Add("ConFieldName") '3
                writeDs.Tables(2).Columns.Add("ConFieldModule") '4
                writeDs.Tables(2).Columns.Add("ConFieldObject") '5
                writeDs.Tables(2).Columns.Add("ConFieldTableCode") '6
                writeDs.Tables(2).Columns.Add("ConFieldTableName") '7

                writeDs.Tables(3).Columns.Add(kEys) '0
                writeDs.Tables(3).Columns.Add("DepFieldCode") '1
                writeDs.Tables(3).Columns.Add("DepFieldName") '2
                writeDs.Tables(3).Columns.Add("ConFieldCode") '3
                writeDs.Tables(3).Columns.Add("ConFieldName") '4
                writeDs.Tables(3).Columns.Add("ConFieldModule") '5
                writeDs.Tables(3).Columns.Add("ConFieldObject") '6
                writeDs.Tables(3).Columns.Add("ConFieldTableCode") '7
                writeDs.Tables(3).Columns.Add("ConFieldTableName") '8

                writeDs.Tables(3).PrimaryKey = iAveprim

                TreeView1.BeginUpdate()

                TreeView1.Nodes.Clear()

                TreeView1.Nodes.Add("root", "All dependencies", 2, 2)

                Cursor.Current = Cursors.WaitCursor

                For i = 0 To depeDs.Tables.Count - 1
                    xObj = Nothing
                    xObj = Split(depeDs.Tables(i).TableName, "#") 'md01#company master

                    posi = TreeView1.Nodes(0).Nodes.IndexOfKey(CStr(xObj(0)))
                    If posi < 0 Then
                        TreeView1.Nodes(0).Nodes.Add(CStr(xObj(0)), CStr(xObj(0)) & " / " & CStr(xObj(1)), 2, 2)
                        MySource.Add(CStr(xObj(0)))
                        posi = TreeView1.Nodes(0).Nodes.Count - 1
                        TreeView1.Nodes(0).Nodes(posi).Tag = CStr(xObj(1))
                    End If

                    For j = 0 To depeDs.Tables(i).Rows.Count - 1
                        'verificamos si existe!
                        pos1 = TreeView1.Nodes(0).Nodes(posi).Nodes.IndexOfKey(CStr(depeDs.Tables(i).Rows(j).Item(0)))
                        'pos1=posicion de la tabla dependiente
                        If pos1 < 0 Then
                            'se agrega!
                            TreeView1.Nodes(0).Nodes(posi).Nodes.Add(CStr(depeDs.Tables(i).Rows(j).Item(0)), CStr(depeDs.Tables(i).Rows(j).Item(0)) & " / " & CStr(depeDs.Tables(i).Rows(j).Item(1)), 8, 8)
                            pos1 = TreeView1.Nodes(0).Nodes(posi).Nodes.Count - 1
                            TreeView1.Nodes(0).Nodes(posi).Nodes(pos1).Tag = CStr(depeDs.Tables(i).Rows(j).Item(1))
                            MySource.Add(CStr(depeDs.Tables(i).Rows(j).Item(0)))
                            MySource.Add(CStr(depeDs.Tables(i).Rows(j).Item(1)))
                        End If

                        pos2 = TreeView1.Nodes(0).Nodes(posi).Nodes(pos1).Nodes.IndexOfKey(CStr(depeDs.Tables(i).Rows(j).Item(2)))
                        If pos2 < 0 Then
                            'lo agregamos!
                            TreeView1.Nodes(0).Nodes(posi).Nodes(pos1).Nodes.Add(CStr(depeDs.Tables(i).Rows(j).Item(2)), CStr(depeDs.Tables(i).Rows(j).Item(2)) & " / " & CStr(depeDs.Tables(i).Rows(j).Item(3)), 4, 4)
                            pos2 = TreeView1.Nodes(0).Nodes(posi).Nodes(pos1).Nodes.Count - 1
                            MySource.Add(CStr(depeDs.Tables(i).Rows(j).Item(2)))
                            MySource.Add(CStr(depeDs.Tables(i).Rows(j).Item(3)))
                        End If

                        TreeView1.Nodes(0).Nodes(posi).Nodes(pos1).Nodes(pos2).Nodes.Add(CStr(depeDs.Tables(i).Rows(j).Item(7)), "Object: " & CStr(depeDs.Tables(i).Rows(j).Item(7)), 4, 4)
                        TreeView1.Nodes(0).Nodes(posi).Nodes(pos1).Nodes(pos2).Nodes.Add(CStr(depeDs.Tables(i).Rows(j).Item(6)), "Module: " & CStr(depeDs.Tables(i).Rows(j).Item(6)), 4, 4)
                        TreeView1.Nodes(0).Nodes(posi).Nodes(pos1).Nodes(pos2).Nodes.Add(CStr(depeDs.Tables(i).Rows(j).Item(8)), "Table: " & CStr(depeDs.Tables(i).Rows(j).Item(8)) & " / " & CStr(depeDs.Tables(i).Rows(j).Item(9)), 4, 4)
                        TreeView1.Nodes(0).Nodes(posi).Nodes(pos1).Nodes(pos2).Nodes.Add(CStr(depeDs.Tables(i).Rows(j).Item(4)), "Field: " & CStr(depeDs.Tables(i).Rows(j).Item(4)) & "/" & CStr(depeDs.Tables(i).Rows(j).Item(5)), 4, 4)
                        'yave formado por: object code + table code + field code
                        MySource.Add(CStr(depeDs.Tables(i).Rows(j).Item(7)))
                        MySource.Add(CStr(depeDs.Tables(i).Rows(j).Item(6)))
                        MySource.Add(CStr(depeDs.Tables(i).Rows(j).Item(8)))
                        MySource.Add(CStr(depeDs.Tables(i).Rows(j).Item(9)))
                        MySource.Add(CStr(depeDs.Tables(i).Rows(j).Item(4)))
                        MySource.Add(CStr(depeDs.Tables(i).Rows(j).Item(5)))


                    Next

                    'TreeView1.Nodes(CStr(xObj(0))).Nodes.Add(CStr(xObj(1)), catDs.Tables(i).Columns(0).ColumnName, 1, 1)

                Next

                Cursor.Current = Cursors.Default

                TreeView1.EndUpdate()

                ToolStripTextBox1.AutoCompleteCustomSource = MySource
                ToolStripTextBox1.AutoCompleteMode = AutoCompleteMode.Suggest
                ToolStripTextBox1.AutoCompleteSource = AutoCompleteSource.CustomSource

            Case Is = 3
                'records
                'cargamos las compañias, de acuerdo al catálogo gb0007
                MySource.Clear()
                TreeView1.BeginUpdate()

                TreeView1.Nodes.Clear()
                TreeView1.Nodes.Add("root", "All companies", 2, 2)

                For i = 0 To catDs.Tables.Count - 1
                    If catDs.Tables(i).TableName = "gb#gb0007" Then
                        'la encontramos!
                        For j = 0 To catDs.Tables(i).Rows.Count - 1
                            If catDs.Tables(i).Rows(j).Item(0) = "CatalogName" Then Continue For
                            TreeView1.Nodes(0).Nodes.Add(catDs.Tables(i).Rows(j).Item("A"), catDs.Tables(i).Rows(j).Item("A") & " / " & catDs.Tables(i).Rows(j).Item("ZZ"), 2, 2)
                            TreeView1.Nodes(0).Nodes(TreeView1.Nodes(0).Nodes.Count - 1).Tag = catDs.Tables(i).Rows(j).Item("ZZ")
                            MySource.Add(catDs.Tables(i).Rows(j).Item("A"))
                            MySource.Add(catDs.Tables(i).Rows(j).Item("ZZ"))
                        Next

                        Exit For

                    End If
                Next

                'ahora ciclamos por todos los templates y los agregamos a todas los segundos nodos!
                For i = 0 To TreeView1.Nodes(0).Nodes.Count - 1
                    For j = 0 To tempDs.Tables.Count - 1
                        xObj = Split(tempDs.Tables(j).TableName, "#")
                        TreeView1.Nodes(0).Nodes(i).Nodes.Add(CStr(xObj(0)), CStr(xObj(0)) & " / " & CStr(xObj(1)), 2, 2)
                        TreeView1.Nodes(0).Nodes(i).Nodes(TreeView1.Nodes(0).Nodes(i).Nodes.Count - 1).Tag = CStr(xObj(1))
                        posi = TreeView1.Nodes(0).Nodes(i).Nodes.Count - 1
                        MySource.Add(CStr(xObj(0)))
                        MySource.Add(CStr(xObj(1)))

                        For w = 0 To tempDs.Tables(j).Rows.Count - 1
                            'If MySource.IndexOf(CStr(tempDs.Tables(j).Rows(w).Item(1))) >= 0 Then Continue For

                            If TreeView1.Nodes(0).Nodes(i).Nodes(posi).Nodes.IndexOfKey(tempDs.Tables(j).Rows(w).Item(1)) >= 0 Then Continue For

                            TreeView1.Nodes(0).Nodes(i).Nodes(posi).Nodes.Add(tempDs.Tables(j).Rows(w).Item(1), tempDs.Tables(j).Rows(w).Item(1) & " / " & tempDs.Tables(j).Rows(w).Item(2), 8, 8)
                            TreeView1.Nodes(0).Nodes(i).Nodes(posi).Nodes(TreeView1.Nodes(0).Nodes(i).Nodes(posi).Nodes.Count - 1).Tag = tempDs.Tables(j).Rows(w).Item(2)

                            If MySource.IndexOf(CStr(tempDs.Tables(j).Rows(w).Item(1))) < 0 Then
                                MySource.Add(CStr(tempDs.Tables(j).Rows(w).Item(1)))
                                MySource.Add(CStr(tempDs.Tables(j).Rows(w).Item(2)))
                            End If

                        Next

                    Next
                Next

                TreeView1.EndUpdate()

            Case Is = 4
                'templates
                Label1.Text = "No object selected"
                Label2.Text = ""

                writeDs.Tables(1).Rows.Clear()
                writeDs.Tables(2).Rows.Clear()
                writeDs.Tables(3).Rows.Clear()

                'aqui NO hay llave primaria!
                writeDs.Tables(1).Columns.Clear() 'new
                writeDs.Tables(2).Columns.Clear() 'deleted
                writeDs.Tables(3).PrimaryKey = Nothing
                writeDs.Tables(3).Columns.Clear() 'current

                Dim iAveprim(1) As DataColumn
                Dim kEys As New DataColumn()
                kEys.ColumnName = "FieldCode"
                iAveprim(0) = kEys

                writeDs.Tables(1).Columns.Add("FieldCode") '0
                writeDs.Tables(1).Columns.Add("FieldName") '1
                writeDs.Tables(1).Columns.Add("isKey") '2
                writeDs.Tables(1).Columns.Add("Position") '3
                writeDs.Tables(1).Columns.Add("Letter") '4

                writeDs.Tables(2).Columns.Add("FieldCode") '0
                writeDs.Tables(2).Columns.Add("FieldName") '1
                writeDs.Tables(2).Columns.Add("isKey") '2
                writeDs.Tables(2).Columns.Add("Position") '3
                writeDs.Tables(2).Columns.Add("Letter") '4

                writeDs.Tables(3).Columns.Add(kEys) '0
                'writeDs.Tables(3).Columns.Add("FieldCode") '1
                writeDs.Tables(3).Columns.Add("Name") '1
                writeDs.Tables(3).Columns.Add("isKey") '2
                writeDs.Tables(3).Columns.Add("Position") '3
                writeDs.Tables(3).Columns.Add("Letter") '4

                writeDs.Tables(3).Columns.Add("MOC") '5
                writeDs.Tables(3).Columns.Add("FillingRule") '6
                writeDs.Tables(3).Columns.Add("DataType") '7
                writeDs.Tables(3).Columns.Add("MaxChar") '8
                writeDs.Tables(3).Columns.Add("ULCase") '9
                writeDs.Tables(3).Columns.Add("Blanks") '10
                writeDs.Tables(3).Columns.Add("CatalogName") '11
                writeDs.Tables(3).Columns.Add("CatalogCode") '12
                writeDs.Tables(3).Columns.Add("ValueColumn") '13
                writeDs.Tables(3).Columns.Add("NonRep") '14
                writeDs.Tables(3).Columns.Add("NonAllowedChars") '15

                'writeDs.Tables(3).Columns.Add("ConditionalPath") '17
                'writeDs.Tables(3).Columns.Add("ConditionalObject") '18
                'writeDs.Tables(3).Columns.Add("ConditionalTable") '19
                'writeDs.Tables(3).Columns.Add("ConditionalField") '20
                'writeDs.Tables(3).Columns.Add("ConditionalType") '21
                'writeDs.Tables(3).Columns.Add("ConditionalRule") '22
                'writeDs.Tables(3).Columns.Add("ConditionalValue") '23
                'writeDs.Tables(3).Columns.Add("ConstructionRule") '24

                writeDs.Tables(3).PrimaryKey = iAveprim

                MySource.Clear()
                TreeView1.BeginUpdate()
                TreeView1.Nodes.Clear()

                TreeView1.Nodes.Add("root", "All templates", 2, 2)

                Cursor.Current = Cursors.WaitCursor

                For i = 0 To tempDs.Tables.Count - 1
                    xObj = Nothing
                    xObj = Split(tempDs.Tables(i).TableName, "#") 'md01#company master

                    posi = TreeView1.Nodes(0).Nodes.IndexOfKey(CStr(xObj(0)))
                    If posi < 0 Then
                        TreeView1.Nodes(0).Nodes.Add(CStr(xObj(0)), CStr(xObj(0)) & " / " & CStr(xObj(1)), 2, 2)
                        MySource.Add(CStr(xObj(0)))
                        posi = TreeView1.Nodes(0).Nodes.Count - 1
                        TreeView1.Nodes(0).Nodes(posi).Tag = CStr(xObj(1)) & "#" & CStr(xObj(2))
                    End If

                    tempDs.Tables(i).DefaultView.Sort = "Position ASC"

                    Dim filterDT As DataTable = tempDs.Tables(i).DefaultView.ToTable()

                    For j = 0 To filterDT.Rows.Count - 1

                        pos1 = TreeView1.Nodes(0).Nodes(posi).Nodes.IndexOfKey(CStr(filterDT.Rows(j).Item(1)))
                        'pos1=posicion de la tabla dependiente
                        If pos1 < 0 Then
                            'se agrega!
                            TreeView1.Nodes(0).Nodes(posi).Nodes.Add(CStr(filterDT.Rows(j).Item(1)), CStr(filterDT.Rows(j).Item(1)) & " / " & CStr(filterDT.Rows(j).Item(2)), 8, 8)
                            pos1 = TreeView1.Nodes(0).Nodes(posi).Nodes.Count - 1
                            MySource.Add(CStr(filterDT.Rows(j).Item(1)))
                            MySource.Add(CStr(filterDT.Rows(j).Item(2)))
                            TreeView1.Nodes(0).Nodes(posi).Nodes(pos1).Tag = CStr(filterDT.Rows(j).Item(2))
                        End If

                        'If CStr(filterDT.Rows(j).Item(3)) = "TableName" Then Continue For 'no se agrega al nodo!

                        TreeView1.Nodes(0).Nodes(posi).Nodes(pos1).Nodes.Add(CStr(filterDT.Rows(j).Item(3)), CStr(filterDT.Rows(j).Item(3)) & " / " & CStr(filterDT.Rows(j).Item(4)), 8, 8)
                        MySource.Add(CStr(filterDT.Rows(j).Item(3)))
                        MySource.Add(CStr(filterDT.Rows(j).Item(4)))

                    Next

                Next

                Cursor.Current = Cursors.Default

                TreeView1.EndUpdate()

                ToolStripTextBox1.AutoCompleteCustomSource = MySource
                ToolStripTextBox1.AutoCompleteMode = AutoCompleteMode.Suggest
                ToolStripTextBox1.AutoCompleteSource = AutoCompleteSource.CustomSource

            Case Is = 5
                'arbol completo!!

                MySource.Clear()

                arbolDs.Tables(0).Rows.Clear() 'empezamos!

                For i = 0 To depeDs.Tables.Count - 1
                    'cada nodo un template!
                    xObj = Nothing
                    xObj = Split(depeDs.Tables(i).TableName, "#")

                    For j = 0 To depeDs.Tables(i).Rows.Count - 1
                        xYave = xObj(0) & "#" & CStr(depeDs.Tables(i).Rows(j).Item(0)) & "#" & CStr(depeDs.Tables(i).Rows(j).Item(2))
                        arbolDs.Tables(0).Rows.Add({xYave, ""})
                        arbolDs.Tables(0).Rows(arbolDs.Tables(0).Rows.Count - 1).Item(1) = "" 'dependent Module
                        arbolDs.Tables(0).Rows(arbolDs.Tables(0).Rows.Count - 1).Item(2) = xObj(0) 'object/template
                        arbolDs.Tables(0).Rows(arbolDs.Tables(0).Rows.Count - 1).Item(3) = CStr(depeDs.Tables(i).Rows(j).Item(0)) 'table
                        arbolDs.Tables(0).Rows(arbolDs.Tables(0).Rows.Count - 1).Item(4) = CStr(depeDs.Tables(i).Rows(j).Item(2)) 'field

                        arbolDs.Tables(0).Rows(arbolDs.Tables(0).Rows.Count - 1).Item(5) = "" 'dependent module name
                        arbolDs.Tables(0).Rows(arbolDs.Tables(0).Rows.Count - 1).Item(6) = xObj(1) '
                        arbolDs.Tables(0).Rows(arbolDs.Tables(0).Rows.Count - 1).Item(7) = CStr(depeDs.Tables(i).Rows(j).Item(1))
                        arbolDs.Tables(0).Rows(arbolDs.Tables(0).Rows.Count - 1).Item(8) = CStr(depeDs.Tables(i).Rows(j).Item(3))

                        arbolDs.Tables(0).Rows(arbolDs.Tables(0).Rows.Count - 1).Item(9) = CStr(depeDs.Tables(i).Rows(j).Item(7)) & "#" & CStr(depeDs.Tables(i).Rows(j).Item(8)) & "#" & CStr(depeDs.Tables(i).Rows(j).Item(4)) 'Yave Condicional

                        arbolDs.Tables(0).Rows(arbolDs.Tables(0).Rows.Count - 1).Item(10) = CStr(depeDs.Tables(i).Rows(j).Item(6)) 'conModule
                        arbolDs.Tables(0).Rows(arbolDs.Tables(0).Rows.Count - 1).Item(11) = CStr(depeDs.Tables(i).Rows(j).Item(7)) 'ConObject
                        arbolDs.Tables(0).Rows(arbolDs.Tables(0).Rows.Count - 1).Item(12) = CStr(depeDs.Tables(i).Rows(j).Item(8)) 'conTable
                        arbolDs.Tables(0).Rows(arbolDs.Tables(0).Rows.Count - 1).Item(13) = CStr(depeDs.Tables(i).Rows(j).Item(4)) 'conField
                        arbolDs.Tables(0).Rows(arbolDs.Tables(0).Rows.Count - 1).Item(14) = "" 'module Name
                        arbolDs.Tables(0).Rows(arbolDs.Tables(0).Rows.Count - 1).Item(15) = "" 'object name
                        arbolDs.Tables(0).Rows(arbolDs.Tables(0).Rows.Count - 1).Item(16) = CStr(depeDs.Tables(i).Rows(j).Item(9)) 'table name
                        arbolDs.Tables(0).Rows(arbolDs.Tables(0).Rows.Count - 1).Item(17) = CStr(depeDs.Tables(i).Rows(j).Item(5)) 'field name

                    Next

                Next

                'debemos ciclar vs tabla de templates para completar el arbol!
                For i = 0 To arbolDs.Tables(0).Rows.Count - 1

                    For j = 0 To tempDs.Tables.Count - 1

                        xObj = Nothing
                        xObj = Split(tempDs.Tables(j).TableName, "#")

                        If arbolDs.Tables(0).Rows(i).Item(2) = xObj(0) Then
                            arbolDs.Tables(0).Rows(i).Item(1) = xObj(2) 'su modulo!
                        End If

                        If arbolDs.Tables(0).Rows(i).Item(11) = xObj(0) Then
                            arbolDs.Tables(0).Rows(i).Item(15) = xObj(1) 'nombre del objeto!
                        End If

                    Next

                Next

                'por último agregamos la info!!
                TreeView1.BeginUpdate()
                TreeView1.Nodes.Clear()
                TreeView1.Nodes.Add("root", "Tree dependencies", 2, 2)

                For i = 0 To depeDs.Tables.Count - 1
                    xObj = Nothing
                    xObj = Split(depeDs.Tables(i).TableName, "#") 'md01#company master

                    posi = TreeView1.Nodes(0).Nodes.IndexOfKey(CStr(xObj(0)))
                    If posi < 0 Then
                        TreeView1.Nodes(0).Nodes.Add(CStr(xObj(0)), CStr(xObj(0)) & " / " & CStr(xObj(1)), 2, 2)
                        MySource.Add(CStr(xObj(0)))
                        posi = TreeView1.Nodes(0).Nodes.Count - 1
                        TreeView1.Nodes(0).Nodes(posi).Tag = CStr(xObj(1))
                    End If

                    For j = 0 To depeDs.Tables(i).Rows.Count - 1
                        'verificamos si existe!
                        pos1 = TreeView1.Nodes(0).Nodes(posi).Nodes.IndexOfKey(CStr(depeDs.Tables(i).Rows(j).Item(0)))
                        'pos1=posicion de la tabla dependiente
                        If pos1 < 0 Then
                            'se agrega!
                            TreeView1.Nodes(0).Nodes(posi).Nodes.Add(CStr(depeDs.Tables(i).Rows(j).Item(0)), CStr(depeDs.Tables(i).Rows(j).Item(0)) & " / " & CStr(depeDs.Tables(i).Rows(j).Item(1)), 8, 8)
                            pos1 = TreeView1.Nodes(0).Nodes(posi).Nodes.Count - 1
                            TreeView1.Nodes(0).Nodes(posi).Nodes(pos1).Tag = CStr(depeDs.Tables(i).Rows(j).Item(1))
                            MySource.Add(CStr(depeDs.Tables(i).Rows(j).Item(0)))
                            MySource.Add(CStr(depeDs.Tables(i).Rows(j).Item(1)))
                        End If

                        pos2 = TreeView1.Nodes(0).Nodes(posi).Nodes(pos1).Nodes.IndexOfKey(CStr(depeDs.Tables(i).Rows(j).Item(2)))
                        If pos2 < 0 Then
                            'lo agregamos!
                            TreeView1.Nodes(0).Nodes(posi).Nodes(pos1).Nodes.Add(CStr(depeDs.Tables(i).Rows(j).Item(2)), CStr(depeDs.Tables(i).Rows(j).Item(2)) & " / " & CStr(depeDs.Tables(i).Rows(j).Item(3)), 4, 4)
                            pos2 = TreeView1.Nodes(0).Nodes(posi).Nodes(pos1).Nodes.Count - 1
                            MySource.Add(CStr(depeDs.Tables(i).Rows(j).Item(2)))
                            MySource.Add(CStr(depeDs.Tables(i).Rows(j).Item(3)))
                        End If

                        TreeView1.Nodes(0).Nodes(posi).Nodes(pos1).Nodes(pos2).Nodes.Add(CStr(depeDs.Tables(i).Rows(j).Item(7)), "Object: " & CStr(depeDs.Tables(i).Rows(j).Item(7)), 4, 4)
                        TreeView1.Nodes(0).Nodes(posi).Nodes(pos1).Nodes(pos2).Nodes.Add(CStr(depeDs.Tables(i).Rows(j).Item(6)), "Module: " & CStr(depeDs.Tables(i).Rows(j).Item(6)), 4, 4)
                        TreeView1.Nodes(0).Nodes(posi).Nodes(pos1).Nodes(pos2).Nodes.Add(CStr(depeDs.Tables(i).Rows(j).Item(8)), "Table: " & CStr(depeDs.Tables(i).Rows(j).Item(8)) & " / " & CStr(depeDs.Tables(i).Rows(j).Item(9)), 4, 4)
                        TreeView1.Nodes(0).Nodes(posi).Nodes(pos1).Nodes(pos2).Nodes.Add(CStr(depeDs.Tables(i).Rows(j).Item(7)) & "#" & CStr(depeDs.Tables(i).Rows(j).Item(8)) & "#" & CStr(depeDs.Tables(i).Rows(j).Item(4)), "Field: " & CStr(depeDs.Tables(i).Rows(j).Item(4)) & "/" & CStr(depeDs.Tables(i).Rows(j).Item(5)), 4, 4)

                        'yave formado por: object code + table code + field code
                        MySource.Add(CStr(depeDs.Tables(i).Rows(j).Item(7)))
                        MySource.Add(CStr(depeDs.Tables(i).Rows(j).Item(6)))
                        MySource.Add(CStr(depeDs.Tables(i).Rows(j).Item(8)))
                        MySource.Add(CStr(depeDs.Tables(i).Rows(j).Item(9)))
                        MySource.Add(CStr(depeDs.Tables(i).Rows(j).Item(4)))
                        MySource.Add(CStr(depeDs.Tables(i).Rows(j).Item(5)))

                    Next

                Next


                For i = 0 To arbolDs.Tables(0).Rows.Count - 1

                    euReka = False
                    oNode = Nothing
                    elQbusca = CStr(arbolDs.Tables(0).Rows(i).Item(0))

                    Call buSkaNuevoNodo(TreeView1.Nodes(0))
                    If euReka = True Then
                        'ojo!, como lo vas a agregar!
                        oNode.Nodes.Add(arbolDs.Tables(0).Rows(i).Item(11), "Object: " & arbolDs.Tables(0).Rows(i).Item(11), 4, 4)
                        oNode.Nodes.Add(arbolDs.Tables(0).Rows(i).Item(10), "Module: " & arbolDs.Tables(0).Rows(i).Item(10), 4, 4)
                        oNode.Nodes.Add(arbolDs.Tables(0).Rows(i).Item(12), "Table: " & arbolDs.Tables(0).Rows(i).Item(12), 4, 4)
                        oNode.Nodes.Add(arbolDs.Tables(0).Rows(i).Item(9), "Field: " & arbolDs.Tables(0).Rows(i).Item(13), 4, 4)
                        'se agrega>el modulo>el template>la tabla, 
                        Exit For
                    End If

                Next

                TreeView1.EndUpdate()

        End Select


    End Sub

    Private Async Function CargaOpcion(ByVal laOpcion As Integer) As Task(Of String)

        'usaDataset.Tables(0).Clear()
        Dim elRet As String = ""
        usaDataset.Tables(0).Rows.Clear()

        Dim i As Integer
        Dim xObj As Object = Nothing
        Dim xSet As New DataSet
        Dim ySet As New DataSet
        Dim posi As Integer = 0

        Select Case laOpcion

            Case Is = 0
                'se elimina!
                elRet = "ok"


            Case Is = 1 'catalogos

                usaDataset.Tables(0).Rows.Add({"catpro"})
                catDs.Tables.Clear()
                catDs = Await PullUrlWs(usaDataset, "catpro")
                'aquii falta el sub para cargar los titulos del datagridview!
                ReloadCatPro()

                'asi estaba:
                'usaDataset.Tables(0).Rows.Add({"catalogs"})
                'catDs.Tables.Clear()
                'catDs = Await PullUrlWs(usaDataset, "catalogs")
                'aquii falta el sub para cargar los titulos del datagridview!
                'Call ReloadCatSimple()

                elRet = "ok"

            Case Is = 2 'dependencias
                usaDataset.Tables(0).Rows.Add({"dependencies"})
                depeDs.Tables.Clear()
                depeDs = Await PullUrlWs(usaDataset, "dependencies")

                elRet = "ok"

            Case Is = 3 'records, primero se jala de que compañía quiere ver
                'se sacan los catálogos, y templates!
                usaDataset.Tables(0).Rows.Add({"catpro"})
                catDs.Tables.Clear()
                catDs = Await PullUrlWs(usaDataset, "catpro")
                ReloadCatPro()
                'Call ReloadCatSimple()

                usaDataset.Tables(0).Rows.Clear()
                usaDataset.Tables(0).Rows.Add({"templates"})
                tempDs.Tables.Clear()
                tempDs = Await PullUrlWs(usaDataset, "templates")

                usaDataset.Tables(0).Rows.Clear()
                usaDataset.Tables(0).Rows.Add({"relations"})
                relaTionDs.Tables.Clear()
                relaTionDs = Await PullUrlWs(usaDataset, "relations")

                elRet = "ok"

            Case Is = 4 'Templates OJO, este es global!!

                'ahora tmb catalogos_
                usaDataset.Tables(0).Rows.Add({"catpro"})
                catDs.Tables.Clear()
                catDs = Await PullUrlWs(usaDataset, "catpro")
                ReloadCatPro()
                'Call ReloadCatSimple()

                usaDataset.Tables(0).Rows.Clear()
                usaDataset.Tables(0).Rows.Add({"dependencies"})
                depeDs.Tables.Clear()
                depeDs = Await PullUrlWs(usaDataset, "dependencies")

                usaDataset.Tables(0).Rows.Clear()
                usaDataset.Tables(0).Rows.Add({"templates"})
                tempDs.Tables.Clear()
                tempDs = Await PullUrlWs(usaDataset, "templates")

                usaDataset.Tables(0).Rows.Clear()
                usaDataset.Tables(0).Rows.Add({"construction"})
                ConstruDs.Tables.Clear()
                ConstruDs = Await PullUrlWs(usaDataset, "construction")

                usaDataset.Tables(0).Rows.Clear()
                usaDataset.Tables(0).Rows.Add({"relations"})
                relaTionDs.Tables.Clear()
                relaTionDs = Await PullUrlWs(usaDataset, "relations")

                elRet = "ok"

            Case Is = 5

                usaDataset.Tables(0).Rows.Add({"dependencies"})
                depeDs.Tables.Clear()
                depeDs = Await PullUrlWs(usaDataset, "dependencies")

                usaDataset.Tables(0).Rows.Clear()
                usaDataset.Tables(0).Rows.Add({"templates"})
                tempDs.Tables.Clear()
                tempDs = Await PullUrlWs(usaDataset, "templates")

                elRet = "ok"

        End Select

        Return elRet

    End Function

    Private Sub CargaHeadersDg(ByVal laOpcion As Integer)
        Select Case laOpcion
            Case Is = 1

                DataGridView1.Columns.Clear()
                'Se va a definir el ancho de las columnas dependiendo de lo que seleccione!

                'DataGridView1.Columns.Add("Key", "Key") '0
                'DataGridView1.Columns.Add("Description", "Description") '1

                'For i = 0 To DataGridView1.Columns.Count - 1
                'DataGridView1.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                'Next

                'Asi estaba antes:
                'DataGridView1.Columns.Clear()
                'DataGridView1.Columns.Add("Key", "Key") '0
                'DataGridView1.Columns.Add("Description", "Description") '1

                'For i = 0 To DataGridView1.Columns.Count - 1
                'DataGridView1.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                'Next

            Case Is = 2

                DataGridView1.ColumnHeadersDefaultCellStyle.Font = New System.Drawing.Font("Calibri", 12, FontStyle.Bold)

                DataGridView1.Columns.Clear()
                'DataGridView1.Columns.Add("DepTableCode", "Dependent Table Code") '0
                'DataGridView1.Columns.Add("DepTableName", "Dependent Table Name") '1
                DataGridView1.Columns.Add("DepFieldCode", "Dependent Field Code") '0
                DataGridView1.Columns.Add("DepFieldName", "Dependent Field Name") '1


                DataGridView1.Columns.Add("ConFieldCode", "Conditional Field Code") '2
                DataGridView1.Columns.Add("ConFieldName", "Conditional Field Name") '3
                DataGridView1.Columns.Add("ConFieldModule", "Conditional Field Module") '4
                DataGridView1.Columns.Add("ConFieldObject", "Conditional Field Object") '5
                DataGridView1.Columns.Add("ConFieldTableCode", "Conditional Field Table Code") '6
                DataGridView1.Columns.Add("ConFieldTableName", "Conditional Field Table name") '7

                DataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 64, 114, 196)

                For i = 0 To DataGridView1.Columns.Count - 1
                    DataGridView1.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                Next

                For i = 0 To 1
                    DataGridView1.Columns(i).HeaderCell.Style.BackColor = Color.FromArgb(255, 64, 114, 196)
                Next

                For i = 2 To DataGridView1.Columns.Count - 1
                    DataGridView1.Columns(i).HeaderCell.Style.BackColor = Color.FromArgb(255, 237, 125, 49)
                Next

            Case Is = 3
                'records!, aqui esta vacio porque el usuario debe seleccionar primero los templates y despues desplegar la info!!
                DataGridView1.Columns.Clear()'


            Case Is = 4 'templates!

                DataGridView1.ColumnHeadersDefaultCellStyle.Font = New System.Drawing.Font("Calibri", 12, FontStyle.Bold)
                DataGridView1.DataSource = Nothing
                DataGridView1.Rows.Clear()
                DataGridView1.Columns.Clear()

                Dim cHekKeyField As New DataGridViewCheckBoxColumn
                Dim EntryKind As New DataGridViewComboBoxColumn
                Dim FillRule As New DataGridViewComboBoxColumn
                Dim DatType As New DataGridViewComboBoxColumn
                Dim ULCase As New DataGridViewComboBoxColumn
                Dim Blancos As New DataGridViewComboBoxColumn
                Dim NoRep As New DataGridViewComboBoxColumn
                'Dim CondRule As New DataGridViewComboBoxColumn

                'CondRule.Name = "ConditionalRule"
                'CondRule.HeaderText = "Conditional Rule"

                NoRep.Name = "NonRep"
                NoRep.HeaderText = "Non Rep?"

                ULCase.Name = "ULCASE"
                ULCase.HeaderText = "U/L Case"

                Blancos.Name = "Blanks"
                Blancos.HeaderText = "Blanks"

                DatType.Name = "DataType"
                DatType.HeaderText = "Data Type"

                EntryKind.Name = "MOC"
                EntryKind.HeaderText = "MOC"

                'CondRule.Items.Clear()
                'CondRule.Items.Add("No selection")

                'Para Conditional interno:
                'CondRule.Items.Add("OR")
                'CondRule.Items.Add("NULL")
                'CondRule.Items.Add("STARTWITH")
                'CondRule.Items.Add("ENDWITH")
                'CondRule.Items.Add("CONTAINS")
                'CondRule.Items.Add("EXCEPT")
                'CondRule.Items.Add("FIXED VALUE")

                'Para condicional externo:
                'CondRule.Items.Add("MATCHS") 'existe en el otro objeto y la otra columna
                'CondRule.Items.Add("STARTWITH") 'inicia
                'CondRule.Items.Add("ENDWITH") 'termina
                'CondRule.Items.Add("CONTAINS") 'Contiene

                EntryKind.Items.Clear()
                EntryKind.Items.Add("No selection")
                EntryKind.Items.Add("Mandatory")
                EntryKind.Items.Add("Optional")
                EntryKind.Items.Add("Conditional")

                FillRule.Name = "FillingRule"
                FillRule.HeaderText = "Filling Rule"

                FillRule.Items.Clear()
                FillRule.Items.Add("No selection")
                FillRule.Items.Add("A - From Catalog")
                FillRule.Items.Add("B - Construction")
                FillRule.Items.Add("D - Fixed Value")
                FillRule.Items.Add("E - Simple validation")

                DatType.Items.Clear()
                DatType.Items.Add("No selection")
                DatType.Items.Add("Number")
                DatType.Items.Add("Date")
                DatType.Items.Add("Text")
                DatType.Items.Add("Email")
                DatType.Items.Add("Decimal")
                DatType.Items.Add("Indicator")
                DatType.Items.Add("Time")

                ULCase.Items.Clear()
                ULCase.Items.Add("None")
                ULCase.Items.Add("UCASE")
                ULCase.Items.Add("lcase")

                Blancos.Items.Clear()
                Blancos.Items.Add("None")
                Blancos.Items.Add("Left")
                Blancos.Items.Add("Right")
                Blancos.Items.Add("Both")

                NoRep.Items.Clear()
                NoRep.Items.Add("N/A")
                NoRep.Items.Add("No")
                NoRep.Items.Add("Yes")

                DatType.DefaultCellStyle.BackColor = Color.White
                DatType.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
                DatType.DisplayStyleForCurrentCellOnly = True

                FillRule.DefaultCellStyle.BackColor = Color.White
                FillRule.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
                FillRule.DisplayStyleForCurrentCellOnly = True

                EntryKind.DefaultCellStyle.BackColor = Color.White
                EntryKind.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
                EntryKind.DisplayStyleForCurrentCellOnly = True

                ULCase.DefaultCellStyle.BackColor = Color.White
                ULCase.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
                ULCase.DisplayStyleForCurrentCellOnly = True

                Blancos.DefaultCellStyle.BackColor = Color.White
                Blancos.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
                Blancos.DisplayStyleForCurrentCellOnly = True

                NoRep.DefaultCellStyle.BackColor = Color.White
                NoRep.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
                NoRep.DisplayStyleForCurrentCellOnly = True

                'CondRule.DefaultCellStyle.BackColor = Color.White
                'CondRule.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox
                'CondRule.DisplayStyleForCurrentCellOnly = True

                cHekKeyField.HeaderText = "Key Field?"
                cHekKeyField.Name = "isKey"

                'http://csharp.net-informations.com/datagridview/autogridview.htm
                'https://www.c-sharpcorner.com/UploadFile/c5c6e2/datagridview-autocomplete-textbox/

                DataGridView1.Columns.Add("FieldCode", "Field Code") '0
                DataGridView1.Columns.Add("FieldName", "Field Name") '1
                DataGridView1.Columns.Add(cHekKeyField) '2 '"isKey", "Key Field?"
                DataGridView1.Columns.Add(EntryKind) '3
                DataGridView1.Columns.Add(FillRule) '4
                DataGridView1.Columns.Add(DatType) '5
                DataGridView1.Columns.Add("MaxChar", "Max Char") '6
                DataGridView1.Columns.Add(ULCase) '7
                DataGridView1.Columns.Add(Blancos) '8
                'autocomplete edittext!
                DataGridView1.Columns.Add("CatalogName", "Cat Name") '9
                DataGridView1.Columns.Add("CatalogCode", "Cat Code") '10
                DataGridView1.Columns.Add("Value", "Value") '11
                DataGridView1.Columns.Add(NoRep) '12'ojo, columna comodin!, puede tener N/A ó un checkbox!, 
                DataGridView1.Columns.Add("NotAllowed", "Not Allowed Chars") '13

                DataGridView1.Columns.Add("ConditionalPath", "Conditional Path") '14
                DataGridView1.Columns.Add("ConditionalObject", "Conditional Object") '15
                DataGridView1.Columns.Add("ConditionalTable", "Conditional Table") '16
                DataGridView1.Columns.Add("ConditionalField", "Conditional Field") '17
                DataGridView1.Columns.Add("ConditionalType", "Conditional Type") '18
                DataGridView1.Columns.Add("ConditionalRule", "Conditional Rule") '19'conditional Rule
                DataGridView1.Columns.Add("ConditionalValue", "Conditional Value") '20
                DataGridView1.Columns.Add("MatchingFields", "Matching Fields") '21
                DataGridView1.Columns.Add("ConditionalScope", "Conditional Scope") '22
                'condition type?, Internal,external, si es interno, entonces seleccionas la regla,
                'si es externo, seleccionas el catálogo, condition Type
                'la diferencia es que el interno(misma tabla), es sobre el mismo renglon, y muchas reglas aplican
                ', si es externo, solo seria por ejemplo que exista, exists, que inicie, que termine, 
                DataGridView1.Columns.Add("ConstructionRule", "Cons Rule Ok?") '23 'N/A, Ok, Not set

                DataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 64, 114, 196)

                For i = 0 To DataGridView1.Columns.Count - 1
                    DataGridView1.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                Next

                For i = 0 To 13 'DataGridView1.Columns.Count - 1
                    DataGridView1.Columns(i).HeaderCell.Style.BackColor = Color.FromArgb(255, 64, 114, 196)
                Next

                For i = 14 To DataGridView1.Columns.Count - 1
                    DataGridView1.Columns(i).HeaderCell.Style.BackColor = Color.FromArgb(255, 255, 140, 0)
                Next



        End Select

    End Sub

    Private Async Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect

        'e.Node.Text
        'TreeView1.Nodes.IndexOfKey("")
        Dim Posi As Integer = -1
        'Dim elNom As String = e.Node.Name
        Dim elCamino As String = e.Node.FullPath
        Dim xObj As Object = Nothing
        Dim yObj As Object = Nothing
        Dim zObj As Object = Nothing
        Dim cadFind As String = ""

        'ciclar a todos los nodos hacia abajo
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim elTit As String = ""
        Dim tablaFind As String = ""
        Dim tabName As String = ""
        Dim k As Integer = 0
        Dim retKey As String = ""

        elNode = e.Node
        puedoEditar = False
        ToolStripLabel2.Text = ""

        Await DejoDeUsar()

        If RoleUsuario = "Editor" Then Module5.AgregaTiempo()

        Select Case ToolStripComboBox1.SelectedIndex
            Case Is = 0
                'Nada!
                NodoNameActual = ""

            Case Is = 1
                'marca el nodo, busca
                'Posi = TreeView1.Nodes.IndexOf(e.Node)

                xObj = Split(e.Node.FullPath, "\")
                If xObj.Length = 1 Then
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Rows.Clear()
                    NodoNameActual = ""
                    puSSyCat = -1
                    moduloSelek = e.Node.Name 'segun yo!
                    pathFinduse = "" 'nada que buscar!
                    'eliminar uso en caso de que haya algo!

                Else
                    Label1.Text = "Catalogs"
                    pathFinduse = RaizFire
                    pathFinduse = pathFinduse & "/" & "inuse"
                    pathFinduse = pathFinduse & "/" & "catalogs"

                    If xObj.Length = 2 Then
                        cadFind = e.Node.Parent.Name & "#" & e.Node.Name
                        Label2.Text = CStr(xObj(0)) & "/" & e.Node.Text
                        moduloSelek = e.Node.Parent.Name
                        pathFinduse = pathFinduse & "/" & e.Node.Parent.Name
                        pathFinduse = pathFinduse & "/" & e.Node.Name
                    Else
                        cadFind = e.Node.Parent.Parent.Name & "#" & e.Node.Parent.Name
                        Label2.Text = CStr(xObj(0)) & "/" & e.Node.Parent.Text
                        moduloSelek = e.Node.Parent.Parent.Name
                        pathFinduse = pathFinduse & "/" & e.Node.Parent.Parent.Name
                        pathFinduse = pathFinduse & "/" & e.Node.Parent.Name
                    End If

                    If NodoNameActual = "" Then
                        NodoNameActual = cadFind
                    Else
                        If NodoNameActual = cadFind Then Exit Sub
                        NodoNameActual = cadFind
                    End If

                    For i = 0 To catDs.Tables.Count - 1
                        If catDs.Tables(i).TableName = cadFind Then
                            Posi = i
                            Exit For
                        End If
                    Next


                    editDs.Clear()

                    If RoleUsuario = "Editor" Then
                        Dim enCuentra As DataRow
                        enCuentra = ModuPermit.Tables(0).Rows.Find(moduloSelek.ToUpper())
                        If IsNothing(enCuentra) = False Then

                            'significa que SI puede editar este modulo, verificar que NO haya nadie
                            editDs = Await PullDtFb(pathFinduse, "inuse")
                            puedoEditar = Await siPuedoEditar(pathFinduse)

                        End If

                    End If

                    If puedoEditar = True Then
                        'escribo ANTES de mostrar que lo estoy usando/editando!
                        're-escribo SOLO si el usuario no era YO!
                        inUseDt.Rows(0).Item(0) = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                        inUseDt.Rows(0).Item(1) = ""
                        inUseDt.Rows(0).Item(2) = UsuarioCorreo
                        inUseDt.Rows(0).Item(3) = UsuarioNombre
                        inUseDt.Rows(0).Item(4) = "X"

                        'editDs.Tables(0).Rows(0).Item(0) = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                        'editDs.Tables(0).Rows(0).Item(1) = ""
                        'editDs.Tables(0).Rows(0).Item(2) = UsuarioCorreo
                        'editDs.Tables(0).Rows(0).Item(3) = UsuarioNombre
                        'editDs.Tables(0).Rows(0).Item(4) = "X"
                        'Debo ponerlo SOLO si lo estaba editando!!
                        'retKey = Await HazPost1Set(pathFinduse, editDs.Tables(0), -1)
                        retKey = Await HazPost1Set(pathFinduse, inUseDt, -1)

                        If retKey = "fail" Then
                            'editDs.Tables(0).ExtendedProperties.Item("Key") = ""
                            inUseDt.ExtendedProperties.Item("inEdit") = False
                            inUseDt.ExtendedProperties.Item("lastPath") = ""
                            inUseDt.ExtendedProperties.Item("Key") = ""
                            ToolStripLabel2.ForeColor = Color.Crimson
                            ToolStripLabel2.Text = "Error ocurred"
                            puedoEditar = False
                        Else
                            ToolStripLabel2.ForeColor = Color.DarkGreen
                            ToolStripLabel2.Text = "Available for edit"
                            inUseDt.ExtendedProperties.Item("inEdit") = True
                            inUseDt.ExtendedProperties.Item("lastPath") = pathFinduse
                            inUseDt.ExtendedProperties.Item("Key") = retKey
                            'editDs.Tables(0).ExtendedProperties.Item("inEdit") = True
                            'editDs.Tables(0).ExtendedProperties.Add("lastPath", pathFinduse)
                            'editDs.Tables(0).ExtendedProperties.Item("Key") = retKey
                        End If

                    Else

                        'poner que NO puede editar en el label y en rojo!
                        NoPuedesEditar(editDs, ToolStripLabel2)

                    End If


                    If Posi >= 0 Then

                        Dim bS As New BindingSource
                        cataNombre = catDs.Tables(Posi).TableName
                        puSSyCat = Posi
                        DataGridView1.DataSource = Nothing
                        DataGridView1.Rows.Clear()
                        DataGridView1.Columns.Clear()

                        DataGridView1.AllowUserToAddRows = False
                        DataGridView1.AllowUserToDeleteRows = False

                        For i = 0 To DataGridView1.Columns.Count - 1
                            DataGridView1.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                        Next

                        estoyAgregandoRows = True

                        'ahora primero agregamos las columnas!
                        For j = 2 To catDs.Tables(Posi).Columns.Count - 1
                            DataGridView1.Columns.Add(catDs.Tables(Posi).Columns(j).ColumnName, catDs.Tables(Posi).Columns(j).ExtendedProperties.Item("FieldName"))
                            DataGridView1.Columns(DataGridView1.Columns.Count - 1).DataPropertyName = catDs.Tables(Posi).Columns(j).ColumnName
                        Next

                        bS.DataSource = catDs.Tables(Posi)
                        DataGridView1.AutoGenerateColumns = False
                        DataGridView1.DataSource = bS

                        For i = 0 To DataGridView1.Rows.Count - 1
                            DataGridView1.Rows(i).HeaderCell.Value = CStr(i + 1)
                        Next

                        estoyAgregandoRows = False


                        DataGridView1.RowHeadersWidth = 70

                        For i = 0 To DataGridView1.Columns.Count - 1
                            DataGridView1.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                        Next

                        'DataGridView1.AllowUserToAddRows = puedoEditar ' True
                        'DataGridView1.AllowUserToDeleteRows = puedoEditar ' True
                        'DataGridView1.ReadOnly = Not puedoEditar

                    Else
                        'es un nodo nuevo!
                        'simplemente se limpia el grid y se deja libre para que ponga la info!
                        DataGridView1.DataSource = Nothing
                        DataGridView1.Rows.Clear()

                        puSSyCat = -1
                        cataNombre = ""
                    End If

                    DataGridView1.AllowUserToAddRows = puedoEditar ' True
                    DataGridView1.AllowUserToDeleteRows = puedoEditar ' True
                    DataGridView1.ReadOnly = Not puedoEditar


                End If


            Case Is = 2

                xObj = Split(e.Node.FullPath, "\")

                If xObj.Length = 1 Or xObj.Length = 2 Then
                    'error!
                    NodoNameActual = ""
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Rows.Clear()
                Else

                    Select Case xObj.Length

                        Case Is = 3
                            tablaFind = elNode.Name
                            tabName = elNode.Tag

                            cadFind = elNode.Parent.Name & "#" & elNode.Parent.Tag
                            elTit = elNode.Parent.Name & " / " & elNode.Parent.Tag & " / " & tablaFind & " / " & tabName


                        Case Is = 4
                            tablaFind = elNode.Parent.Name
                            tabName = elNode.Parent.Tag
                            cadFind = elNode.Parent.Parent.Name & "#" & elNode.Parent.Parent.Tag
                            elTit = elNode.Parent.Parent.Name & " / " & elNode.Parent.Parent.Tag & " / " & tablaFind & " / " & tabName


                        Case Is = 5
                            tablaFind = elNode.Parent.Parent.Name
                            tabName = elNode.Parent.Parent.Tag
                            cadFind = elNode.Parent.Parent.Parent.Name & "#" & elNode.Parent.Parent.Parent.Tag
                            elTit = elNode.Parent.Parent.Parent.Name & " / " & elNode.Parent.Parent.Parent.Tag & " / " & tablaFind & " / " & tabName


                    End Select

                    If NodoNameActual = "" Then
                        NodoNameActual = tablaFind
                    Else
                        If NodoNameActual = tablaFind Then Exit Sub
                        NodoNameActual = tablaFind
                    End If

                    Label1.Text = "Dependencies"
                    Label2.Text = elTit

                    For i = 0 To depeDs.Tables.Count - 1
                        If depeDs.Tables(i).TableName = cadFind Then
                            Posi = i
                            Exit For
                        End If
                    Next

                    If Posi >= 0 Then
                        DataGridView1.DataSource = Nothing
                        DataGridView1.Rows.Clear()

                        DataGridView1.AllowUserToAddRows = False
                        DataGridView1.AllowUserToDeleteRows = False

                        For i = 0 To DataGridView1.Columns.Count - 1
                            DataGridView1.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                        Next

                        Dim filterDT As DataTable = depeDs.Tables(Posi).Clone()

                        Dim result() As DataRow = depeDs.Tables(Posi).Select("DepTableCode = '" & tablaFind & "'")

                        For Each row As DataRow In result
                            filterDT.ImportRow(row)
                        Next

                        For j = 0 To filterDT.Rows.Count - 1
                            DataGridView1.Rows.Add()
                            DataGridView1.Rows(DataGridView1.Rows.Count - 1).HeaderCell.Value = CStr(j + 1)
                            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = depeDs.Tables(Posi).Rows(j).Item(2)
                            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = depeDs.Tables(Posi).Rows(j).Item(3)
                            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = depeDs.Tables(Posi).Rows(j).Item(4)
                            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3).Value = depeDs.Tables(Posi).Rows(j).Item(5)
                            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = depeDs.Tables(Posi).Rows(j).Item(6)
                            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(5).Value = depeDs.Tables(Posi).Rows(j).Item(7)
                            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = depeDs.Tables(Posi).Rows(j).Item(8)
                            DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(7).Value = depeDs.Tables(Posi).Rows(j).Item(9)
                        Next

                        DataGridView1.RowHeadersWidth = 70

                        For i = 0 To DataGridView1.Columns.Count - 1
                            DataGridView1.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                        Next

                        DataGridView1.AllowUserToAddRows = True
                        DataGridView1.AllowUserToDeleteRows = True
                    Else
                        DataGridView1.DataSource = Nothing
                        DataGridView1.Rows.Clear()
                        DataGridView1.AllowUserToAddRows = True
                        DataGridView1.AllowUserToDeleteRows = True
                    End If
                End If




            Case Is = 3
                'records
                xObj = Split(e.Node.FullPath, "\")

                If xObj.Length <= 2 Then
                    DoneEvaluacion = False
                    NodoNameActual = ""
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Rows.Clear()
                    DataGridView1.Columns.Clear()
                    pathFinduse = ""
                    PussyTemp = -1
                    Exit Sub
                End If

                If xObj.Length = 3 Then
                    DoneEvaluacion = False
                    NodoNameActual = ""
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Rows.Clear()
                    DataGridView1.Columns.Clear()
                    pathFinduse = ""

                    For i = 0 To tempDs.Tables.Count - 1
                        yObj = Nothing
                        yObj = Split(tempDs.Tables(i).TableName, "#")
                        If yObj(0) = elNode.Name Then
                            PussyTemp = i
                            moduloSelek = CStr(yObj(2))
                            Exit For
                        End If
                    Next

                    compaSelekted = elNode.Parent.Name
                    objetoSelek = elNode.Name

                    '///////////////////////////////7
                    'OJJOOOOOOO, esta linea de puedoeditar NO VAA, se debe cambiar por el procedimiento
                    'de bloqueo completo de hojas!
                    'puedoEditar = True

                    Exit Sub

                End If

                'primero mostramos las columnas!
                tablaFind = elNode.Parent.Parent.Name & "#" & elNode.Name

                objetoSelek = elNode.Parent.Name
                tableSelek = elNode.Name

                Label1.Text = "Records"
                Label2.Text = elNode.Parent.Parent.Name & " / " & elNode.Parent.Parent.Tag & " / " & elNode.Parent.Name & " / " & elNode.Parent.Tag & " / " & elNode.Name & " / " & elNode.Tag

                compaSelekted = elNode.Parent.Parent.Name

                If NodoNameActual = "" Then
                    NodoNameActual = tablaFind
                Else
                    If NodoNameActual = tablaFind Then Exit Sub
                    NodoNameActual = tablaFind
                End If

                pathFinduse = RaizFire
                pathFinduse = pathFinduse & "/" & "inuse"
                pathFinduse = pathFinduse & "/" & "records"
                pathFinduse = pathFinduse & "/" & compaSelekted
                pathFinduse = pathFinduse & "/" & objetoSelek
                pathFinduse = pathFinduse & "/" & tableSelek

                DataGridView1.DataSource = Nothing
                DataGridView1.Rows.Clear()
                DataGridView1.Columns.Clear()

                JalaRegistrosHermanos(elNode.Parent.Parent.Name)

                DoneEvaluacion = False

                For i = 0 To tempDs.Tables.Count - 1
                    xObj = Nothing
                    xObj = Split(tempDs.Tables(i).TableName, "#")

                    If xObj(0) = elNode.Parent.Name Then

                        PussyTemp = i
                        moduloSelek = CStr(xObj(2))

                        ValidaDt.PrimaryKey = Nothing
                        ValidaDt.Columns.Clear()
                        ValidaDt.Rows.Clear()
                        ValidaDt = tempDs.Tables(i).Clone()

                        Dim filterDT As DataTable = tempDs.Tables(i).Clone()

                        Dim result() As DataRow = tempDs.Tables(i).Select("TableCode = '" & elNode.Name & "'", "Position ASC")

                        For Each row As DataRow In result
                            filterDT.ImportRow(row)
                            ValidaDt.ImportRow(row)
                        Next

                        DataGridView1.ColumnHeadersDefaultCellStyle.Font = New System.Drawing.Font("Calibri", 12, FontStyle.Bold)
                        'OJO, dar formato!
                        DataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 242, 242, 242)

                        For j = 0 To DataGridView1.Columns.Count - 1
                            DataGridView1.Columns(j).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                        Next

                        For j = 0 To filterDT.Rows.Count - 1
                            DataGridView1.Columns.Add(filterDT.Rows(j).Item(3), filterDT.Rows(j).Item(4))
                            DataGridView1.Columns(DataGridView1.Columns.Count - 1).DataPropertyName = filterDT.Rows(j).Item(3)
                            DataGridView1.Columns(DataGridView1.Columns.Count - 1).Tag = filterDT.Rows(j).Item(5) 'si es campo Yave o no!
                            'DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderCell = New DataGridViewAutoFilter.DataGridViewAutoFilterColumnHeaderCell
                            DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderText = filterDT.Rows(j).Item(4)
                            DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderCell.ToolTipText = CStr(filterDT.Rows(j).Item(8))
                            DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                            Select Case CStr(filterDT.Rows(j).Item(8))
                                Case Is = "Mandatory"
                                    DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderCell.Style.BackColor = Color.FromArgb(255, 199, 206)
                                    DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderCell.Style.ForeColor = Color.FromArgb(156, 0, 6)

                                Case Is = "Optional"
                                    DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderCell.Style.BackColor = Color.FromArgb(198, 239, 206)
                                    DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderCell.Style.ForeColor = Color.FromArgb(0, 97, 0)


                                Case Is = "Conditional"
                                    DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderCell.Style.BackColor = Color.FromArgb(255, 235, 156)
                                    DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderCell.Style.ForeColor = Color.FromArgb(156, 87, 0)

                            End Select

                        Next

                        'agregar 3 columnas extras para validar!
                        DataGridView1.Columns.Add("ROWOK", "¿Row Ok?")
                        DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderCell.Style.BackColor = Color.FromArgb(242, 242, 242)
                        DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderCell.Style.ForeColor = Color.FromArgb(63, 63, 63)
                        DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                        'DataGridView1.Columns(DataGridView1.Columns.Count - 1).DataPropertyName = "ROWOK"
                        'DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderCell = New DataGridViewAutoFilter.DataGridViewAutoFilterColumnHeaderCell
                        'DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderText = "¿Row Ok?"

                        DataGridView1.Columns.Add("ROWREP", "Repeteability?")
                        DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderCell.Style.BackColor = Color.FromArgb(242, 242, 242)
                        DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderCell.Style.ForeColor = Color.FromArgb(63, 63, 63)
                        DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                        'DataGridView1.Columns(DataGridView1.Columns.Count - 1).DataPropertyName = "ROWREP"
                        'DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderCell = New DataGridViewAutoFilter.DataGridViewAutoFilterColumnHeaderCell
                        'DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderText = "Repeteability?"

                        DataGridView1.Columns.Add("ROWCOMMS", "Internal-Relations Ok?")
                        DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderCell.Style.BackColor = Color.FromArgb(242, 242, 242)
                        DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderCell.Style.ForeColor = Color.FromArgb(63, 63, 63)
                        DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                        'DataGridView1.Columns(DataGridView1.Columns.Count - 1).DataPropertyName = "ROWCOMMS"
                        'DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderCell = New DataGridViewAutoFilter.DataGridViewAutoFilterColumnHeaderCell
                        'DataGridView1.Columns(DataGridView1.Columns.Count - 1).HeaderText = "Internal-Relations Ok?"

                        'Esto si estaba!!
                        'DataGridView1.AllowUserToAddRows = True
                        'DataGridView1.AllowUserToDeleteRows = True

                        Exit For
                    End If

                Next

                editDs.Clear()
                If RoleUsuario = "Editor" Then
                    Dim enCuentra As DataRow
                    enCuentra = ModuPermit.Tables(0).Rows.Find(moduloSelek.ToUpper())
                    If IsNothing(enCuentra) = False Then
                        'significa que SI puede editar este modulo, verificar que NO haya nadie
                        editDs = Await PullDtFb(pathFinduse, "inuse")
                        puedoEditar = Await siPuedoEditar(pathFinduse)
                    End If
                End If

                If puedoEditar = True Then
                    'escribo ANTES de mostrar que lo estoy usando/editando!
                    're-escribo SOLO si el usuario no era YO!

                    inUseDt.Rows(0).Item(0) = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                    inUseDt.Rows(0).Item(1) = ""
                    inUseDt.Rows(0).Item(2) = UsuarioCorreo
                    inUseDt.Rows(0).Item(3) = UsuarioNombre
                    inUseDt.Rows(0).Item(4) = "X"

                    'editDs.Tables(0).Rows(0).Item(0) = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                    'editDs.Tables(0).Rows(0).Item(1) = ""
                    'editDs.Tables(0).Rows(0).Item(2) = UsuarioCorreo
                    'editDs.Tables(0).Rows(0).Item(3) = UsuarioNombre
                    'editDs.Tables(0).Rows(0).Item(4) = "X"
                    'Debo ponerlo SOLO si lo estaba editando!!
                    retKey = Await HazPost1Set(pathFinduse, inUseDt, -1)

                    If retKey = "fail" Then

                        inUseDt.ExtendedProperties.Item("lastPath") = ""
                        inUseDt.ExtendedProperties.Item("Key") = ""
                        inUseDt.ExtendedProperties.Item("inEdit") = False

                        'editDs.Tables(0).ExtendedProperties.Item("Key") = ""
                        ToolStripLabel2.ForeColor = Color.Crimson
                        ToolStripLabel2.Text = "Error ocurred"
                        puedoEditar = False
                    Else
                        ToolStripLabel2.ForeColor = Color.DarkGreen
                        ToolStripLabel2.Text = "Available for edit"

                        inUseDt.ExtendedProperties.Item("lastPath") = pathFinduse
                        inUseDt.ExtendedProperties.Item("Key") = ""
                        inUseDt.ExtendedProperties.Item("inEdit") = True

                        'editDs.Tables(0).ExtendedProperties.Item("inEdit") = True
                        'editDs.Tables(0).ExtendedProperties.Add("lastPath", pathFinduse)
                        'editDs.Tables(0).ExtendedProperties.Item("Key") = retKey
                    End If

                Else
                    'poner que NO puede editar en el label y en rojo!

                    NoPuedesEditar(editDs, ToolStripLabel2)

                End If

                DataGridView1.AllowUserToAddRows = puedoEditar 'True
                DataGridView1.AllowUserToDeleteRows = puedoEditar 'True
                DataGridView1.ReadOnly = Not puedoEditar

                MuestraRecords()


            Case Is = 4
                'templates
                xObj = Split(e.Node.FullPath, "\")
                If xObj.Length = 1 Or xObj.Length = 2 Then
                    'error!
                    'o nada?!
                    tablaNombre = ""
                    NodoNameActual = ""
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Rows.Clear()
                    PussyTemp = -1
                    pathFinduse = ""

                Else

                    Select Case xObj.Length
                        Case Is = 2 'este ya no!
                            yObj = Split(CStr(elNode.Tag), "#")
                            cadFind = elNode.Name & "#" & elNode.Tag
                            elTit = elNode.Name & " / " & CStr(yObj(0)) & " / " & CStr(yObj(1))


                        Case Is = 3
                            tablaFind = elNode.Name
                            yObj = Split(CStr(elNode.Parent.Tag), "#")
                            cadFind = elNode.Parent.Name & "#" & elNode.Parent.Tag
                            elTit = elNode.Parent.Name & " / " & CStr(yObj(0)) & " / " & elNode.Text & " / " & CStr(yObj(1))
                            tabName = elNode.Parent.Name
                            'si es 3, seleccionó una tabla especifica!

                        Case Is = 4
                            tablaFind = elNode.Parent.Name
                            yObj = Split(CStr(elNode.Parent.Parent.Tag), "#")
                            cadFind = elNode.Parent.Parent.Name & "#" & elNode.Parent.Parent.Tag
                            elTit = elNode.Parent.Parent.Name & " / " & CStr(yObj(0)) & " / " & elNode.Parent.Text & " / " & CStr(yObj(1))
                            tabName = elNode.Parent.Parent.Name

                            'si es 4 seleccionó

                    End Select

                    moduloSelek = CStr(yObj(1)).ToLower()
                    objetoName = CStr(yObj(0))
                    objetoSelek = tabName
                    tableSelek = tablaFind

                    pathFinduse = RaizFire
                    pathFinduse = pathFinduse & "/" & "inuse"
                    pathFinduse = pathFinduse & "/" & "templates"
                    pathFinduse = pathFinduse & "/" & objetoSelek
                    pathFinduse = pathFinduse & "/" & tableSelek

                    If NodoNameActual = "" Then
                        NodoNameActual = tablaFind
                    Else
                        If NodoNameActual = tablaFind Then Exit Sub
                        NodoNameActual = tablaFind
                    End If

                    Label1.Text = "Templates"
                    Label2.Text = elTit

                    For i = 0 To tempDs.Tables.Count - 1
                        If tempDs.Tables(i).TableName = cadFind Then
                            PussyTemp = i
                            Posi = i
                            tablaNombre = cadFind
                            Exit For
                        End If
                    Next

                    filtDepe.PrimaryKey = Nothing
                    filtDepe.Columns.Clear()
                    filtDepe.Rows.Clear()
                    filtDepe = depeDs.Tables(0).Clone()

                    For i = 0 To depeDs.Tables.Count - 1
                        zObj = Split(depeDs.Tables(i).TableName, "#")
                        If CStr(zObj(0)) = tabName Then
                            'es el mismo objeto, filtramos sus dependencias de esa tabla!
                            Dim resultados() As DataRow = depeDs.Tables(i).Select("DepTableCode = '" & tablaFind & "'")
                            For Each row As DataRow In resultados
                                filtDepe.ImportRow(row)
                            Next
                            Exit For
                        End If
                    Next

                    'tablecode#fieldcode, esto es el campo yave de los templates
                    Dim unYve As String = ""
                    Dim enCuentra As DataRow
                    Dim z As Long
                    For i = 0 To filtDepe.Rows.Count - 1
                        unYve = filtDepe.Rows(i).Item(0) & "#" & filtDepe.Rows(i).Item(2)
                        enCuentra = tempDs.Tables(Posi).Rows.Find(CStr(unYve))
                        If IsNothing(enCuentra) = True Then Continue For
                        z = tempDs.Tables(Posi).Rows.IndexOf(enCuentra)
                        tempDs.Tables(Posi).Rows(z).Item(19) = filtDepe.Rows(i).Item(7) & ">" & filtDepe.Rows(i).Item(8) & ">" & filtDepe.Rows(i).Item(4)
                        tempDs.Tables(Posi).Rows(z).Item(20) = filtDepe.Rows(i).Item(7) 'objeto condicional
                        tempDs.Tables(Posi).Rows(z).Item(21) = filtDepe.Rows(i).Item(8) 'conditional table
                        tempDs.Tables(Posi).Rows(z).Item(22) = filtDepe.Rows(i).Item(4) 'conditional field
                        tempDs.Tables(Posi).Rows(z).Item(23) = filtDepe.Rows(i).Item(10) 'conditional type

                        tempDs.Tables(Posi).Rows(z).Item(24) = filtDepe.Rows(i).Item(11) 'Conditional Rule
                        tempDs.Tables(Posi).Rows(z).Item(25) = filtDepe.Rows(i).Item(12) 'Conditional Value
                        tempDs.Tables(Posi).Rows(z).Item(27) = filtDepe.Rows(i).Item(13) 'matching fields
                        tempDs.Tables(Posi).Rows(z).Item(28) = filtDepe.Rows(i).Item(14) 'Scope
                    Next

                    filtCat.PrimaryKey = Nothing
                    filtCat.Columns.Clear()
                    filtCat.Rows.Clear()
                    filtCat = CatSimple.Tables(0).Clone()

                    Dim result() As DataRow = CatSimple.Tables(0).Select("Module='gb' OR Module='" & CStr(yObj(1)) & "'") '")

                    For Each row As DataRow In result
                        filtCat.ImportRow(row)
                    Next

                    FuenteCatalogos.Clear()
                    For i = 0 To filtCat.Rows.Count - 1
                        FuenteCatalogos.Add(filtCat.Rows(i).Item(2) & " " & filtCat.Rows(i).Item(0))
                    Next

                    'Aqui ponemos lo de build
                    If ConstruDs.Tables.IndexOf(objetoSelek) >= 0 Then
                        'existen reglas de construcción para el template seleccionado, falta buscar para la tabla seleccionada!
                        'buscamos para
                        Dim filtBuild As DataTable
                        filtBuild = ConstruDs.Tables(objetoSelek).Clone()

                        Dim yRows() As DataRow = ConstruDs.Tables(objetoSelek).Select("TableCode='" & tableSelek & "'")

                        If yRows.Length > 0 Then

                            For Each miRow As DataRow In yRows
                                filtBuild.ImportRow(miRow)
                            Next

                            For i = 0 To tempDs.Tables(Posi).Rows.Count - 1

                                If tempDs.Tables(Posi).Rows(i).Item(1) <> tableSelek Then Continue For

                                'Ver si jala!
                                Dim uRows() As DataRow = filtBuild.Select("TableCode='" & tableSelek & "' AND FieldCode='" & tempDs.Tables(Posi).Rows(i).Item(3) & "'")
                                If uRows.Length > 0 Then
                                    Dim wiTab As DataTable = filtBuild.Select("TableCode='" & tableSelek & "' AND FieldCode='" & tempDs.Tables(Posi).Rows(i).Item(3) & "'").CopyToDataTable().DefaultView.ToTable(True, "RuleKey")
                                    tempDs.Tables(Posi).Rows(i).Item(26) = CStr(wiTab.Rows.Count) & " rule(s) found"

                                End If

                            Next

                        End If

                    End If


                    If Posi >= 0 Then

                        'aqui bloqueamos!
                        editDs.Clear()

                        If RoleUsuario = "Editor" Then

                            enCuentra = ModuPermit.Tables(0).Rows.Find(moduloSelek.ToUpper())
                            If IsNothing(enCuentra) = False Then
                                'significa que SI puede editar este modulo, verificar que NO haya nadie
                                editDs = Await PullDtFb(pathFinduse, "inuse")
                                puedoEditar = Await siPuedoEditar(pathFinduse)

                            End If

                        End If

                        If puedoEditar = True Then
                            'escribo ANTES de mostrar que lo estoy usando/editando!
                            're-escribo SOLO si el usuario no era YO!

                            inUseDt.Rows(0).Item(0) = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                            inUseDt.Rows(0).Item(1) = ""
                            inUseDt.Rows(0).Item(2) = UsuarioCorreo
                            inUseDt.Rows(0).Item(3) = UsuarioNombre
                            inUseDt.Rows(0).Item(4) = "X"


                            'editDs.Tables(0).Rows(0).Item(0) = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                            'editDs.Tables(0).Rows(0).Item(1) = ""
                            'editDs.Tables(0).Rows(0).Item(2) = UsuarioCorreo
                            'editDs.Tables(0).Rows(0).Item(3) = UsuarioNombre
                            'editDs.Tables(0).Rows(0).Item(4) = "X"

                            'Debo ponerlo SOLO si lo estaba editando!!
                            retKey = Await HazPost1Set(pathFinduse, inUseDt, -1)

                            If retKey = "fail" Then

                                inUseDt.ExtendedProperties.Item("Key") = ""
                                inUseDt.ExtendedProperties.Item("inEdit") = False
                                inUseDt.ExtendedProperties.Item("lastPath") = ""

                                'editDs.Tables(0).ExtendedProperties.Item("Key") = ""
                                ToolStripLabel2.ForeColor = Color.Crimson
                                ToolStripLabel2.Text = "Error ocurred"
                                puedoEditar = False
                            Else
                                ToolStripLabel2.ForeColor = Color.DarkGreen
                                ToolStripLabel2.Text = "Available for edit"

                                inUseDt.ExtendedProperties.Item("Key") = retKey
                                inUseDt.ExtendedProperties.Item("inEdit") = True
                                inUseDt.ExtendedProperties.Item("lastPath") = pathFinduse

                                'editDs.Tables(0).ExtendedProperties.Item("inEdit") = True
                                'editDs.Tables(0).ExtendedProperties.Add("lastPath", pathFinduse)
                                'editDs.Tables(0).ExtendedProperties.Item("Key") = retKey
                            End If


                        Else
                            'poner que NO puede editar en el label y en rojo!
                            NoPuedesEditar(editDs, ToolStripLabel2)

                        End If

                        estoyAgregandoRows = True

                        DataGridView1.DataSource = Nothing
                        DataGridView1.Rows.Clear()

                        DataGridView1.AllowUserToAddRows = False
                        DataGridView1.AllowUserToDeleteRows = False

                        For i = 0 To DataGridView1.Columns.Count - 1
                            DataGridView1.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                        Next

                        tempDs.Tables(Posi).DefaultView.Sort = "Position ASC"

                        Dim filterDT As DataTable = tempDs.Tables(Posi).DefaultView.ToTable()
                        k = 0
                        For j = 0 To filterDT.Rows.Count - 1

                            If filterDT.Rows(j).Item(1) = tablaFind Then
                                'es la tabla!
                                k = k + 1
                                DataGridView1.Rows.Add()
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).HeaderCell.Value = CStr(k)
                                'DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = tempDs.Tables(Posi).Rows(j).Item(1)
                                'DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = tempDs.Tables(Posi).Rows(j).Item(2)
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = filterDT.Rows(j).Item(3)
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = filterDT.Rows(j).Item(4)
                                If filterDT.Rows(j).Item(5) = "X" Then
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = True ' "X"
                                Else
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(2).Value = False '""
                                End If

                                'MOC
                                Dim comMoc As DataGridViewComboBoxCell
                                comMoc = DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(3)

                                Select Case filterDT.Rows(j).Item(8).ToString()
                                    Case Is = ""
                                        comMoc.Value = "No selection"

                                    Case Is = "Mandatory"
                                        comMoc.Value = "Mandatory"

                                    Case Is = "Optional"
                                        comMoc.Value = "Optional"

                                    Case Is = "Conditional"
                                        comMoc.Value = "Conditional"

                                End Select

                                If filterDT.Rows(j).Item(5) = "X" Then
                                    'es mandatorios
                                    comMoc.Value = "Mandatory"
                                    comMoc.ReadOnly = True
                                End If


                                Dim comFilRul As DataGridViewComboBoxCell
                                comFilRul = DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4)
                                Select Case filterDT.Rows(j).Item(9).ToString()
                                    Case Is = ""
                                        comFilRul.Value = "No selection"
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = ""
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).ReadOnly = True

                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).Value = ""
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).ReadOnly = True

                                    Case Is = "A - From Catalog"
                                        comFilRul.Value = "A - From Catalog"
                                        'DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = ""
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).ReadOnly = False

                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).Value = ""
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).ReadOnly = True

                                    Case Is = "B - Construction"
                                        comFilRul.Value = "B - Construction"
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = ""
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).ReadOnly = True

                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).Value = ""
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).ReadOnly = True

                                    Case Is = "D - Fixed Value"
                                        comFilRul.Value = "D - Fixed Value"
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = ""
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).ReadOnly = True

                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).ReadOnly = False

                                    Case Is = "E - Simple validation"
                                        comFilRul.Value = "E - Simple validation"
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = ""
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).ReadOnly = True

                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).Value = ""
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).ReadOnly = True

                                End Select


                                Dim comDatType As DataGridViewComboBoxCell
                                comDatType = DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(5)

                                Select Case filterDT.Rows(j).Item(10).ToString()
                                    Case Is = ""
                                        comDatType.Value = "No selection"
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = ""
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).ReadOnly = True

                                    Case Is = "Number"
                                        comDatType.Value = "Number"
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = filterDT.Rows(j).Item(11).ToString()
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).ReadOnly = False

                                    Case Is = "Date"
                                        comDatType.Value = "Date"
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = filterDT.Rows(j).Item(11).ToString()
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).ReadOnly = False

                                    Case Is = "Text"
                                        comDatType.Value = "Text"
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = filterDT.Rows(j).Item(11).ToString()
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).ReadOnly = False

                                    Case Is = "Email"
                                        comDatType.Value = "Email"
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = filterDT.Rows(j).Item(11).ToString()
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).ReadOnly = False

                                    Case Is = "Decimal"
                                        comDatType.Value = "Decimal"
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = filterDT.Rows(j).Item(11).ToString()
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).ReadOnly = False

                                    Case Is = "Indicator"
                                        comDatType.Value = "Indicator"
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = filterDT.Rows(j).Item(11).ToString()
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).ReadOnly = False

                                    Case Is = "Time"
                                        comDatType.Value = "Time"
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).Value = filterDT.Rows(j).Item(11).ToString()
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(6).ReadOnly = False

                                End Select

                                Dim ComUlCase As DataGridViewComboBoxCell
                                ComUlCase = DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(7)

                                Select Case filterDT.Rows(j).Item(12).ToString()
                                    Case Is = ""
                                        ComUlCase.Value = "None"

                                    Case Is = "UCASE"
                                        ComUlCase.Value = "UCASE"

                                    Case Is = "lcase"
                                        ComUlCase.Value = "lcase"

                                End Select


                                Dim comBla As DataGridViewComboBoxCell
                                comBla = DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(8)

                                Select Case filterDT.Rows(j).Item(13).ToString()
                                    Case Is = ""
                                        comBla.Value = "None"

                                    Case Is = "Left"
                                        comBla.Value = "Left"

                                    Case Is = "Right"
                                        comBla.Value = "Right"

                                    Case Is = "Both"
                                        comBla.Value = "Both"

                                End Select

                                'Catalog code
                                Select Case filterDT.Rows(j).Item(14).ToString()
                                    Case Is = ""
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(10).Value = "N/A"
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(10).Tag = "None" 'no tenia catálogo

                                    Case Else
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(10).Value = filterDT.Rows(j).Item(14).ToString()
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(10).Tag = filterDT.Rows(j).Item(29).ToString() & ":" & filterDT.Rows(j).Item(14).ToString() & ":" & filterDT.Rows(j).Item(30).ToString() & ":" & filterDT.Rows(j).Item(31).ToString()
                                        'catalog route>gb:gb0001:FIELDMATCH:MATNR#A-BUKRS#B


                                End Select

                                'CatalogName
                                Select Case filterDT.Rows(j).Item(15).ToString()
                                    Case Is = ""
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = "N/A"


                                    Case Else
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).Value = filterDT.Rows(j).Item(15).ToString()
                                End Select

                                'Value Column
                                Select Case filterDT.Rows(j).Item(16).ToString()
                                    Case Is = ""
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).Value = "N/A"

                                    Case Else
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(11).Value = filterDT.Rows(j).Item(16).ToString()
                                End Select


                                Dim comNoRep As DataGridViewComboBoxCell
                                comNoRep = DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(12)
                                Select Case filterDT.Rows(j).Item(17).ToString()
                                    Case Is = ""
                                        comNoRep.Value = "N/A"

                                    Case Is = "No"
                                        comNoRep.Value = "No"

                                    Case Is = "Yes"
                                        comNoRep.Value = "Yes"
                                End Select

                                'comMoc
                                If filterDT.Rows(j).Item(5) = "X" Then
                                    comNoRep.Value = "N/A"
                                    comNoRep.ReadOnly = True
                                Else
                                    comNoRep.ReadOnly = False
                                End If


                                'Not allowed Chars
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(13).Value = filterDT.Rows(j).Item(18).ToString()

                                'Conditional path
                                Select Case filterDT.Rows(j).Item(19).ToString()
                                    Case Is = ""
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(14).Value = "N/A"

                                    Case Else
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(14).Value = filterDT.Rows(j).Item(19).ToString()
                                End Select

                                'conditional Object
                                Select Case filterDT.Rows(j).Item(20).ToString()
                                    Case Is = ""
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(15).Value = "N/A"

                                    Case Else
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(15).Value = filterDT.Rows(j).Item(20).ToString()
                                End Select

                                'Conditional Table
                                Select Case filterDT.Rows(j).Item(21).ToString()
                                    Case Is = ""
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(16).Value = "N/A"

                                    Case Else
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(16).Value = filterDT.Rows(j).Item(21).ToString()
                                End Select

                                'Conditional Field
                                Select Case filterDT.Rows(j).Item(22).ToString()
                                    Case Is = ""
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(17).Value = "N/A"

                                    Case Else
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(17).Value = filterDT.Rows(j).Item(22).ToString()
                                End Select

                                'Conditional Type & Conditional Rule:
                                'Dim ComCondit As DataGridViewComboBoxCell
                                'ComCondit = DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(19)
                                Select Case filterDT.Rows(j).Item(23).ToString()
                                    Case Is = ""
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(18).Value = "N/A"
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(19).Value = "N/A"
                                        'ComCondit.Value = "No selection"

                                    Case Else
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(18).Value = filterDT.Rows(j).Item(23).ToString()
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(19).Value = filterDT.Rows(j).Item(24).ToString()

                                End Select

                                'ConditionalValue
                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(20).Value = filterDT.Rows(j).Item(25).ToString()

                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(21).Value = filterDT.Rows(j).Item(27).ToString() 'matching fields

                                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(22).Value = filterDT.Rows(j).Item(28).ToString() 'scope

                                Select Case filterDT.Rows(j).Item(26).ToString()
                                    Case Is = ""
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(23).Value = "N/A"

                                    Case Else
                                        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(23).Value = filterDT.Rows(j).Item(26).ToString()
                                End Select

                                'poner los readonly
                                'Para el catálogo!
                                If DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(4).Value = "A - From Catalog" Then
                                    'se habilita, catalog code y catalog name
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).ReadOnly = False
                                Else
                                    'se deshabilita!
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(9).ReadOnly = True
                                End If

                                If DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(18).Value = "N/A" Then
                                    'se bloquea
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(19).ReadOnly = True 'se bloquea!
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(20).ReadOnly = True 'se bloquea!
                                Else
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(19).ReadOnly = False 'se desbloquea!
                                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(20).ReadOnly = False 'se desbloquea!
                                End If

                            End If

                        Next

                        DataGridView1.Columns(0).ReadOnly = True ' siempre será de solo lectura! Not puedoEditar
                        DataGridView1.Columns(1).ReadOnly = Not puedoEditar
                        DataGridView1.Columns(2).ReadOnly = Not puedoEditar
                        DataGridView1.Columns(3).ReadOnly = Not puedoEditar
                        DataGridView1.Columns(4).ReadOnly = Not puedoEditar
                        DataGridView1.Columns(5).ReadOnly = Not puedoEditar
                        DataGridView1.Columns(6).ReadOnly = Not puedoEditar
                        DataGridView1.Columns(7).ReadOnly = Not puedoEditar
                        DataGridView1.Columns(8).ReadOnly = Not puedoEditar

                        DataGridView1.Columns(9).ReadOnly = True 'CatalogCode
                        DataGridView1.Columns(10).ReadOnly = True 'CatalogCode

                        DataGridView1.Columns(11).ReadOnly = Not puedoEditar
                        DataGridView1.Columns(12).ReadOnly = Not puedoEditar
                        DataGridView1.Columns(13).ReadOnly = Not puedoEditar

                        DataGridView1.Columns(14).ReadOnly = True 'ConditionalPath
                        DataGridView1.Columns(15).ReadOnly = True 'ConditionalObject
                        DataGridView1.Columns(16).ReadOnly = True 'ConditionalTable
                        DataGridView1.Columns(17).ReadOnly = True 'ConditionalField
                        DataGridView1.Columns(18).ReadOnly = True 'ConditionalType
                        DataGridView1.Columns(19).ReadOnly = True 'Conditional rule
                        DataGridView1.Columns(20).ReadOnly = True 'Conditional Value
                        DataGridView1.Columns(21).ReadOnly = True 'ConstructionRule
                        DataGridView1.Columns(22).ReadOnly = True 'ConstructionRule
                        DataGridView1.Columns(23).ReadOnly = True 'ConstructionRule
                        DataGridView1.RowHeadersWidth = 70

                        estoyAgregandoRows = False

                        For i = 0 To DataGridView1.Columns.Count - 1
                            DataGridView1.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                            DataGridView1.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                        Next

                        DataGridView1.AllowUserToAddRows = False ' puedoEditar ' ahora siempre en false!
                        'DataGridView1.AllowUserToDeleteRows = True

                        DataGridView1.Columns(1).Frozen = True

                    Else
                        DataGridView1.DataSource = Nothing
                        DataGridView1.Rows.Clear()
                        DataGridView1.AllowUserToAddRows = False ' True
                        'DataGridView1.AllowUserToDeleteRows = True
                    End If


                End If


        End Select

    End Sub

    Private Sub TreeView1_Click(sender As Object, e As EventArgs) Handles TreeView1.Click

    End Sub

    Private Sub TreeView1_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick

    End Sub

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        ToolStripComboBox1.SelectedIndex = 0
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub ToolStripComboBox1_Click(sender As Object, e As EventArgs) Handles ToolStripComboBox1.Click

    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click

    End Sub

    Private Async Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click

        'Add Node

        If CategSelected = 0 Then
            MsgBox("Please select a category first!", vbCritical, TitBox)
            Exit Sub
        End If

        If CategSelected = 3 Then
            MsgBox("Sorry you can't add nodes to this category", vbCritical, TitBox)
            Exit Sub
        End If

        Dim i As Integer = 0
        Dim Posi As Integer = -1
        Dim xObj As Object
        Dim yObj As Object
        Dim zObj As Object

        Dim cadFind As String = ""
        Dim pathBuild As String = ""
        Dim conseCnumba As Integer = 0
        Dim yaEntro As Boolean
        Dim valMovil As String = ""
        Dim consStr As String = ""
        Dim elCero As String = ""
        Dim nodyav As String = ""
        Dim tabNombre As String = ""
        Dim tabCode As String = ""
        Dim miCamino As String = ""
        Dim unRetorno As String = ""
        Dim moduCode As String = ""

        Dim xDs As New DataSet
        xDs.Tables.Add()
        xDs.Tables(0).Columns.Add()
        xDs.Tables(0).Columns.Add()

        writeDs.Tables(0).Rows.Clear()

        Select Case CategSelected
            Case Is = 1 'catalogos
                'Solo aqui se puede agregar, sobre el nodo seleccionado!
                writeDs.Tables(0).Rows.Add({"catalogs"})

                pathBuild = "catalogs"

                If IsNothing(elNode) = False Then

                    xObj = Split(elNode.FullPath, "\")

                    Select Case xObj.Length
                        Case Is = 1
                            moduCode = elNode.Name

                        Case Is = 2
                            moduCode = elNode.Parent.Name

                    End Select


                    Dim enCuentra As DataRow
                    enCuentra = ModuPermit.Tables(0).Rows.Find(moduCode.ToUpper())
                    If IsNothing(enCuentra) = True Then
                        MsgBox("Sorry you are not allowed to add catalogs on the selected module", vbCritical, TitBox)
                        Exit Sub
                    End If

                    writeDs.Tables(0).Rows.Add({CStr(xObj(0))}) 'siempre!'solo ahiii
                    pathBuild = pathBuild & "/" & CStr(xObj(0))
                    'se muestra el form para que introduzca el nombre de la tabla!
                    'se le pide: CatalogName y solito se pone el consecutivo de xobj(0)+ xobj(0) &0001...
                    'cuando le pique y este vacío, que pueda pegar la info!
                    'al cambiar el nodo, se pone el grid vacío, y se pone para que pueda pegar todo!

                    'se busca la última tabla de ese catalogo!
                    i = 0
                    yaEntro = False
                    Do While i < catDs.Tables.Count
                        yObj = Nothing
                        yObj = Split(catDs.Tables(i).TableName, "#")

                        If yaEntro = True Then
                            If moduCode = CStr(yObj(0)) Then
                                valMovil = CStr(yObj(1)) 'ultimo valor
                            Else
                                Exit Do
                            End If
                        Else
                            If moduCode = CStr(yObj(0)) Then
                                valMovil = CStr(yObj(1)) 'ultimo valor
                                yaEntro = True
                            End If
                        End If

                        i = i + 1
                    Loop

                    If valMovil = "" Then
                        conseCnumba = 1
                        consStr = CStr(conseCnumba)
                    Else
                        conseCnumba = CInt(Microsoft.VisualBasic.Right(valMovil, 4))
                        conseCnumba = conseCnumba + 1
                        consStr = CStr(conseCnumba)
                    End If

                    elCero = ""
                    For i = 1 To 4 - Len(consStr)
                        elCero = elCero & "0"
                    Next

                    nodyav = moduCode & elCero & consStr

                    Form2.huboExito = False
                    Form2.queOpcion = 1
                    Form2.elTitulo = "Add catalog"
                    Form2.keyName = "Table code:"
                    Form2.tabName = "Catalog name:"
                    Form2.keyValue = nodyav 'codigo de la tabla!!
                    Form2.tabValue = ""
                    Form2.pathLabel = "Path:"
                    Form2.elCamino = pathBuild
                    Form2.ShowDialog()

                    If Form2.huboExito = False Then
                        writeDs.Tables(0).Rows.Clear()
                        Exit Sub
                    End If

                    'lo agregamos!
                    'se agrega a catds!
                    'Posi = TreeView1.Nodes.IndexOf(elNode)
                    miCamino = RaizFire
                    'miCamino = miCamino & "/catalogs/" & CStr(xObj(0))
                    miCamino = miCamino & "/catpro/" & moduCode
                    miCamino = miCamino & "/" & nodyav

                    unRetorno = ""
                    unRetorno = Await HazPutEnFbSimple(miCamino, "CatalogName", Form2.tabValue)

                    If unRetorno = "Ok" Then
                        Select Case xObj.Length
                            Case Is = 1
                                elNode.Nodes.Add(nodyav, Form2.tabValue, 1, 1)

                            Case Is = 2
                                elNode.Parent.Nodes.Add(nodyav, Form2.tabValue, 1, 1)

                            Case Is = 3
                                elNode.Parent.Parent.Nodes.Add(nodyav, Form2.tabValue, 1, 1)

                        End Select

                        Dim iAveprim(1) As DataColumn
                        Dim kEys As New DataColumn()

                        kEys.ColumnName = "KeyField"
                        iAveprim(0) = kEys
                        'Aqui ahora inmediatamente debe definir la cantidad de columnas a agregar!
                        catDs.Tables.Add(moduCode & "#" & nodyav) 'SOLO se agrega, pero esta limpia!
                        'falto asignar la llave primariaa!!
                        catDs.Tables(catDs.Tables.Count - 1).Columns.Add(kEys)
                        'catDs.Tables(catDs.Tables.Count - 1).Columns.Add("Value", GetType(String))'este si estaba, en la versión anterior!
                        catDs.Tables(catDs.Tables.Count - 1).PrimaryKey = iAveprim
                        catDs.Tables(catDs.Tables.Count - 1).ExtendedProperties.Add("ModuleCode", moduCode)
                        catDs.Tables(catDs.Tables.Count - 1).ExtendedProperties.Add("CatalogCode", nodyav)
                        catDs.Tables(catDs.Tables.Count - 1).ExtendedProperties.Add("CatalogName", Form2.tabValue)
                        catDs.Tables(catDs.Tables.Count - 1).ExtendedProperties.Add("ColumnCount", 0)
                        catDs.Tables(catDs.Tables.Count - 1).ExtendedProperties.Add("ModuleName", "Undefined")

                    Else
                        MsgBox(unRetorno, vbCritical, TitBox)

                    End If

                    'AQUII falta agregarlo a FB!!


                Else
                    'error!!
                    MsgBox("Please select a node to add the catalog on the corresponding structure!", vbCritical, TitBox)

                End If

            Case Is = 2 'dependencias
                'aqui tmb se puede
                'OJO, primero se deben cargar los templates en caso de que no se haya hecho antes!
                'If tempDs.Tables.Count = 0 Then
                '    'las cargamos
                '    usaDataset.Tables(0).Rows.Clear()
                '    usaDataset.Tables(0).Rows.Add({"templates"})
                '    tempDs.Tables.Clear()
                '    tempDs = Await PullUrlWs(usaDataset, "templates")
                'End If

                'writeDs.Tables(0).Rows.Add({"dependencies"})

                ''nomas ahii!!!
                'xObj = Split(elNode.FullPath, "\")

                'Select Case xObj.Length
                '    Case Is = 1
                '        'esta hasta arriba, agrega un template!
                '        'add template
                '        Form2.xtraDs = tempDs
                '        Form2.yTraDs = depeDs
                '        Form2.queOpcion = 2
                '        Form2.huboExito = False

                '        Form2.ShowDialog()

                '        If Form2.huboExito = False Then Exit Sub

                '        pathBuild = RaizFire & "/" & "dependencies"
                '        pathBuild = pathBuild & "/" & Form2.pathValue

                '        xDs.Tables(0).Rows.Add({"ObjectName", Form2.keyValue})

                '        Await HazPutEnFireBase(pathBuild, xDs)

                '        'hago rebuild!
                '        Await CargaOpcion(CategSelected)

                '        elNode.Nodes.Add(Form2.pathValue, Form2.pathValue & " / " & Form2.keyValue, 2, 2)
                '        elNode.Nodes(elNode.Nodes.Count - 1).Tag = Form2.keyValue
                '        'se agrega el nodo a FB


                '    Case Is = 2
                '        'esta en un template, agrega una tabla/hoja!!
                '        'yObj = Split(elNode.Tag, "#")
                '        zObj = CStr("")
                '        For i = 0 To tempDs.Tables.Count - 1
                '            yObj = Split(tempDs.Tables(i).TableName, "#")
                '            If CStr(yObj(0)) = elNode.Name Then
                '                zObj = CStr(yObj(2))
                '                Exit For
                '            End If
                '        Next

                '        Dim enCuentra As DataRow
                '        enCuentra = ModuPermit.Tables(0).Rows.Find(CStr(zObj).ToUpper())
                '        If IsNothing(enCuentra) = True Then
                '            MsgBox("Sorry you are not allowed to make changes on the selected module", vbCritical, "SAP MD")
                '            Exit Sub
                '        End If


                '        Form2.xtraDs = tempDs
                '        Form2.yTraDs = depeDs
                '        Form2.queOpcion = 3
                '        Form2.huboExito = False
                '        Form2.pathValue = elNode.Name

                '        Form2.ShowDialog()

                '        If Form2.huboExito = False Then Exit Sub

                '        pathBuild = RaizFire & "/" & "dependencies"
                '        pathBuild = pathBuild & "/" & Form2.pathValue & "/" & Form2.keyValue

                '        xDs.Tables(0).Rows.Add({"TableName", Form2.tabValue})

                '        Await HazPutEnFireBase(pathBuild, xDs)

                '        'hago rebuild!
                '        Await CargaOpcion(CategSelected)

                '        elNode.Nodes.Add(Form2.keyValue, Form2.keyValue & " / " & Form2.tabValue, 8, 8)
                '        elNode.Nodes(elNode.Nodes.Count - 1).Tag = Form2.tabValue


                '    Case Is = 3
                '        'esta en una tabla, agrega un campo!
                '        'aqui ya viene la interfaz!

                '        zObj = CStr("")
                '        For i = 0 To tempDs.Tables.Count - 1
                '            yObj = Split(tempDs.Tables(i).TableName, "#")
                '            If CStr(yObj(0)) = elNode.Parent.Name Then
                '                zObj = CStr(yObj(2))
                '                Exit For
                '            End If
                '        Next

                '        Dim enCuentra As DataRow
                '        enCuentra = ModuPermit.Tables(0).Rows.Find(CStr(zObj).ToUpper())
                '        If IsNothing(enCuentra) = True Then
                '            MsgBox("Sorry you are not allowed to make changes on the selected module", vbCritical, "SAP MD")
                '            Exit Sub
                '        End If


                '        Form3.xtraDs = tempDs
                '        Form3.yTraDs = depeDs
                '        Form3.depeTemplate = elNode.Parent.Name
                '        Form3.depeTabla = elNode.Name

                '        Form3.resDepFieldCode = ""
                '        Form3.resDepFieldName = ""
                '        Form3.resConTempCode = ""
                '        Form3.resConTempName = ""
                '        Form3.resConTempModule = ""
                '        Form3.resConTableCode = ""
                '        Form3.resConTableName = ""
                '        Form3.resConFieldCode = ""
                '        Form3.resConFieldName = ""

                '        Form3.resConType = ""
                '        Form3.resConRule = ""
                '        Form3.resConVal = ""
                '        Form3.resMachFields = ""

                '        Form3.depeCampo = ""

                '        Form3.elPath = "dependencies/" & elNode.Parent.Name & "/" & elNode.Name

                '        Form3.huboExito = False

                '        Form3.ShowDialog()

                '        If Form3.huboExito = False Then Exit Sub

                '        'se escribe en firebase!
                '        pathBuild = RaizFire & "/" & "dependencies"
                '        pathBuild = pathBuild & "/" & Form3.depeTemplate
                '        pathBuild = pathBuild & "/" & Form3.depeTabla
                '        pathBuild = pathBuild & "/" & Form3.resDepFieldCode

                '        xDs.Tables(0).Rows.Add({"MyName", Form3.resDepFieldName}) '0
                '        xDs.Tables(0).Rows.Add({"Object", Form3.resConTempCode}) '1
                '        xDs.Tables(0).Rows.Add({"Module", Form3.resConTempModule}) '2
                '        xDs.Tables(0).Rows.Add({"TableCode", Form3.resConTableCode}) '3
                '        xDs.Tables(0).Rows.Add({"TableName", Form3.resConTableName}) '4
                '        xDs.Tables(0).Rows.Add({"FieldCode", Form3.resConFieldCode}) '5
                '        xDs.Tables(0).Rows.Add({"FieldName", Form3.resConFieldName}) '6

                '        xDs.Tables(0).Rows.Add({"ConditionalType", Form3.resConType}) '7
                '        xDs.Tables(0).Rows.Add({"ConditionalRule", Form3.resConRule}) '8
                '        xDs.Tables(0).Rows.Add({"ConditionalValue", Form3.resConVal}) '9

                '        Await HazPutEnFireBase(pathBuild, xDs)

                '        'hago rebuild!
                '        Await CargaOpcion(CategSelected)

                '        elNode.Nodes.Add(Form3.resDepFieldCode, Form3.resDepFieldCode & " / " & Form3.resDepFieldName, 4, 4)
                '        Posi = elNode.Nodes.Count - 1

                '        elNode.Nodes(Posi).Nodes.Add(Form3.resConTempCode, "Object: " & Form3.resConTempCode, 4, 4)
                '        elNode.Nodes(Posi).Nodes.Add(Form3.resConTempModule, "Module: " & Form3.resConTempModule, 4, 4)
                '        elNode.Nodes(Posi).Nodes.Add(Form3.resConTableCode, "Table: " & Form3.resConTableCode & " / " & Form3.resConTableName, 4, 4)
                '        elNode.Nodes(Posi).Nodes.Add(Form3.resConFieldCode, "Field: " & Form3.resConFieldCode & " / " & Form3.resConFieldName, 4, 4)


                'End Select



            Case Is = 3'records, este NO


            Case Is = 4 'templates, este sii!

                If IsNothing(elNode) = True Then
                    'error!!
                    MsgBox("Please select a node to add the template on the corresponding structure!", vbCritical, TitBox)
                    Exit Sub
                End If

                writeDs.Tables(0).Rows.Add({"templates"})


                xObj = Split(elNode.FullPath, "\")

                Select Case xObj.Length
                    Case Is = 1
                        'agrego un nuevo template
                        'writeDs.Tables(0).Rows.Add({CStr(xObj(0))})
                        'pathBuild = pathBuild & "/" & CStr(xObj(0))
                        'lo agrego!!
                        pathBuild = RaizFire & "/" & "templates"

                        Form2.huboExito = False
                        Form2.queOpcion = 5
                        Form2.elTitulo = "Add template"
                        Form2.keyName = "Template code:"
                        Form2.tabName = "Template name:"
                        Form2.keyValue = "" 'codigo de la tabla!!
                        Form2.tabValue = ""
                        Form2.pathLabel = "Module:"
                        Form2.pathValue = ""
                        Form2.elCamino = ""
                        Form2.ModuDs = ModuloDs
                        Form2.ShowDialog()

                        If Form2.huboExito = False Then
                            writeDs.Tables(0).Rows.Clear()
                            Exit Sub
                        End If

                        Posi = TreeView1.Nodes(0).Nodes.IndexOfKey(Form2.keyValue)
                        If Posi >= 0 Then
                            MsgBox("That template code already exists!!, please choose another name!!", vbCritical, TitBox)
                            Exit Sub
                        End If



                        Form2.keyValue = Form2.keyValue.ToLower()

                        pathBuild = pathBuild & "/" & Form2.keyValue
                        xDs.Tables(0).Rows.Add({"Module", Form2.pathValue})
                        xDs.Tables(0).Rows.Add({"ObjectName", Form2.tabValue})

                        Await HazPutEnFireBase(pathBuild, xDs)

                        'hago rebuild!
                        Await CargaOpcion(CategSelected)

                        TreeView1.Nodes(0).Nodes.Add(Form2.keyValue, Form2.keyValue & " / " & Form2.tabValue, 2, 2)
                        TreeView1.Nodes(0).Nodes(TreeView1.Nodes(0).Nodes.Count - 1).Tag = Form2.tabValue & "#" & Form2.pathValue

                        TreeView1.Nodes(0).Nodes(TreeView1.Nodes(0).Nodes.Count - 1).EnsureVisible()
                        TreeView1.SelectedNode = TreeView1.Nodes(0).Nodes(TreeView1.Nodes(0).Nodes.Count - 1)


                    Case Is = 2
                        'agrego una nueva tabla!
                        xObj = Split(elNode.Tag, "#")
                        Dim enCuentra As DataRow
                        enCuentra = ModuPermit.Tables(0).Rows.Find(CStr(xObj(1)).ToUpper())
                        If IsNothing(enCuentra) = True Then
                            MsgBox("Sorry you are not allowed to make changes on this template", vbCritical, TitBox)
                            Exit Sub
                        End If

                        cadFind = elNode.Name & "#" & elNode.Tag
                        writeDs.Tables(0).Rows.Add({elNode.Name})
                        pathBuild = "templates"
                        pathBuild = pathBuild & "/" & elNode.Name

                        tabCode = elNode.Name
                        tabNombre = elNode.Tag

                        Posi = tempDs.Tables.IndexOf(cadFind)
                        If Posi < 0 Then
                            'quiza es un nuevo nodo!
                            tempDs.Tables.Add(cadFind)
                            Posi = tempDs.Tables.Count - 1
                        End If

                        If tempDs.Tables(Posi).Rows.Count = 0 Then
                            nodyav = tabCode & "-0001" 'primer valor!
                        Else
                            yObj = Split(tempDs.Tables(Posi).Rows(tempDs.Tables(Posi).Rows.Count - 1).Item(1), "-")
                            valMovil = yObj(1)
                            conseCnumba = CInt(Microsoft.VisualBasic.Right(valMovil, 4))
                            conseCnumba = conseCnumba + 1
                            consStr = CStr(conseCnumba)

                            elCero = ""
                            For i = 1 To 4 - Len(consStr)
                                elCero = elCero & "0"
                            Next

                            nodyav = tabCode & "-" & elCero & consStr
                        End If

                        Form2.huboExito = False
                        Form2.queOpcion = 4
                        Form2.elTitulo = "Add table"
                        Form2.keyName = "Table code:"
                        Form2.tabName = "Table name:"
                        Form2.pathLabel = "Path:"
                        Form2.keyValue = nodyav 'codigo de la tabla!!
                        Form2.tabValue = ""

                        Form2.elCamino = pathBuild
                        Form2.ShowDialog()

                        If Form2.huboExito = False Then
                            writeDs.Tables(0).Rows.Clear()
                            Exit Sub
                        End If

                        pathBuild = RaizFire & "/templates/" & tabCode & "/" & nodyav
                        xDs.Tables(0).Rows.Add({"TableName", Form2.tabValue})

                        Await HazPutEnFireBase(pathBuild, xDs)

                        elNode.Nodes.Add(nodyav, nodyav & " / " & Form2.tabValue, 8, 8)
                        elNode.Nodes(elNode.Nodes.Count - 1).Tag = Form2.tabValue
                        elNode.Nodes(elNode.Nodes.Count - 1).EnsureVisible()
                        TreeView1.SelectedNode = elNode.Nodes(elNode.Nodes.Count - 1)
                        'Y si mejor agregamos de una vez un campo Dummy?
                        'elNode = TreeView1.SelectedNode


                    Case Else
                        Exit Sub

                End Select





        End Select




    End Sub

    Private Sub TreeView1_MouseDown(sender As Object, e As MouseEventArgs) Handles TreeView1.MouseDown
        'Dim selectedTreeview As TreeView = CType(sender, TreeView)
        'Dim pt As Point = CType(sender, TreeView).PointToClient(New Point(e.X, e.Y))
        Dim targetNode As TreeNode = TreeView1.HitTest(e.X, e.Y).Node '  selectedTreeview.GetNodeAt(pt)

        If IsNothing(targetNode) = True Then elNode = Nothing Else elNode = targetNode

    End Sub

    Private Sub TreeView1_MouseClick(sender As Object, e As MouseEventArgs) Handles TreeView1.MouseClick

        'Dim selectedTreeview As TreeView = CType(sender, TreeView)
        'Dim pt As Point = CType(sender, TreeView).PointToClient(New Point(e.X, e.Y))
        Dim targetNode As TreeNode = TreeView1.HitTest(e.X, e.Y).Node '  selectedTreeview.GetNodeAt(pt)

        If IsNothing(targetNode) = True Then elNode = Nothing Else elNode = targetNode

    End Sub

    Public Shared Sub SetDoubleBuffered(ByVal control As Control)
        GetType(Control).InvokeMember("DoubleBuffered", System.Reflection.BindingFlags.SetProperty Or System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.NonPublic, Nothing, control, New Object() {True})
    End Sub

    Private Sub ReloadOneNode(ByVal laOpcion As Integer)

        Dim xObj As Object
        Dim cadFind As String = ""
        Dim Posi As Integer = -1
        Dim unTree As New TreeNode
        Dim pos1 As Integer = 0
        Dim tabCode As String = ""

        Select Case laOpcion
            Case Is = 1
                'catalogos
                Exit Sub 'ya no hay nada que actualizar!

                xObj = Split(elNode.FullPath, "\")
                If xObj.Length = 2 Then
                    cadFind = elNode.Parent.Name & "#" & elNode.Name
                    unTree = elNode
                Else
                    cadFind = elNode.Parent.Parent.Name & "#" & elNode.Parent.Name
                    unTree = elNode.Parent
                End If
                For i = 0 To catDs.Tables.Count - 1
                    If catDs.Tables(i).TableName = cadFind Then
                        Posi = i
                        Exit For
                    End If
                Next

                If Posi < 0 Then
                    MsgBox("Unable reload node!", vbCritical, TitBox)
                    Exit Sub
                End If

                For i = 0 To unTree.Nodes.Count - 1
                    MySource.Remove(unTree.Nodes(i).Name)
                    MySource.Remove(unTree.Nodes(i).Text)
                Next

                TreeView1.BeginUpdate()
                unTree.Nodes.Clear()
                For i = 0 To catDs.Tables(Posi).Rows.Count - 1
                    unTree.Nodes.Add(CStr(catDs.Tables(Posi).Rows(i).Item(0)), CStr(catDs.Tables(Posi).Rows(i).Item(0)) & " / " & CStr(catDs.Tables(Posi).Rows(i).Item(1)), 7, 7)
                    MySource.Add(CStr(catDs.Tables(Posi).Rows(i).Item(0)))
                    MySource.Add(CStr(catDs.Tables(Posi).Rows(i).Item(1)))
                Next
                TreeView1.EndUpdate()

                elNode = unTree

                ToolStripTextBox1.AutoCompleteCustomSource = MySource
                ToolStripTextBox1.AutoCompleteMode = AutoCompleteMode.Suggest
                ToolStripTextBox1.AutoCompleteSource = AutoCompleteSource.CustomSource

            Case Is = 2
                'dependencias



            Case Is = 3
                'records



            Case Is = 4
                'templates
                xObj = Split(elNode.FullPath, "\")
                Select Case xObj.Length
                    Case Is = 2
                        'este ya no va a entrar!
                        Exit Sub
                        cadFind = elNode.Name & "#" & elNode.Tag
                        unTree = elNode

                    Case Is = 3
                        cadFind = elNode.Parent.Name & "#" & elNode.Parent.Tag
                        unTree = elNode ' elNode.Parent
                        tabCode = elNode.Name

                    Case Is = 4
                        cadFind = elNode.Parent.Parent.Name & "#" & elNode.Parent.Parent.Tag
                        unTree = elNode.Parent ' elNode.Parent.Parent
                        tabCode = elNode.Parent.Name

                End Select


                For i = 0 To tempDs.Tables.Count - 1
                    If tempDs.Tables(i).TableName = cadFind Then
                        Posi = i
                        Exit For
                    End If
                Next

                If Posi < 0 Then
                    MsgBox("Unable reload node!", vbCritical, TitBox)
                    Exit Sub
                End If

                For i = 0 To unTree.Nodes.Count - 1
                    MySource.Remove(unTree.Nodes(i).Name)
                    MySource.Remove(unTree.Nodes(i).Text)
                Next

                TreeView1.BeginUpdate()
                unTree.Nodes.Clear()

                'filtramos la tabla!
                Dim filterDT As DataTable = tempDs.Tables(Posi).Clone()

                Dim result() As DataRow = tempDs.Tables(Posi).Select("TableCode = '" & tabCode & "'", "Position ASC")

                For Each row As DataRow In result
                    filterDT.ImportRow(row)
                Next


                For j = 0 To filterDT.Rows.Count - 1

                    'pos1 = unTree.Nodes.IndexOfKey(CStr(filterDT.Rows(j).Item(1)))
                    'If pos1 < 0 Then
                    '    unTree.Nodes.Add(CStr(filterDT.Rows(j).Item(1)), CStr(filterDT.Rows(j).Item(1)) & " / " & CStr(filterDT.Rows(j).Item(2)), 8, 8)
                    '    pos1 = TreeView1.Nodes(0).Nodes(Posi).Nodes.Count - 1
                    '    MySource.Add(CStr(filterDT.Rows(j).Item(1)))
                    '    MySource.Add(CStr(filterDT.Rows(j).Item(2)))
                    'End If

                    unTree.Nodes.Add(CStr(filterDT.Rows(j).Item(3)), CStr(filterDT.Rows(j).Item(3)) & " / " & CStr(filterDT.Rows(j).Item(4)), 8, 8)
                    MySource.Add(CStr(filterDT.Rows(j).Item(3)))
                    MySource.Add(CStr(filterDT.Rows(j).Item(4)))

                Next

                elNode = unTree

                TreeView1.EndUpdate()


        End Select



    End Sub

    Private Async Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        'eliminar el nodo!
        If CategSelected = 0 Then
            MsgBox("Please select a category first!", vbCritical, TitBox)
            Exit Sub
        End If

        Dim xObj As Object
        Dim nodYav As String = ""
        Dim nodText As String = ""
        Dim x As Integer = 0
        Dim pathDel As String = ""
        Dim elRet As String
        Dim cadFind As String = ""
        Dim tablaFind As String = ""
        Dim tabName As String = ""
        Dim campoCode As String = ""
        Dim campoNombre As String = ""
        Dim objCode As String = ""
        Dim objName As String = ""
        Dim cadDep As String = ""
        Dim cadCon As String = ""
        Dim moduSel As Object = Nothing
        Dim Posi As Integer = -1
        Dim laTabDel As New DataTable
        laTabDel.Columns.Add("")

        Select Case CategSelected

            Case Is = 1
                If IsNothing(elNode) = False Then

                    xObj = Split(elNode.FullPath, "\")

                    If xObj.Length = 1 Then
                        MsgBox("Please select a node on the second level to delete!", vbCritical, TitBox)
                        Exit Sub
                    Else

                        Dim enCuentra As DataRow
                        enCuentra = ModuPermit.Tables(0).Rows.Find(elNode.Parent.Name.ToUpper())
                        If IsNothing(enCuentra) = True Then
                            MsgBox("Sorry you are not allowed to add or delete on the selected module", vbCritical, TitBox)
                            Exit Sub
                        End If

                        'OJO, en el excel falta un sub para poder comparar vs los catálogos nuevos

                        If puedoEditar = False Then
                            MsgBox("Sorry you are not allowed to make changes to this template at this time, please try again later or wait for it to become available for editing", vbExclamation, TitBox)
                            Exit Sub
                        End If

                        pathDel = RaizFire
                        pathDel = pathDel & "/" & "catpro"
                        pathDel = pathDel & "/" & elNode.Parent.Name

                        Select Case xObj.Length
                            Case Is = 2
                                nodYav = elNode.Name
                                nodText = elNode.Text
                                cadFind = elNode.Parent.Name & "#" & elNode.Name
                                'elNode.Parent.Nodes.Add(nodyav, Form2.tabValue, 1, 1)
                            Case Is = 3
                                nodYav = elNode.Parent.Name
                                nodText = elNode.Parent.Text
                                cadFind = elNode.Parent.Parent.Name & "#" & elNode.Parent.Name
                                'elNode.Parent.Parent.Nodes.Add(nodYav, Form2.tabValue, 1, 1)
                        End Select

                        'pathDel = pathDel & "/" & nodYav
                        For i = 0 To catDs.Tables.Count - 1
                            If catDs.Tables(i).TableName = cadFind Then
                                Posi = i
                                Exit For
                            End If
                        Next

                        laTabDel.Rows.Add({nodYav})

                        x = MsgBox("ATTENTION!" & vbCrLf & "Are you sure you want to delete the whole catalog?" & vbCrLf & "This action can not be undone!!", vbYesNo + vbExclamation, "¿Delete full catalog?")

                        If x <> 6 Then Exit Sub

                        elRet = Await HazDeleteEnFbSimple(pathDel, elNode.Name)
                        'elRet = Await HazDeleteAFireBase(pathDel, laTabDel)
                        MySource.Remove(elNode.Name)
                        MySource.Remove(elNode.Text)

                        ToolStripTextBox1.AutoCompleteCustomSource = MySource
                        ToolStripTextBox1.AutoCompleteMode = AutoCompleteMode.Suggest
                        ToolStripTextBox1.AutoCompleteSource = AutoCompleteSource.CustomSource

                        If Posi >= 0 Then catDs.Tables.RemoveAt(Posi)

                        elNode.Remove()
                        elNode = Nothing
                        'eliminamos el
                        MsgBox(elRet, vbInformation, TitBox)

                    End If

                Else
                    MsgBox("Please select a node to delete on the corresponding structure!", vbCritical, TitBox)
                End If


            Case Is = 2
                'dependencias!, eliminar dependencias a nivel field code
                xObj = Split(elNode.FullPath, "\")

                If xObj.Length < 4 Then
                    'error!
                    MsgBox("Please select a field node!!", vbCritical, TitBox)
                    Exit Sub
                Else

                    cadCon = ""
                    pathDel = RaizFire
                    pathDel = pathDel & "/" & "dependencies"

                    Select Case xObj.Length

                        Case Is = 4
                            tablaFind = elNode.Parent.Name
                            tabName = elNode.Parent.Tag
                            cadFind = elNode.Parent.Parent.Name & "#" & elNode.Parent.Parent.Tag
                            objCode = elNode.Parent.Parent.Name
                            objName = elNode.Parent.Parent.Tag
                            campoCode = elNode.Name
                            campoNombre = elNode.Tag
                            'elTit = elNode.Parent.Parent.Name & " / " & elNode.Parent.Parent.Tag & " / " & tablaFind & " / " & tabName
                            For i = 0 To elNode.Nodes.Count - 1
                                If i <> 0 Then cadCon = cadCon & vbCrLf
                                cadCon = cadCon & elNode.Nodes(i).Text
                            Next

                        Case Is = 5
                            tablaFind = elNode.Parent.Parent.Name
                            tabName = elNode.Parent.Parent.Tag
                            cadFind = elNode.Parent.Parent.Parent.Name & "#" & elNode.Parent.Parent.Parent.Tag
                            objCode = elNode.Parent.Parent.Parent.Name
                            objName = elNode.Parent.Parent.Parent.Tag

                            campoCode = elNode.Parent.Name
                            campoNombre = elNode.Parent.Tag

                            For i = 0 To elNode.Parent.Nodes.Count - 1
                                If i <> 0 Then cadCon = cadCon & vbCrLf
                                cadCon = cadCon & elNode.Parent.Nodes(i).Text
                            Next

                    End Select

                    For i = 0 To tempDs.Tables.Count - 1
                        moduSel = Split(tempDs.Tables(i).TableName, "#")
                        If moduSel(0) = objCode Then

                            Dim enCuentra As DataRow
                            enCuentra = ModuPermit.Tables(0).Rows.Find(CStr(moduSel(2)).ToUpper())
                            If IsNothing(enCuentra) = True Then
                                MsgBox("Sorry you are not allowed to add or delete on the selected module", vbCritical, TitBox)
                                Exit Sub
                            End If

                            Exit For
                        End If
                    Next

                    If puedoEditar = False Then
                        MsgBox("Sorry you are not allowed to make changes to this template at this time, please try again later or wait for it to become available for editing", vbExclamation, TitBox)
                        Exit Sub
                    End If

                    x = MsgBox("ATTENTION!" & vbCrLf & "Are you sure you want to delete the dependencie of the field: " & campoCode & " / " & campoNombre & " of the table " & tablaFind & " / " & tabName & " of the object " & objCode & " / " & objName & ", of the conditional object: " & vbCrLf & cadCon & vbCrLf & "This action can-not be undone!." & vbCrLf & " Records already set could be affected. Are you sure?", vbExclamation + vbYesNo, TitBox)

                    If x <> 6 Then Exit Sub

                    'laTabDel.Rows.Add({campoCode})

                    pathDel = pathDel & "/" & objCode
                    pathDel = pathDel & "/" & tablaFind
                    elRet = Await HazDeleteEnFbSimple(pathDel, campoCode)
                    'elRet = Await HazDeleteAFireBase(pathDel, laTabDel)
                    If elRet = "Ok" Then

                        Select Case xObj.Length
                            Case Is = 4
                                elNode.Remove()

                            Case Is = 5
                                elNode.Parent.Remove()

                        End Select

                        elNode = Nothing
                        'eliminamos el
                        NodoNameActual = ""
                        MsgBox("Dependencie deleted!", vbInformation, TitBox)
                    Else
                        MsgBox(elRet, vbCritical, TitBox)
                    End If


                End If

            Case Is = 3
                'records
                'solo eliminar los registros a nivel records!!
                'Esto NO se puede, no se pueden eliminar registros!!, ó simplemente eliminar todos a nivel records!
                xObj = Split(elNode.FullPath, "\")
                If xObj.Length < 4 Then
                    MsgBox("Please select a table node!!", vbCritical, TitBox)
                    Exit Sub
                End If

                tablaFind = elNode.Parent.Parent.Name
                tabName = elNode.Parent.Parent.Tag

                objCode = elNode.Parent.Name
                objName = elNode.Parent.Tag

                campoCode = elNode.Name
                campoNombre = elNode.Tag

                For i = 0 To tempDs.Tables.Count - 1
                    moduSel = Split(tempDs.Tables(i).TableName, "#")
                    If moduSel(0) = objCode Then
                        Dim enCuentra As DataRow
                        enCuentra = ModuPermit.Tables(0).Rows.Find(CStr(moduSel(2)).ToUpper())
                        If IsNothing(enCuentra) = True Then
                            MsgBox("Sorry you are not allowed to add or delete on the selected module", vbCritical, TitBox)
                            Exit Sub
                        End If

                        Exit For
                    End If
                Next

                If puedoEditar = False Then
                    MsgBox("Sorry you are not allowed to make changes to this template at this time, please try again later or wait for it to become available for editing", vbExclamation, TitBox)
                    Exit Sub
                End If

                x = MsgBox("ATTENTION!!" & vbCrLf & "Are you sure you want to delete the full records of the table " & campoCode & " / " & campoNombre & " of the object " & objCode & " / " & objName & " of the company " & tablaFind & " / " & tabName & " ??" & vbCrLf & "This action can not be undone!!" & vbCrLf & "Are you sure?", vbExclamation + vbYesNo, TitBox)

                If x <> 6 Then Exit Sub

                pathDel = RaizFire
                pathDel = pathDel & "/" & "records"
                pathDel = pathDel & "/" & tablaFind
                pathDel = pathDel & "/" & objCode

                elRet = Await HazDeleteEnFbSimple(pathDel, campoCode)

                If elRet = "Ok" Then
                    elNode.Remove()
                    elNode = Nothing
                    NodoNameActual = ""
                    DataGridView1.DataSource = Nothing
                    DataGridView1.Rows.Clear()
                    DataGridView1.Columns.Clear()

                    MsgBox("Records gone!!", vbInformation, TitBox)

                Else
                    MsgBox("Error deleting: " & vbCrLf & elRet, vbCritical, TitBox)
                End If

            Case Is = 4
                'templates!, eliminar templates completos a nivel tabla!, OJO, esto es peligroso!!
                'Avisar 2 veces de la afectacion posible a registros y dependencias!!
                xObj = Split(elNode.FullPath, "\")

                If xObj.Length < 3 Then
                    MsgBox("Please select a table node!!", vbCritical, TitBox)
                    Exit Sub
                End If

                pathDel = RaizFire
                pathDel = pathDel & "/" & "templates"

                Select Case xObj.Length

                    Case Is = 3
                        objCode = elNode.Parent.Name
                        objName = elNode.Parent.Tag
                        tablaFind = elNode.Name
                        tabName = elNode.Tag
                        moduSel = Split(elNode.Parent.Tag, "#")

                    Case Is = 4
                        objCode = elNode.Parent.Parent.Name
                        objName = elNode.Parent.Parent.Tag
                        tablaFind = elNode.Parent.Name
                        tabName = elNode.Parent.Tag
                        moduSel = Split(elNode.Parent.Parent.Tag, "#")
                End Select

                Dim enCuentra As DataRow
                enCuentra = ModuPermit.Tables(0).Rows.Find(CStr(moduSel(1)).ToUpper())
                If IsNothing(enCuentra) = True Then
                    MsgBox("Sorry you are not allowed to make changes on the selected module", vbCritical, TitBox)
                    Exit Sub
                End If

                If puedoEditar = False Then
                    MsgBox("Sorry you are not allowed to make changes to this template at this time, please try again later or wait for it to become available for editing", vbExclamation, TitBox)
                    Exit Sub
                End If

                x = MsgBox("ATTENTION!!" & vbCrLf & "Are you sure you want to delete the complete template: " & tablaFind & " / " & tabName & " of the object " & objCode & " / " & objName & " ??" & vbCrLf & "This action can not be undone!!" & vbCrLf & "Notice that dependencies related to this table or records already set of this table will be affected!!. Are you sure?", vbExclamation + vbYesNo, TitBox)

                If x <> 6 Then Exit Sub

                x = MsgBox("Are you sure!?", vbExclamation + vbYesNo, TitBox)

                If x <> 6 Then Exit Sub

                pathDel = pathDel & "/" & objCode

                elRet = Await HazDeleteEnFbSimple(pathDel, tablaFind)

                If elRet = "Ok" Then
                    'hacemos reload!
                    Select Case xObj.Length
                        Case Is = 3
                            elNode.Remove()
                        Case Is = 4
                            elNode.Parent.Remove()
                    End Select

                    NodoNameActual = ""
                    elNode = Nothing
                    'eliminamos el
                    MsgBox("Template deleted!", vbInformation, TitBox)
                Else
                    MsgBox("Error deleting: " & vbCrLf & elRet, vbCritical, TitBox)
                End If


            Case Is = 5
                'este NO aplica!!

        End Select

    End Sub

    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
        Try
            If e.Control And (e.KeyCode = Keys.C) Then
                Dim d As DataObject = DataGridView1.GetClipboardContent()
                Clipboard.SetDataObject(d)
                e.Handled = True
            ElseIf (e.Control And e.KeyCode = Keys.V) Then
                PasteUnboundRecords()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PasteUnboundRecords()
        Dim j As Integer
        Try

            Select Case CategSelected
                Case Is = 0
                    MsgBox("Please select a category first!", vbCritical, TitBox)

                Case Is = 1
                    Dim rowLines As String() = Clipboard.GetText(TextDataFormat.Text).Split(New String(0) {vbCr & vbLf}, StringSplitOptions.None)
                    Dim currentRowIndex As Integer = (If(DataGridView1.CurrentRow IsNot Nothing, DataGridView1.CurrentRow.Index, 0))
                    Dim currentColumnIndex As Integer = (If(DataGridView1.CurrentCell IsNot Nothing, DataGridView1.CurrentCell.ColumnIndex, 0))
                    Dim currentColumnCount As Integer = DataGridView1.Columns.Count

                    Dim colAs As String() = rowLines(0).Split(New String(0) {vbTab}, StringSplitOptions.None)

                    If currentColumnIndex <> 0 Then
                        MsgBox("Please paste EXACTLY " & catDs.Tables(cataNombre).Columns.Count - 2 & " columns starting from the first column at the left!", vbCritical, TitBox)
                        Exit Sub
                    End If

                    If colAs.Length > catDs.Tables(cataNombre).Columns.Count - 2 Then
                        MsgBox("Please paste EXACTLY " & catDs.Tables(cataNombre).Columns.Count - 2 & " columns starting from the first column at the left!", vbCritical, TitBox)
                        Exit Sub
                    End If

                    'DataGridView1.AllowUserToAddRows = True
                    If rowLines.Length = 2 Then
                        DataGridView1.Rows.Add(1)
                        DataGridView1.Rows(DataGridView1.Rows.Count - 2).HeaderCell.Value = CStr(currentRowIndex + 1)
                    Else
                        If DataGridView1.Rows.Count - currentRowIndex < rowLines.Length - 1 Then
                            For i = 1 To (rowLines.Length - 1) - (DataGridView1.Rows.Count - currentRowIndex) + 1
                                DataGridView1.Rows.Add(1)
                                DataGridView1.Rows(currentRowIndex + i - 1).HeaderCell.Value = CStr(currentRowIndex + i)
                            Next
                        End If
                    End If



                    'DataGridView1.Rows.Add(1)

                    j = 0
                    For rowLine As Integer = 0 To rowLines.Length - 1
                        If rowLine = rowLines.Length - 1 AndAlso String.IsNullOrEmpty(rowLines(rowLine)) Then
                            Exit For
                        End If
                        Application.DoEvents()
                        Dim columnsData As String() = rowLines(rowLine).Split(New String(0) {vbTab}, StringSplitOptions.None)
                        If (currentColumnIndex + columnsData.Length) > DataGridView1.Columns.Count Then
                            For columnCreationCounter As Integer = 0 To ((currentColumnIndex + columnsData.Length) - currentColumnCount) - 1
                                If columnCreationCounter = rowLines.Length - 1 Then
                                    Exit For
                                End If
                            Next
                        End If
                        If DataGridView1.Rows.Count > (currentRowIndex + rowLine) Then
                            For columnsDataIndex As Integer = 0 To columnsData.Length - 1
                                If currentColumnIndex + columnsDataIndex <= DataGridView1.Columns.Count - 1 Then
                                    DataGridView1.Rows(currentRowIndex + rowLine).Cells(currentColumnIndex + columnsDataIndex).Value = columnsData(columnsDataIndex)
                                End If
                            Next
                        Else
                            Dim pasteCells As String() = New String(DataGridView1.Columns.Count - 1) {}
                            For cellStartCounter As Integer = currentColumnIndex To DataGridView1.Columns.Count - 1
                                If columnsData.Length > (cellStartCounter - currentColumnIndex) Then
                                    pasteCells(cellStartCounter) = columnsData(cellStartCounter - currentColumnIndex)
                                End If
                            Next
                        End If
                    Next

                Case Is = 2
                    'dependencias


                Case Is = 3
                    'records
                    Dim rowLines As String() = Clipboard.GetText(TextDataFormat.Text).Split(New String(0) {vbCr & vbLf}, StringSplitOptions.None)
                    Dim currentRowIndex As Integer = (If(DataGridView1.CurrentRow IsNot Nothing, DataGridView1.CurrentRow.Index, 0))
                    Dim currentColumnIndex As Integer = (If(DataGridView1.CurrentCell IsNot Nothing, DataGridView1.CurrentCell.ColumnIndex, 0))
                    Dim currentColumnCount As Integer = DataGridView1.Columns.Count

                    Dim colAs As String() = rowLines(0).Split(New String(0) {vbTab}, StringSplitOptions.None)

                    If currentColumnIndex <> 0 Then
                        MsgBox("Please paste only 5 columns data starting from 'Table Code' column", vbCritical, TitBox)
                        Exit Sub
                    End If

                    If colAs.Length > DataGridView1.Columns.Count - 3 Then
                        MsgBox("Please paste exact " & DataGridView1.Columns.Count - 3 & " columns data starting on the first column", vbCritical, TitBox)
                        Exit Sub
                    End If

                    If DataGridView1.Rows.Count - currentRowIndex < rowLines.Length - 1 Then
                        For i = 1 To (rowLines.Length - 1) - (DataGridView1.Rows.Count - currentRowIndex) + 1
                            DataGridView1.Rows.Add(1)
                            DataGridView1.Rows(currentRowIndex + i - 1).HeaderCell.Value = CStr(currentRowIndex + i)
                        Next
                    End If

                    'DataGridView1.Rows.Add(1)

                    j = 0
                    For rowLine As Integer = 0 To rowLines.Length - 1
                        If rowLine = rowLines.Length - 1 AndAlso String.IsNullOrEmpty(rowLines(rowLine)) Then
                            Exit For
                        End If
                        Application.DoEvents()
                        Dim columnsData As String() = rowLines(rowLine).Split(New String(0) {vbTab}, StringSplitOptions.None)
                        If (currentColumnIndex + columnsData.Length) > DataGridView1.Columns.Count Then
                            For columnCreationCounter As Integer = 0 To ((currentColumnIndex + columnsData.Length) - currentColumnCount) - 1
                                If columnCreationCounter = rowLines.Length - 1 Then
                                    Exit For
                                End If
                            Next
                        End If
                        If DataGridView1.Rows.Count > (currentRowIndex + rowLine) Then
                            For columnsDataIndex As Integer = 0 To columnsData.Length - 1
                                If currentColumnIndex + columnsDataIndex <= DataGridView1.Columns.Count - 1 Then
                                    DataGridView1.Rows(currentRowIndex + rowLine).Cells(currentColumnIndex + columnsDataIndex).Value = columnsData(columnsDataIndex)
                                End If
                            Next
                        Else
                            Dim pasteCells As String() = New String(DataGridView1.Columns.Count - 1) {}
                            For cellStartCounter As Integer = currentColumnIndex To DataGridView1.Columns.Count - 1
                                If columnsData.Length > (cellStartCounter - currentColumnIndex) Then
                                    pasteCells(cellStartCounter) = columnsData(cellStartCounter - currentColumnIndex)
                                End If
                            Next
                        End If
                    Next


                Case Is = 4
                    'templates
                    Dim rowLines As String() = Clipboard.GetText(TextDataFormat.Text).Split(New String(0) {vbCr & vbLf}, StringSplitOptions.None)
                    Dim currentRowIndex As Integer = (If(DataGridView1.CurrentRow IsNot Nothing, DataGridView1.CurrentRow.Index, 0))
                    Dim currentColumnIndex As Integer = (If(DataGridView1.CurrentCell IsNot Nothing, DataGridView1.CurrentCell.ColumnIndex, 0))
                    Dim currentColumnCount As Integer = DataGridView1.Columns.Count

                    Dim colAs As String() = rowLines(0).Split(New String(0) {vbTab}, StringSplitOptions.None)

                    If currentColumnIndex <> 0 Then
                        MsgBox("Please paste only 5 columns data starting on 'Table Code' column", vbCritical, TitBox)
                        Exit Sub
                    End If

                    If colAs.Length > 2 Then
                        MsgBox("Please paste only 2 columns data starting on 'Table Code' column", vbCritical, TitBox)
                        Exit Sub
                    End If

                    'DataGridView1.AllowUserToAddRows = True
                    If DataGridView1.Rows.Count - currentRowIndex < rowLines.Length - 1 Then
                        For i = 1 To (rowLines.Length - 1) - (DataGridView1.Rows.Count - currentRowIndex) + 1
                            DataGridView1.Rows.Add(1)
                            DataGridView1.Rows(currentRowIndex + i - 1).HeaderCell.Value = CStr(currentRowIndex + i)
                        Next
                    End If

                    'DataGridView1.Rows.Add(1)

                    j = 0
                    For rowLine As Integer = 0 To rowLines.Length - 1
                        If rowLine = rowLines.Length - 1 AndAlso String.IsNullOrEmpty(rowLines(rowLine)) Then
                            Exit For
                        End If
                        Application.DoEvents()
                        Dim columnsData As String() = rowLines(rowLine).Split(New String(0) {vbTab}, StringSplitOptions.None)
                        If (currentColumnIndex + columnsData.Length) > DataGridView1.Columns.Count Then
                            For columnCreationCounter As Integer = 0 To ((currentColumnIndex + columnsData.Length) - currentColumnCount) - 1
                                If columnCreationCounter = rowLines.Length - 1 Then
                                    Exit For
                                End If
                            Next
                        End If
                        If DataGridView1.Rows.Count > (currentRowIndex + rowLine) Then
                            For columnsDataIndex As Integer = 0 To columnsData.Length - 1
                                If currentColumnIndex + columnsDataIndex <= DataGridView1.Columns.Count - 1 Then
                                    DataGridView1.Rows(currentRowIndex + rowLine).Cells(currentColumnIndex + columnsDataIndex).Value = columnsData(columnsDataIndex)
                                End If
                            Next
                        Else
                            Dim pasteCells As String() = New String(DataGridView1.Columns.Count - 1) {}
                            For cellStartCounter As Integer = currentColumnIndex To DataGridView1.Columns.Count - 1
                                If columnsData.Length > (cellStartCounter - currentColumnIndex) Then
                                    pasteCells(cellStartCounter) = columnsData(cellStartCounter - currentColumnIndex)
                                End If
                            Next
                        End If
                    Next

                Case Is = 5

            End Select


        Catch ex As Exception
            'Log Exception
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Async Sub ToolStripButton8_Click(sender As Object, e As EventArgs) Handles ToolStripButton8.Click



        Dim xObj As Object
        Dim i As Integer
        Dim cadFind As String = ""
        Dim unResp As String = ""
        Dim xResp As String = ""
        Dim Posi As Integer = -1
        Dim nomCat As String = ""
        Dim z As Long
        Dim encuentra As DataRow
        Dim X As Integer = 0
        Dim unReg As String = ""
        Dim verFal As String = ""
        Dim tabNombre As String = ""
        Dim tabCode As String = ""
        Dim moduSel As Object = ""
        Dim unPaths As String = ""

        Select Case CategSelected

            Case Is = 1
                'catalogos!

                xObj = Split(elNode.FullPath, "\")

                If xObj.Length = 1 Then
                    'seleccionó el nodo principal, nada que agregar!
                    MsgBox("Please select a catalog to update first!", vbCritical, TitBox)
                    Exit Sub
                End If

                'Dim enCuentra As DataRow
                encuentra = ModuPermit.Tables(0).Rows.Find(elNode.Parent.Name.ToUpper())
                If IsNothing(enCuentra) = True Then
                    MsgBox("Sorry you are not allowed to make changes on the selected module", vbCritical, TitBox)
                    Exit Sub
                End If

                If puedoEditar = False Then
                    MsgBox("Sorry, you are not allowed to edit this catalog at this time, please try again later or wait for it to be available to edit.", vbExclamation, TitBox)
                    Exit Sub
                End If

                Await SigoVivo()
                Module5.AgregaTiempo()

                writeDs.Tables(0).Rows.Clear() 'tabla de los niveles
                writeDs.Tables(1).Rows.Clear() 'nuevos
                writeDs.Tables(2).Rows.Clear() 'deleted
                writeDs.Tables(3).Rows.Clear() 'current
                writeDs.Tables(4).Rows.Clear() 'updated

                writeDs.Tables(0).Rows.Add({"catalogs"}) 'nivel 1
                writeDs.Tables(0).Rows.Add({CStr(xObj(0))}) 'nivel 2

                If xObj.Length = 2 Then
                    cadFind = elNode.Parent.Name & "#" & elNode.Name
                    writeDs.Tables(0).Rows.Add({elNode.Name}) 'nivl 3
                    nomCat = elNode.Text
                Else
                    cadFind = elNode.Parent.Parent.Name & "#" & elNode.Parent.Name
                    writeDs.Tables(0).Rows.Add({elNode.Parent.Name}) 'nivel 3
                    nomCat = elNode.Parent.Text
                End If

                For i = 0 To catDs.Tables.Count - 1
                    If catDs.Tables(i).TableName = cadFind Then
                        Posi = i
                        Exit For
                    End If
                Next

                'Vas a necesitar una tabla con un campo llave al inicio y que evalúe que todas las combinaciones
                'de la primera a la ultima -1 columnas NO se repita, si existe alguna repetidda, pues Angg, error!
                'Eliminar los renglones que tengan alguna columna vacía!
                'Y eliminar duplicados también!
                QuitaDuplisCats(DataGridView1, unResp)
                'Call QuitaDuplicadosDeGrid(DataGridView1, 0, unResp)'este era el anterior!

                DataGridView1.AllowUserToAddRows = False
                DataGridView1.AllowUserToDeleteRows = False

                If DataGridView1.Rows.Count = 0 Then
                    MsgBox("Sorry, there must be at least one record per catalog, if you want to delete the whole table please select the node and click on 'Delete node' button!", vbCritical, TitBox)
                    DataGridView1.AllowUserToAddRows = True
                    DataGridView1.AllowUserToDeleteRows = True
                    Exit Sub
                End If


                'AQUI ME QUEDE!!
                'tabla de registros nuevos
                'tabla de registros para eliminar
                'tabla de registros para actualizar!
                Dim remoDt As DataTable
                Dim updDt As DataTable
                Dim niuDt As DataTable
                niuDt = catDs.Tables(cataNombre).Clone()
                updDt = catDs.Tables(cataNombre).Clone()
                remoDt = catDs.Tables(cataNombre).Clone()
                niuDt.Columns.RemoveAt(1) 'eliminamos el campo de fbkey, por que son nuevos!
                updDt.Columns.RemoveAt(1) 'eliminamos el campo de yave local
                TomaInfoParaCats(catDs.Tables(cataNombre), DataGridView1, niuDt, remoDt, updDt)

                If (niuDt.Rows.Count + updDt.Rows.Count + remoDt.Rows.Count) > 0 Then

                    toyTrabajando = True

                    Cursor.Current = Cursors.WaitCursor
                    ToolStripLabel1.Visible = True
                    ToolStripLabel1.Text = "Uploading..."

                    ToolStripButton8.Enabled = False

                    If niuDt.Rows.Count > 0 Then
                        unPaths = RaizFire
                        unPaths = unPaths & "/" & "catpro"
                        unPaths = unPaths & "/" & elNode.Parent.Name
                        unPaths = unPaths & "/" & elNode.Name
                        xResp = Await HazPostEnFireBaseConPathYColumnas(unPaths, niuDt, "records", 0)
                    End If

                    If remoDt.Rows.Count > 0 Then
                        unPaths = RaizFire
                        unPaths = unPaths & "/" & "catpro"
                        unPaths = unPaths & "/" & elNode.Parent.Name
                        unPaths = unPaths & "/" & elNode.Name
                        unPaths = unPaths & "/" & "records"
                        xResp = Await HazDeleteAFireBase(unPaths, remoDt)

                    End If


                    If updDt.Rows.Count > 0 Then
                        unPaths = RaizFire
                        unPaths = unPaths & "/" & "catpro"
                        unPaths = unPaths & "/" & elNode.Parent.Name
                        unPaths = unPaths & "/" & elNode.Name
                        unPaths = unPaths & "/" & "records"
                        xResp = Await HazUpdateEnFireBaseConYave(unPaths, 0, updDt)

                    End If

                    toyTrabajando = False

                    ToolStripButton8.Enabled = True
                    Cursor.Current = Cursors.Default
                    ToolStripLabel1.Visible = False
                    ToolStripLabel1.Text = "Ready"

                    MsgBox("Update complete!", vbInformation, TitBox)

                    unReg = Await CargaOpcion(CategSelected)

                Else

                    MsgBox("No changes detected!", vbInformation, TitBox)

                End If

                DataGridView1.AllowUserToAddRows = True
                DataGridView1.AllowUserToDeleteRows = True



            Case Is = 2
                'dependencias
                xObj = Split(elNode.FullPath, "\")
                If xObj.Length = 1 Or xObj.Length = 2 Then
                    'seleccionó el nodo principal, nada que agregar!
                    MsgBox("Please select a dependecie table to update first!", vbCritical, TitBox)
                    Exit Sub
                End If



            Case Is = 3
                'records



            Case Is = 4
                'templates

                xObj = Split(elNode.FullPath, "\")

                If xObj.Length = 1 Or xObj.Length = 2 Then
                    'seleccionó el nodo principal, nada que agregar!
                    MsgBox("Please select first a template to update!", vbCritical, TitBox)
                    Exit Sub
                End If

                writeDs.Tables(0).Rows.Clear() 'tabla de los niveles
                writeDs.Tables(1).Rows.Clear() 'nuevos
                writeDs.Tables(2).Rows.Clear() 'deleted
                writeDs.Tables(3).Rows.Clear() 'current

                writeDs.Tables(0).Rows.Add({"templates"}) 'nivel 1

                unPaths = RaizFire
                unPaths = unPaths & "/templates"

                Select Case xObj.Length
                    Case Is = 2
                        cadFind = elNode.Name & "#" & elNode.Tag
                        writeDs.Tables(0).Rows.Add({elNode.Name}) 'nivel 2

                    Case Is = 3
                        moduSel = Split(elNode.Parent.Tag, "#")
                        tabCode = elNode.Name
                        tabNombre = elNode.Tag
                        cadFind = elNode.Parent.Name & "#" & elNode.Parent.Tag
                        writeDs.Tables(0).Rows.Add({elNode.Parent.Name})
                        writeDs.Tables(0).Rows.Add({tabCode})

                        unPaths = unPaths & "/" & elNode.Parent.Name
                        unPaths = unPaths & "/" & tabCode

                    Case Is = 4
                        moduSel = Split(elNode.Parent.Parent.Tag, "#")
                        tabCode = elNode.Parent.Name
                        tabNombre = elNode.Parent.Tag
                        cadFind = elNode.Parent.Parent.Name & "#" & elNode.Parent.Parent.Tag
                        writeDs.Tables(0).Rows.Add({elNode.Parent.Parent.Name})
                        writeDs.Tables(0).Rows.Add({tabCode})

                        unPaths = unPaths & "/" & elNode.Parent.Parent.Name
                        unPaths = unPaths & "/" & tabCode

                End Select

                'Dim enCuentra As DataRow
                encuentra = ModuPermit.Tables(0).Rows.Find(CStr(moduSel(1)).ToUpper())
                If IsNothing(enCuentra) = True Then
                    MsgBox("Sorry you are not allowed to make changes on the selected template", vbCritical, TitBox)
                    Exit Sub
                End If

                If puedoEditar = False Then
                    MsgBox("Sorry, you are not allowed to edit this template at this time, please try again later or when it becomes available.", vbExclamation, TitBox)
                    Exit Sub
                End If

                Await SigoVivo()
                Module5.AgregaTiempo()

                For i = 0 To tempDs.Tables.Count - 1
                    If tempDs.Tables(i).TableName = cadFind Then
                        Posi = i
                        Exit For
                    End If
                Next

                toyTrabajando = True

                'Segun yo este metodo ya no sería necesario!!
                'Call QuitaDuplicadosMultiplesGrid(DataGridView1, "0", unResp)

                DataGridView1.AllowUserToAddRows = False
                DataGridView1.AllowUserToDeleteRows = False

                'OJO, este metodo de abajo va a cambiar por uno más robusto, que tome toodos los campos necesarios!!
                'nuevos y deleted

                For i = 0 To DataGridView1.Rows.Count - 1
                    DataGridView1.Rows(i).HeaderCell.Value = CStr(i + 1)
                Next

                Call TomaInfoParaTemplates(DataGridView1, writeDs, 3, 0, 3, 13)

                'tomamos la info!!, 
                Cursor.Current = Cursors.WaitCursor
                ToolStripLabel1.Visible = True
                ToolStripLabel1.Text = "Uploading..."
                ToolStripButton8.Enabled = False

                'se debe re-escribir!!

                Await HazPutEnFbSimple(unPaths, "TableName", tabNombre)

                xResp = Await HazPutEnFireBasePathYColumnas(unPaths, writeDs.Tables(3), 0)

                toyTrabajando = False

                MsgBox(xResp, vbInformation, TitBox)

                ToolStripButton8.Enabled = True

                Cursor.Current = Cursors.Default
                ToolStripLabel1.Visible = False
                ToolStripLabel1.Text = "Ready"

                unReg = Await CargaOpcion(CategSelected)

                If unReg = "ok" Then
                    Call ReloadOneNode(CategSelected)
                End If

                DataGridView1.AllowUserToAddRows = False ' True
                DataGridView1.AllowUserToDeleteRows = False ' True

                'comparativa local!
                'todo lo nuevo vs lo que estaba, 


        End Select

    End Sub

    Private Async Sub MuestraRecords()

        If CategSelected <> 3 Then Exit Sub
        'bajar info de firebase
        If IsNothing(elNode) = True Then
            MsgBox("Please select a valid node!", vbCritical, TitBox)
            Exit Sub
        End If

        Dim xObj As Object
        xObj = Split(elNode.FullPath, "\")

        If xObj.Length < 4 Then
            MsgBox("Please select a node at a table level!", vbCritical, TitBox)
            Exit Sub
        End If

        Dim elCamino As String = RaizFire
        elCamino = elCamino & "/records/" & elNode.Parent.Parent.Name 'compania
        elCamino = elCamino & "/" & elNode.Parent.Name 'objeto
        elCamino = elCamino & "/" & elNode.Name 'la tabla
        elCamino = elCamino & "/records"

        usaDataset.Tables(0).Rows.Clear()
        usaDataset.Tables(0).Rows.Add({"records"})
        usaDataset.Tables(0).Rows.Add({elNode.Parent.Parent.Name})
        usaDataset.Tables(0).Rows.Add({elNode.Parent.Name})
        usaDataset.Tables(0).Rows.Add({elNode.Name})
        usaDataset.Tables(0).Rows.Add({"records"})

        'xdepeds

        Dim xDs As New DataSet
        xDs = Await PullUrlWs(usaDataset, "records")

        DataGridView1.DataSource = Nothing
        DataGridView1.Rows.Clear()

        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AllowUserToDeleteRows = False

        If xDs.Tables.Count > 0 Then
            Dim k As Integer = 0

            Dim bS As New BindingSource

            RecDt.Rows.Clear()
            RecDt.Columns.Clear()
            RecDt = xDs.Tables(0)
            'agregar las columnas extras!!

            'RecDt.Columns.Add("ROWOK")
            'RecDt.Columns.Add("ROWREP")
            'RecDt.Columns.Add("ROWCOMMS")

            bS.DataSource = RecDt ' xDs.Tables(0)

            DataGridView1.AutoGenerateColumns = False

            DataGridView1.DataSource = bS

            For i = 0 To DataGridView1.Rows.Count - 1
                DataGridView1.Rows(i).HeaderCell.Value = CStr(i + 1)
                DataGridView1.Rows(i).Tag = "NO OK" 'empezamos suponiendo que estan mal!
            Next


        Else
            MsgBox("No records found for this object!!", vbInformation, TitBox)
        End If

        DataGridView1.RowHeadersWidth = 70

        DataGridView1.AllowUserToAddRows = puedoEditar ' True
        DataGridView1.AllowUserToDeleteRows = puedoEditar 'True

        For i = 0 To DataGridView1.Columns.Count - 1
            DataGridView1.Columns(i).ReadOnly = Not (puedoEditar)
        Next

    End Sub

    Private Async Sub ToolStripButton11_Click(sender As Object, e As EventArgs) Handles ToolStripButton11.Click

        'este boton va ser para importar registros de un excel


        If IsNothing(elNode) = True Then
            MsgBox("Please select a valid node!", vbCritical, TitBox)
            Exit Sub
        End If

        Await SigoVivo()
        Module5.AgregaTiempo()

        Dim xCad As String = ""
        Dim xPath As String = ""
        Dim xObj As Object
        xObj = Split(elNode.FullPath, "\")
        Dim enCuentra As DataRow

        Select Case ToolStripComboBox1.SelectedIndex
            Case Is = 1
                'catalogos
                If xObj.Length < 2 Then
                    MsgBox("Please select a node at a catalog level!", vbCritical, TitBox)
                    Exit Sub
                End If

                enCuentra = ModuPermit.Tables(0).Rows.Find(CStr(catDs.Tables(cataNombre).ExtendedProperties.Item("ModuleCode")).ToUpper())
                If IsNothing(enCuentra) = True Then
                    MsgBox("Sorry you are not allowed to add catalogs on the selected module", vbCritical, TitBox)
                    Exit Sub
                End If


                If puedoEditar = False Then
                    MsgBox("Sorry, you are not allowed to edit this object at this time, please try again later or wait for it to be available to edit.", vbExclamation, TitBox)
                    Exit Sub
                End If

                Dim diLOg As New OpenFileDialog
                Dim oFileName As String
                oFileName = ""
                diLOg.InitialDirectory = Environment.SpecialFolder.MyDocuments
                diLOg.Filter = "Excel CSV File (*.csv)|*.csv" '"Excel File (*.xls)|*.xls"

                'al importar hay varias opciones
                '1. Importar una hoja simple
                '2. Importar todo el template!
                '3. Darle en importar y guardar
                '4. Exportar una hoja simple
                '5. Exportar todo el template!

                Dim borroRecords As Boolean = False

                If diLOg.ShowDialog() = Windows.Forms.DialogResult.OK Then

                    Dim path As String = diLOg.FileName

                    Dim elDt As New DataTable

                    elDt = LoadFromCSV(path)

                    If elDt.Rows.Count = 0 Then Exit Sub

                    DataGridView1.DataSource = Nothing
                    If borroRecords = True Then
                        DataGridView1.Rows.Clear()
                    End If

                    estoyAgregandoRows = True

                    DataGridView1.AllowUserToAddRows = False
                    DataGridView1.AllowUserToDeleteRows = False

                    Dim bS As New BindingSource

                    bS.DataSource = elDt

                    DataGridView1.AutoGenerateColumns = False

                    DataGridView1.DataSource = bS

                    For i = 0 To DataGridView1.Rows.Count - 1
                        DataGridView1.Rows(i).HeaderCell.Value = CStr(i + 1)
                    Next

                    estoyAgregandoRows = False

                    DataGridView1.AllowUserToAddRows = puedoEditar ' True
                    DataGridView1.AllowUserToDeleteRows = puedoEditar ' True

                End If


            Case Is = 2
                'dependencias


            Case Is = 3
                'records

                Select Case xObj.Length
                    Case Is = 1
                        'crazy!!, importar todas las companias con todos sus objetos!


                    Case Is = 2
                        'No tan crazy, importar todos los objetos de 1 compañía
                        'este seria un ciclo por excel!


                    Case Is = 3
                        'Here, importar todas las tablas de un objeto completo!
                        enCuentra = ModuPermit.Tables(0).Rows.Find(moduloSelek.ToUpper())
                        If IsNothing(enCuentra) = True Then
                            MsgBox("Sorry you are not allowed to import records to the selected module!!", vbCritical, TitBox)
                            Exit Sub
                        End If

                        If PussyTemp < 0 Then
                            MsgBox("Please choose a valid template!!", vbCritical, TitBox)
                            Exit Sub
                        End If

                        If elNode.Nodes.Count = 0 Then
                            MsgBox("This template does not contain any tables!!, please choose another one!!", vbCritical, TitBox)
                            Exit Sub
                        End If


                        'aquii, AAANTESS de iniciar, verificar que TOODAS las hojas esten NO ocupadas
                        'SOLO si NO estan ocupadas, entonces permite la importación masiva!!
                        'Si NO estan ocupadas deben 'ocuparse' para que nadie mas las pueda editar!
                        'ciclo para validar que NO esten ocupados!!

                        'primero!!, crear el dataset con la estructura del template vacío!!
                        Dim puedoAvanzar As Boolean = True
                        Dim cadX As String = ""
                        Dim allDs As New DataSet
                        Dim filtrosDt As New DataTable
                        filtrosDt = tempDs.Tables(PussyTemp).Clone()

                        '.ExtendedProperties.Item("Found")
                        allDs.ExtendedProperties.Add("AllOk", False)
                        allDs.ExtendedProperties.Add("TabsFound", False)
                        allDs.ExtendedProperties.Add("ColsFound", False)

                        For i = 0 To elNode.Nodes.Count - 1
                            filtrosDt.Rows.Clear()
                            Dim result() As DataRow = tempDs.Tables(PussyTemp).Select("TableCode = '" & elNode.Nodes(i).Name & "'", "Position ASC")
                            For Each row As DataRow In result
                                filtrosDt.ImportRow(row)
                            Next

                            allDs.Tables.Add(elNode.Nodes(i).Name)
                            allDs.Tables(allDs.Tables.Count - 1).ExtendedProperties.Add("TableCode", elNode.Nodes(i).Name)
                            allDs.Tables(allDs.Tables.Count - 1).ExtendedProperties.Add("TableName", "")
                            allDs.Tables(allDs.Tables.Count - 1).ExtendedProperties.Add("Found", False)
                            allDs.Tables(allDs.Tables.Count - 1).ExtendedProperties.Add("PuedoEdit", False)

                            For j = 0 To filtrosDt.Rows.Count - 1
                                allDs.Tables(allDs.Tables.Count - 1).Columns.Add(filtrosDt.Rows(j).Item("FieldCode"))
                                allDs.Tables(allDs.Tables.Count - 1).Columns(allDs.Tables(allDs.Tables.Count - 1).Columns.Count - 1).ExtendedProperties.Add("FieldCode", filtrosDt.Rows(j).Item("FieldCode"))
                                allDs.Tables(allDs.Tables.Count - 1).Columns(allDs.Tables(allDs.Tables.Count - 1).Columns.Count - 1).ExtendedProperties.Add("HeaderText", filtrosDt.Rows(j).Item("FieldName"))
                                allDs.Tables(allDs.Tables.Count - 1).Columns(allDs.Tables(allDs.Tables.Count - 1).Columns.Count - 1).ExtendedProperties.Add("MOC", filtrosDt.Rows(j).Item("MOC"))
                                allDs.Tables(allDs.Tables.Count - 1).Columns(allDs.Tables(allDs.Tables.Count - 1).Columns.Count - 1).ExtendedProperties.Add("Found", False)
                                allDs.Tables(allDs.Tables.Count - 1).ExtendedProperties.Item("TableName") = filtrosDt.Rows(j).Item("TableName")
                            Next

                            allDs.Tables(allDs.Tables.Count - 1).ExtendedProperties.Add("FromCol", 0)
                            allDs.Tables(allDs.Tables.Count - 1).ExtendedProperties.Add("ToCol", allDs.Tables(allDs.Tables.Count - 1).Columns.Count - 1)

                            'agregar un campo llave!!?
                            allDs.Tables(allDs.Tables.Count - 1).Columns.Add("CampoYave")
                            allDs.Tables(allDs.Tables.Count - 1).Columns(allDs.Tables(allDs.Tables.Count - 1).Columns.Count - 1).ExtendedProperties.Add("Found", True)

                            xPath = RaizFire
                            xPath = xPath & "/" & "inuse"
                            xPath = xPath & "/" & "records"
                            xPath = xPath & "/" & compaSelekted
                            xPath = xPath & "/" & objetoSelek
                            xPath = xPath & "/" & elNode.Nodes(i).Name
                            editDs.Clear()
                            editDs = Await PullDtFb(xPath, "inuse")

                            allDs.Tables(allDs.Tables.Count - 1).ExtendedProperties.Item("PuedoEdit") = Await siPuedoEditar(xPath)

                            If allDs.Tables(allDs.Tables.Count - 1).ExtendedProperties.Item("PuedoEdit") = False Then
                                If cadX <> "" Then cadX = cadX & vbCrLf
                                cadX = cadX & "Table " & allDs.Tables(allDs.Tables.Count - 1).TableName & " is currently in use by " & editDs.Tables(0).Rows(0).Item(3)
                                puedoAvanzar = False
                            End If

                        Next


                        If puedoAvanzar = False Then
                            MsgBox("Massive import not possible!!" & vbCrLf & "One or more tables are currently in use by your colleages, below the detail" & vbCrLf & cadX & vbCrLf & "Please wait for them to become available!, try again later", vbExclamation, TitBox)
                            Exit Sub
                        End If

                        inUseDt.Rows(0).Item(0) = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                        inUseDt.Rows(0).Item(1) = ""
                        inUseDt.Rows(0).Item(2) = UsuarioCorreo
                        inUseDt.Rows(0).Item(3) = UsuarioNombre
                        inUseDt.Rows(0).Item(4) = "X"

                        'OJO, aqui antes falta bloquear masivo también!!
                        For i = 0 To allDs.Tables.Count - 1

                            xPath = RaizFire
                            xPath = xPath & "/" & "inuse"
                            xPath = xPath & "/" & "records"
                            xPath = xPath & "/" & compaSelekted
                            xPath = xPath & "/" & objetoSelek
                            xPath = xPath & "/" & allDs.Tables(i).TableName

                            Await HazPost1Set(xPath, inUseDt, -1)

                        Next

                        Dim diLOg As New OpenFileDialog
                        diLOg.InitialDirectory = Environment.SpecialFolder.MyDocuments
                        diLOg.Filter = "Excel File (*.xlsx)|*.xlsx" '"Excel File (*.xls)|*.xls"

                        If diLOg.ShowDialog = System.Windows.Forms.DialogResult.OK Then

                            ToolStripLabel1.Text = "Ready"
                            ToolStripLabel1.Visible = True

                            cadX = ""

                            If ImportaExcelRecords(diLOg.FileName, allDs, cadX, ToolStripLabel1) = True Then

                                MsgBox("All tables and columns found correctly on the archive!!", vbInformation, TitBox)
                                'solo aqui deberia continuar al guardado!

                                For i = 0 To allDs.Tables.Count - 1
                                    ToolStripLabel1.Text = "Uploading data of table..." & allDs.Tables(i).TableName
                                    Application.DoEvents()
                                    xCad = ""
                                    xPath = RaizFire
                                    xPath = xPath & "/records/" & compaSelekted
                                    xPath = xPath & "/" & objetoSelek
                                    xPath = xPath & "/" & allDs.Tables(i).TableName
                                    xPath = xPath & "/" & "records"

                                    xCad = DtToJsonWithKey(allDs.Tables(i), "CampoYave")
                                    cadX = Await HazPostMasivoFbSingleJson(xPath, xCad)

                                Next

                                cadX = "Update complete!"
                                'MsgBox("Update complete!", vbInformation, TitBox)
                            Else
                                puedoAvanzar = False
                                cadX = "ATTENTION!!" & vbCrLf & "Some tables/sheets or columns/fields were not found on the archive you provided, because of that import was NOT possible!!, please refer below detailed report." & vbCrLf & cadX
                                'MsgBox("ATTENTION!!" & vbCrLf & "Some tables/sheets or columns/fields were not found on the archive you provided, please refer below detailed report." & vbCrLf & cadX, vbExclamation, TitBox)
                            End If

                            'Y aqui falta desbloquear masivo!!
                            For i = 0 To allDs.Tables.Count - 1

                                xPath = RaizFire
                                xPath = xPath & "/" & "inuse"
                                xPath = xPath & "/" & "records"
                                xPath = xPath & "/" & compaSelekted
                                xPath = xPath & "/" & objetoSelek
                                xPath = xPath & "/" & allDs.Tables(i).TableName

                                Await HazDeleteEnFbSimple(xPath, "")

                            Next

                            If puedoAvanzar = True Then
                                MsgBox(cadX, vbInformation, TitBox)
                            Else
                                MsgBox(cadX, vbExclamation, TitBox)
                            End If

                            ToolStripLabel1.Text = "Ready"
                            ToolStripLabel1.Visible = False

                            'aqui dependiendo de lo que se busque sería la onda de crear el
                            'procedimiento de forma masiva para postear toda la info!!

                            'O si no un visor de state completo?,
                            'Pero que onda con los bloqueos de hojas?, que pasa si alguien mas está editando el objeto!?, tendría que
                            'bloquearse, subirse y posteriormente liberarse!!
                            'si Solo es para simulacion, 
                        Else

                            For i = 0 To allDs.Tables.Count - 1

                                xPath = RaizFire
                                xPath = xPath & "/" & "inuse"
                                xPath = xPath & "/" & "records"
                                xPath = xPath & "/" & compaSelekted
                                xPath = xPath & "/" & objetoSelek
                                xPath = xPath & "/" & allDs.Tables(i).TableName

                                Await HazPost1Set(xPath, inUseDt, -1)

                            Next

                        End If


                    Case Is = 4
                        'importar una tabla de un objeto!
                        enCuentra = ModuPermit.Tables(0).Rows.Find(moduloSelek.ToUpper())
                        If IsNothing(enCuentra) = True Then
                            MsgBox("Sorry you are not allowed to import record to the selected module!!", vbCritical, TitBox)
                            Exit Sub
                        End If

                        If puedoEditar = False Then
                            MsgBox("Sorry, you are not allowed to edit this object at this time, please try again later or wait for it to be available to edit.", vbExclamation, TitBox)
                            Exit Sub
                        End If

                        Dim diLOg As New OpenFileDialog
                        Dim oFileName As String
                        oFileName = ""
                        diLOg.InitialDirectory = Environment.SpecialFolder.MyDocuments
                        diLOg.Filter = "Excel CSV File (*.csv)|*.csv" '"Excel File (*.xls)|*.xls"

                        'al importar hay varias opciones
                        '1. Importar una hoja simple
                        '2. Importar todo el template!
                        '3. Darle en importar y guardar
                        '4. Exportar una hoja simple
                        '5. Exportar todo el template!

                        Dim borroRecords As Boolean = False

                        If diLOg.ShowDialog() = Windows.Forms.DialogResult.OK Then

                            Dim path As String = diLOg.FileName

                            Dim elDt As New DataTable

                            elDt = LoadFromCSV(path)

                            If elDt.Rows.Count = 0 Then Exit Sub

                            DataGridView1.DataSource = Nothing
                            If borroRecords = True Then
                                DataGridView1.Rows.Clear()
                            End If

                            estoyAgregandoRows = True

                            DataGridView1.AllowUserToAddRows = False
                            DataGridView1.AllowUserToDeleteRows = False

                            Dim bS As New BindingSource

                            bS.DataSource = elDt

                            DataGridView1.AutoGenerateColumns = False

                            DataGridView1.DataSource = bS

                            For i = 0 To DataGridView1.Rows.Count - 1
                                DataGridView1.Rows(i).HeaderCell.Value = CStr(i + 1)
                            Next

                            estoyAgregandoRows = False

                            DataGridView1.AllowUserToAddRows = puedoEditar ' True
                            DataGridView1.AllowUserToDeleteRows = puedoEditar ' True

                        End If

                End Select





        End Select

        'Aqui se debe hacer la variante por cada nodo,
        'Si selecciono un nodo tabla continua de forma unitaria
        'Si selecciona un nodo objeto importa un excel con interop
        'Solo hasta ese nivel
        'Importar y guardar ó importar y solo ver?!?
        'Activar view mode, solo para observar y evaluar, y al final guardo por tabla o por objeto
        'Save y Save all





    End Sub

    Private Async Sub ToolStripButton10_Click(sender As Object, e As EventArgs) Handles ToolStripButton10.Click

        If CategSelected <> 3 Then Exit Sub
        'subir info a firebase!

        If IsNothing(elNode) = True Then
            MsgBox("Please select a valid node!", vbCritical, TitBox)
            Exit Sub
        End If

        Dim xObj As Object
        xObj = Split(elNode.FullPath, "\")

        If xObj.Length < 4 Then
            MsgBox("Please select a node at a table level!", vbCritical, TitBox)
            Exit Sub
        End If

        Dim yObj As Object
        Dim objetoCode As String
        objetoCode = elNode.Parent.Name

        For i = 0 To tempDs.Tables.Count - 1
            yObj = Split(tempDs.Tables(i).TableName, "#")
            If yObj(0) = objetoCode Then
                Dim enCuentra As DataRow
                enCuentra = ModuPermit.Tables(0).Rows.Find(CStr(yObj(2)).ToUpper())
                If IsNothing(enCuentra) = True Then
                    MsgBox("Sorry you are not allowed to make changes on the selected template", vbCritical, TitBox)
                    Exit Sub
                End If
                Exit For
            End If
        Next


        If puedoEditar = False Then
            MsgBox("Sorry, you are not allowed to edit this template at this time, please try again later or wait for it to become available!", vbExclamation, TitBox)
            Exit Sub
        End If

        Await SigoVivo()
        Module5.AgregaTiempo()

        'se debe tomar la info del grid!
        'writeDs.Tables(0).Rows.Clear()
        'writeDs.Tables(0).Rows.Add()
        Dim k As Integer = 0
        Dim cadYave As String = ""
        Dim MutaDs As New DataSet 'este dataset va a tomar la info del datagrid, de acuerdo a sus columnas!
        MutaDs.Tables.Clear()
        MutaDs.Tables.Add("Up")
        MutaDs.Tables(0).Columns.Add("CampoYave") 'concatenacion de los campos llave!
        For i = 0 To DataGridView1.Columns.Count - 4
            MutaDs.Tables(0).Columns.Add(DataGridView1.Columns(i).Name) 'codigo del campo!
            MutaDs.Tables(0).Columns(MutaDs.Tables(0).Columns.Count - 1).ExtendedProperties.Add("isKey", DataGridView1.Columns(i).Tag)

            If DataGridView1.Columns(i).Tag = "X" Then
                'es yave!
                k = k + 1

                If k = 1 Then
                    cadYave = i
                Else
                    cadYave = cadYave & "#" & i
                End If

            End If

        Next

        'primero deberíamos limpiar el grid no?
        'quitar blancos, y campos llave
        Dim xResp As String = ""
        ToolStripLabel1.Visible = True

        toyTrabajando = True

        If cadYave <> "" Then
            'SI hay campos llave!
            ToolStripLabel1.Text = "Finding duplicates..."

            MsgBox(xResp, vbInformation, TitBox)

        End If

        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AllowUserToDeleteRows = False

        Dim yavEField As String = ""

        'ToolStripLabel1.Text = "Getting data..."
        ToolStripLabel1.Text = "Fetching data..."
        Application.DoEvents()

        For i = 0 To DataGridView1.Rows.Count - 1
            k = 0
            yavEField = ""

            MutaDs.Tables(0).Rows.Add()
            For j = 0 To DataGridView1.Columns.Count - 4
                MutaDs.Tables(0).Rows(MutaDs.Tables(0).Rows.Count - 1).Item(j + 1) = DataGridView1.Rows(i).Cells(j).Value
                'quitar las comillas
                MutaDs.Tables(0).Rows(MutaDs.Tables(0).Rows.Count - 1).Item(j + 1) = Replace(MutaDs.Tables(0).Rows(MutaDs.Tables(0).Rows.Count - 1).Item(j + 1), """", "")



            Next

            yavEField = Guid.NewGuid().ToString()
            MutaDs.Tables(0).Rows(MutaDs.Tables(0).Rows.Count - 1).Item(0) = yavEField
        Next

        'OJO, 3 pasitos antes:
        'compania y su nombre
        'objeto y su nombre
        'tabla y su nombre!

        Dim cam1 As String = RaizFire
        cam1 = cam1 & "/records/" & elNode.Parent.Parent.Name
        Await HazPutEnFbSimple(cam1, "CompanyName", elNode.Parent.Parent.Tag)

        cam1 = RaizFire
        cam1 = cam1 & "/records/" & elNode.Parent.Parent.Name
        cam1 = cam1 & "/" & elNode.Parent.Name 'objeto
        Await HazPutEnFbSimple(cam1, "ObjectName", elNode.Parent.Tag)

        cam1 = RaizFire
        cam1 = cam1 & "/records/" & elNode.Parent.Parent.Name
        cam1 = cam1 & "/" & elNode.Parent.Name 'objeto
        cam1 = cam1 & "/" & elNode.Name 'la tabla
        Await HazPutEnFbSimple(cam1, "TableName", elNode.Tag)


        If MutaDs.Tables(0).Rows.Count = 0 Then

            Dim x As Integer = 0
            x = MsgBox("WARNING! " & vbCrLf & "There are zero records in the table, you want to delete the whole table?", vbQuestion + vbYesNo, "Delete records?")

            If x <> 6 Then
                DataGridView1.AllowUserToAddRows = puedoEditar ' True
                DataGridView1.AllowUserToDeleteRows = puedoEditar ' True
                toyTrabajando = False
                Exit Sub
            End If

        End If

        'eliminar el nodo completo en caso de que exista!!
        cam1 = RaizFire
        cam1 = cam1 & "/records/" & elNode.Parent.Parent.Name
        cam1 = cam1 & "/" & elNode.Parent.Name 'objeto
        cam1 = cam1 & "/" & elNode.Name 'la tabla
        Await HazDeleteEnFbSimple(cam1, "records")

        Dim elCamino As String = RaizFire
        elCamino = elCamino & "/records/" & elNode.Parent.Parent.Name 'compania
        elCamino = elCamino & "/" & elNode.Parent.Name 'objeto
        elCamino = elCamino & "/" & elNode.Name 'la tabla
        elCamino = elCamino & "/" & "records"

        ToolStripLabel1.Text = "Uploading..."

        ToolStripButton10.Enabled = False
        ToolStripButton11.Enabled = False

        If MutaDs.Tables(0).Rows.Count > 0 Then


        Else
            xResp = "ok"
        End If

        toyTrabajando = False

        ToolStripLabel1.Visible = False
        ToolStripLabel1.Text = "Ready"

        ToolStripButton10.Enabled = True
        ToolStripButton11.Enabled = True

        If xResp = "ok" Then
            MsgBox("Update complete!", vbInformation, TitBox)
        Else
            MsgBox("Failure on upload!", vbInformation, TitBox)
        End If

        DataGridView1.AllowUserToAddRows = puedoEditar ' True
        DataGridView1.AllowUserToDeleteRows = puedoEditar ' True

    End Sub

    Private Async Sub ToolStripButton12_Click(sender As Object, e As EventArgs) Handles ToolStripButton12.Click


        'Exportar registros
        'esperar a mañana a la presentacion!
        'Igual, va a ser la variante dependiendo del nodo seleccionado
        'Si selecciona un nodo tabla continua con el unitario
        'Si selecciona el nodo objeto, exporta todo lo de ese objeto
        'Si tiene seleccionado el nodo compañía, exporta la compañía completa con todos sus hijos


        Dim xObj As Object
        xObj = Split(elNode.FullPath, "\")

        Dim lCol As Integer = 1

        Select Case ToolStripComboBox1.SelectedIndex

            Case Is = 1 'catalogos
                'los catalogos también podría ser una onda de exportar todos los catalogos de un nodo!
                Select Case xObj.Length
                    Case Is = 1
                        'selecciono un nodo de un catalogo!


                    Case Is = 2
                        'seleccionó una tabla!
                        If DataGridView1.Rows.Count - 1 < 0 Then
                            MsgBox("Deploy some information first!!", vbCritical, TitBox)
                            Exit Sub
                        End If

                        Dim diLOg As New SaveFileDialog
                        Dim oFileName As String
                        oFileName = ""
                        diLOg.InitialDirectory = Environment.SpecialFolder.MyDocuments
                        diLOg.Filter = "Excel CSV File (*.csv)|*.csv" '"Excel File (*.xls)|*.xls"
                        'diLOg.Filter = "Excel File (*.xls)|*.xls|Todos los archivos (*.*)|*.*"
                        If diLOg.ShowDialog = System.Windows.Forms.DialogResult.OK Then

                            Call ExportToCsv2(DataGridView1, diLOg.FileName, lCol)

                        End If

                End Select



            Case Is = 2



            Case Is = 3 'registros
                'todos las hojas y objetos de un nodo!
                Select Case xObj.Length
                    Case Is = 1
                        'seleccionó all companies



                    Case Is = 2
                        'seleccionó 1 compañia



                    Case Is = 3
                        'seleccionó un objeto
                        'sería hacer un ciclo de esa companía, de ese objeto!
                        If PussyTemp < 0 Then
                            MsgBox("No template found!", vbExclamation, TitBox)
                            Exit Sub
                        End If


                        If elNode.Nodes.Count = 0 Then
                            MsgBox("This object does not contain any tables!!, choose another one!", vbCritical, TitBox)
                            Exit Sub
                        End If

                        Dim filtrosDt As DataTable = tempDs.Tables(PussyTemp).Clone()

                        Dim diLOg As New SaveFileDialog

                        diLOg.InitialDirectory = Environment.SpecialFolder.MyDocuments
                        diLOg.Filter = "Excel File (*.xlsx)|*.xlsx" '"Excel File (*.xls)|*.xls"

                        If diLOg.ShowDialog = System.Windows.Forms.DialogResult.OK Then

                            Dim allDs As New DataSet
                            Dim xDs As New DataSet
                            Dim elCamino As String = "" ' RaizFire
                            Dim xError As String = ""
                            allDs.ExtendedProperties.Add("TemplateCode", elNode.Name)
                            allDs.ExtendedProperties.Add("TemplateName", elNode.Text)

                            allDs.ExtendedProperties.Add("CompanyCode", elNode.Parent.Name)
                            allDs.ExtendedProperties.Add("CompanyName", elNode.Parent.Text)
                            allDs.ExtendedProperties.Add("Module", moduloSelek)

                            For i = 0 To elNode.Nodes.Count - 1
                                elCamino = RaizFire
                                elCamino = elCamino & "/records/" & elNode.Parent.Name 'compania
                                elCamino = elCamino & "/" & elNode.Name 'objeto
                                elCamino = elCamino & "/" & elNode.Nodes(i).Name

                                elCamino = elCamino & "/records"

                                usaDataset.Tables(0).Rows.Clear()
                                usaDataset.Tables(0).Rows.Add({"records"})
                                usaDataset.Tables(0).Rows.Add({elNode.Parent.Name}) 'compania
                                usaDataset.Tables(0).Rows.Add({elNode.Name}) 'objeto
                                usaDataset.Tables(0).Rows.Add({elNode.Nodes(i).Name}) 'tabla
                                usaDataset.Tables(0).Rows.Add({"records"})

                                'xdepeds
                                xDs.Tables.Clear()
                                xDs = Await PullUrlWs(usaDataset, "records")

                                If xDs.Tables.Count = 0 Then

                                    filtrosDt.Rows.Clear()
                                    Dim result() As DataRow = tempDs.Tables(PussyTemp).Select("TableCode = '" & elNode.Nodes(i).Name & "'", "Position ASC")
                                    For Each row As DataRow In result
                                        filtrosDt.ImportRow(row)
                                    Next

                                    allDs.Tables.Add(elNode.Nodes(i).Name)
                                    allDs.Tables(allDs.Tables.Count - 1).ExtendedProperties.Add("TableCode", elNode.Nodes(i).Name)
                                    allDs.Tables(allDs.Tables.Count - 1).ExtendedProperties.Add("TableName", "")

                                    For j = 0 To filtrosDt.Rows.Count - 1
                                        allDs.Tables(allDs.Tables.Count - 1).Columns.Add(filtrosDt.Rows(j).Item("FieldCode"))
                                        allDs.Tables(allDs.Tables.Count - 1).Columns(allDs.Tables(allDs.Tables.Count - 1).Columns.Count - 1).ExtendedProperties.Add("FieldCode", filtrosDt.Rows(j).Item("FieldCode"))
                                        allDs.Tables(allDs.Tables.Count - 1).Columns(allDs.Tables(allDs.Tables.Count - 1).Columns.Count - 1).ExtendedProperties.Add("HeaderText", filtrosDt.Rows(j).Item("FieldName"))
                                        allDs.Tables(allDs.Tables.Count - 1).Columns(allDs.Tables(allDs.Tables.Count - 1).Columns.Count - 1).ExtendedProperties.Add("MOC", filtrosDt.Rows(j).Item("MOC"))
                                        allDs.Tables(allDs.Tables.Count - 1).ExtendedProperties.Item("TableName") = filtrosDt.Rows(j).Item("TableName")
                                    Next

                                    allDs.Tables(allDs.Tables.Count - 1).ExtendedProperties.Add("FromCol", 0)
                                    allDs.Tables(allDs.Tables.Count - 1).ExtendedProperties.Add("ToCol", allDs.Tables(allDs.Tables.Count - 1).Columns.Count - 1)

                                Else

                                    allDs.Tables.Add(xDs.Tables(0).Copy())
                                    allDs.Tables(allDs.Tables.Count - 1).TableName = elNode.Nodes(i).Name
                                    allDs.Tables(allDs.Tables.Count - 1).ExtendedProperties.Add("TableCode", elNode.Nodes(i).Name)
                                    allDs.Tables(allDs.Tables.Count - 1).ExtendedProperties.Add("TableName", "")
                                    allDs.Tables(allDs.Tables.Count - 1).ExtendedProperties.Add("FromCol", 1)
                                    allDs.Tables(allDs.Tables.Count - 1).ExtendedProperties.Add("ToCol", allDs.Tables(allDs.Tables.Count - 1).Columns.Count - 1)

                                    For j = 0 To allDs.Tables(allDs.Tables.Count - 1).Columns.Count - 1

                                        allDs.Tables(allDs.Tables.Count - 1).Columns(j).ExtendedProperties.Add("FieldCode", "")
                                        allDs.Tables(allDs.Tables.Count - 1).Columns(j).ExtendedProperties.Add("HeaderText", "")
                                        allDs.Tables(allDs.Tables.Count - 1).Columns(j).ExtendedProperties.Add("MOC", "")

                                        filtrosDt.Rows.Clear()
                                        Dim result() As DataRow = tempDs.Tables(PussyTemp).Select("TableCode = '" & elNode.Nodes(i).Name & "' AND FieldCode='" & allDs.Tables(allDs.Tables.Count - 1).Columns(j).ColumnName & "'")
                                        For Each row As DataRow In result
                                            filtrosDt.ImportRow(row)
                                        Next

                                        If filtrosDt.Rows.Count = 0 Then Continue For

                                        allDs.Tables(allDs.Tables.Count - 1).ExtendedProperties.Item("TableName") = filtrosDt.Rows(0).Item("TableName")
                                        allDs.Tables(allDs.Tables.Count - 1).Columns(j).ExtendedProperties.Item("FieldCode") = filtrosDt.Rows(0).Item("FieldCode")
                                        allDs.Tables(allDs.Tables.Count - 1).Columns(j).ExtendedProperties.Item("HeaderText") = filtrosDt.Rows(0).Item("FieldName")
                                        allDs.Tables(allDs.Tables.Count - 1).Columns(j).ExtendedProperties.Item("MOC") = filtrosDt.Rows(0).Item("MOC")

                                    Next

                                End If

                            Next

                            Dim xStat As String = ""
                            ToolStripLabel1.Visible = True
                            ToolStripLabel1.Text = "Ready"
                            If ExportaRecordsReport(allDs, diLOg.FileName, ToolStripLabel1, xError) = True Then

                                Dim X As Integer
                                X = MsgBox("Report exported at: " & diLOg.FileName & " !" & vbCrLf & "You want to open the file now?", vbQuestion + vbYesNo, TitBox)
                                If X = 6 Then
                                    System.Diagnostics.Process.Start(diLOg.FileName)
                                End If
                            Else
                                MsgBox("Sorry there was an error loading excel", vbExclamation, TitBox)
                            End If

                            ToolStripLabel1.Text = "Ready"
                            ToolStripLabel1.Visible = False


                        End If



                    Case Is = 4
                        'seleccionó una hoja
                        If DataGridView1.Rows.Count - 1 < 0 Then
                            MsgBox("Deploy some information first!!", vbCritical, TitBox)
                            Exit Sub
                        End If

                        Dim diLOg As New SaveFileDialog
                        Dim oFileName As String
                        oFileName = ""
                        diLOg.InitialDirectory = Environment.SpecialFolder.MyDocuments
                        diLOg.Filter = "Excel CSV File (*.csv)|*.csv" '"Excel File (*.xls)|*.xls"
                        'diLOg.Filter = "Excel File (*.xls)|*.xls|Todos los archivos (*.*)|*.*"
                        If diLOg.ShowDialog = System.Windows.Forms.DialogResult.OK Then

                            Call ExportToCsv2(DataGridView1, diLOg.FileName, lCol)

                        End If

                End Select


            Case Is = 4 'templates

                Select Case xObj.Length
                    Case Is = 1
                        'seleccionó All templates


                    Case Is = 2
                        'selecciono 1 template


                    Case Is = 3, 4
                        'selecciono una hoja!
                        If DataGridView1.Rows.Count - 1 < 0 Then
                            MsgBox("Deploy some information first!!", vbCritical, TitBox)
                            Exit Sub
                        End If

                        Dim diLOg As New SaveFileDialog
                        Dim oFileName As String
                        oFileName = ""
                        diLOg.InitialDirectory = Environment.SpecialFolder.MyDocuments
                        diLOg.Filter = "Excel CSV File (*.csv)|*.csv" '"Excel File (*.xls)|*.xls"
                        'diLOg.Filter = "Excel File (*.xls)|*.xls|Todos los archivos (*.*)|*.*"
                        If diLOg.ShowDialog = System.Windows.Forms.DialogResult.OK Then

                            Call ExportToCsv2(DataGridView1, diLOg.FileName, lCol)

                        End If

                End Select


        End Select





    End Sub

    Private Sub DataGridView1_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles DataGridView1.CurrentCellDirtyStateChanged
        If DataGridView1.IsCurrentCellDirty Then
            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Async Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click

        Dim xObj As Object
        Dim yObj As Object
        Dim pathBuild As String = ""
        Dim elRegreso As String = ""
        Dim cadeFb As String = ""
        Dim tabValor As String = ""
        Dim objeName As String = ""
        Dim tablaName As String = ""
        Dim moduCode As String = ""
        Dim camiOpcion As String = ""

        Select Case CategSelected
            Case Is = 0
                MsgBox("Please select a category first!!", vbCritical, TitBox)

            Case Is = 1

                xObj = Split(elNode.FullPath, "\")

                Dim enCuentra As DataRow
                enCuentra = ModuPermit.Tables(0).Rows.Find(CStr(xObj(0)).ToUpper())
                If IsNothing(enCuentra) = True Then
                    MsgBox("Sorry you are not allowed to add catalogs on the selected module", vbCritical, TitBox)
                    Exit Sub
                End If

                If puedoEditar = False Then
                    MsgBox("Sorry you are not allowed to make changes to the selected catalog at this time, please try again later or wait for it to become available.", vbExclamation, TitBox)
                    Exit Sub
                End If

                pathBuild = "catalogs"
                pathBuild = pathBuild & "/" & xObj(0)
                pathBuild = pathBuild & "/" & elNode.Name

                cadeFb = RaizFire
                cadeFb = cadeFb & "/catpro/" & xObj(0)
                cadeFb = cadeFb & "/" & elNode.Name

                Form2.huboExito = False
                Form2.queOpcion = 6
                Form2.elTitulo = "Edit catalog"
                Form2.keyName = "Table code:"
                Form2.tabName = "Catalog name:"
                Form2.keyValue = elNode.Name 'codigo de la tabla!!
                Form2.tabValue = elNode.Text
                Form2.pathLabel = "Path:"
                Form2.elCamino = pathBuild
                Form2.ShowDialog()

                If Form2.huboExito = False Then Exit Sub

                elRegreso = Await HazPutEnFbSimple(cadeFb, "CatalogName", Form2.tabValue)

                If elRegreso = "Ok" Then
                    MsgBox("Information updated!!", vbInformation, TitBox)
                    elNode.Text = Form2.tabValue
                    Await CargaOpcion(1)
                End If


            Case Is = 4
                'templates!
                If IsNothing(elNode) = True Then
                    'error!!
                    MsgBox("Please select a node to add the template on the corresponding structure!", vbCritical, TitBox)
                    Exit Sub
                End If

                xObj = Split(elNode.FullPath, "\")

                If xObj.Length <= 1 Or xObj.Length > 3 Then
                    MsgBox("Please select a template node or a table node!!", vbCritical, TitBox)
                    Exit Sub
                End If

                Select Case xObj.Length
                    Case Is = 2
                        yObj = Split(elNode.Tag, "#")

                    Case Else '3
                        yObj = Split(elNode.Parent.Tag, "#")

                End Select

                objeName = yObj(0)
                moduCode = yObj(1)

                Dim enCuentra As DataRow
                enCuentra = ModuPermit.Tables(0).Rows.Find(CStr(yObj(1)).ToUpper())
                If IsNothing(enCuentra) = True Then
                    MsgBox("Sorry you are not allowed to make changes on this template", vbCritical, TitBox)
                    Exit Sub
                End If

                If puedoEditar = False Then
                    MsgBox("Sorry you are not allowed to make changes to the selected catalog at this time, please try again later or wait for it to become available.", vbExclamation, TitBox)
                    Exit Sub
                End If

                pathBuild = "templates"

                cadeFb = RaizFire
                cadeFb = cadeFb & "/templates" ' & xObj(0)

                Select Case xObj.Length
                    Case Is = 2
                        'estoy en un template
                        pathBuild = pathBuild & "/" & elNode.Name
                        cadeFb = cadeFb & "/" & elNode.Name
                        Form2.elTitulo = "Edit template name"
                        Form2.keyName = "Template code:"
                        Form2.tabName = "Template name:"
                        Form2.keyValue = elNode.Name 'codigo de la tabla!!
                        Form2.tabValue = objeName
                        camiOpcion = "ObjectName"


                    Case Is = 3
                        'estoy en una tabla!
                        pathBuild = pathBuild & "/" & elNode.Parent.Name
                        pathBuild = pathBuild & "/" & elNode.Name
                        cadeFb = cadeFb & "/" & elNode.Parent.Name
                        cadeFb = cadeFb & "/" & elNode.Name
                        Form2.elTitulo = "Edit table name"
                        Form2.keyName = "Table code:"
                        Form2.tabName = "Table name:"
                        Form2.keyValue = elNode.Name 'codigo de la tabla!!
                        Form2.tabValue = elNode.Tag
                        camiOpcion = "TableName"

                End Select


                Form2.huboExito = False
                Form2.queOpcion = 6
                'Form2.elTitulo = "Edit catalog"
                'Form2.keyName = "Table code:"
                'Form2.tabName = "Catalog name:"
                'Form2.keyValue = elNode.Name 'codigo de la tabla!!
                'Form2.tabValue = elNode.Text

                Form2.pathLabel = "Path:"
                Form2.elCamino = pathBuild
                Form2.ShowDialog()

                If Form2.huboExito = False Then Exit Sub

                elRegreso = Await HazPutEnFbSimple(cadeFb, camiOpcion, Form2.tabValue)

                If elRegreso = "Ok" Then
                    MsgBox("Information updated!!", vbInformation, TitBox)
                    elNode.Text = elNode.Name & " / " & Form2.tabValue
                    Select Case xObj.Length
                        Case Is = 2
                            elNode.Tag = Form2.tabValue & "#" & moduCode

                        Case Is = 3
                            elNode.Tag = Form2.tabValue

                    End Select

                    Await CargaOpcion(4)

                End If

        End Select


    End Sub

    Private Async Sub ToolStripButton7_Click(sender As Object, e As EventArgs) Handles ToolStripButton7.Click

        'Tuuu!!
        'procedimiento para cambiar todo masivamente!!
        'Dim unCamino As String ' = RaizFire
        'Dim xObj As Object
        'Dim xDs As New DataSet
        'xDs.Tables.Add()
        'xDs.Tables(0).Columns.Add("Key")
        'xDs.Tables(0).Columns.Add("Description")

        'For i = 0 To catDs.Tables.Count - 1
        '    unCamino = RaizFire
        '    xObj = Split(catDs.Tables(i).TableName, "#")
        '    'primero borramos!

        '    unCamino = unCamino & "/catalogs/" & CStr(xObj(0))

        '    Await HazDeleteEnFbSimple(unCamino, CStr(xObj(1)))

        '    'primero se escribe el nombre del catalogo:

        '    unCamino = RaizFire
        '    unCamino = unCamino & "/catalogs/" & CStr(xObj(0)) & "/" & CStr(xObj(1))

        '    Await HazPutEnFbSimple(unCamino, "CatalogName", catDs.Tables(i).Columns(0).ColumnName)

        '    xDs.Tables(0).Rows.Clear()

        '    For j = 0 To catDs.Tables(i).Rows.Count - 1

        '        xDs.Tables(0).Rows.Add({catDs.Tables(i).Rows(j).Item(0), catDs.Tables(i).Rows(j).Item(1)})

        '    Next

        '    unCamino = RaizFire
        '    unCamino = unCamino & "/catalogs/" & CStr(xObj(0))

        '    Await HazPostEnFireBaseConPathYColumnas(unCamino, xDs.Tables(0), CStr(xObj(1)), -1)

        'Next

    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged

        If estoyAgregandoRows = True Then Exit Sub

        Dim cadeX As String = ""
        Dim xObj As Object = Nothing
        Dim enCuentra As DataRow
        Dim z As Long
        Select Case ToolStripComboBox1.SelectedIndex
            Case Is = 0

            Case Is = 1

            Case Is = 2


            Case Is = 3


            Case Is = 4
                'templates!
                If IsNothing(elNode) = True Then Exit Sub

                Dim comEntry As DataGridViewComboBoxCell
                comEntry = DataGridView1.Rows(e.RowIndex).Cells(3)

                Dim comFill As DataGridViewComboBoxCell
                comFill = DataGridView1.Rows(e.RowIndex).Cells(4)

                Dim comRep As DataGridViewComboBoxCell
                comRep = DataGridView1.Rows(e.RowIndex).Cells(12)

                Select Case e.ColumnIndex
                    Case Is = 2
                        'hizo check al Key Field

                        If DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = True Then
                            'se pone como mandatorio el combo de la derecha
                            comEntry.Value = "Mandatory"
                            comEntry.ReadOnly = True

                            comRep.Value = "N/A"
                            comRep.ReadOnly = True

                        Else
                            comEntry.ReadOnly = False

                            comRep.Value = "No"
                            comRep.ReadOnly = False

                        End If

                    Case Is = 3
                        'MOC
                        'If DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = "Conditional" Then
                        '    'se habilita rule y value de conditional
                        '    DataGridView1.Rows(e.RowIndex).Cells(19).ReadOnly = False
                        '    DataGridView1.Rows(e.RowIndex).Cells(20).ReadOnly = False
                        'Else
                        '    'se deshabilita!
                        '    DataGridView1.Rows(e.RowIndex).Cells(19).ReadOnly = True
                        '    DataGridView1.Rows(e.RowIndex).Cells(20).ReadOnly = True
                        'End If

                    Case Is = 4
                        Select Case DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                            Case Is = "No selection"
                                DataGridView1.Rows(e.RowIndex).Cells(9).Value = ""
                                DataGridView1.Rows(e.RowIndex).Cells(9).ReadOnly = True

                                DataGridView1.Rows(e.RowIndex).Cells(11).Value = ""
                                DataGridView1.Rows(e.RowIndex).Cells(11).ReadOnly = True

                            Case Is = "A - From Catalog"
                                'se habilita catalogs

                                DataGridView1.Rows(e.RowIndex).Cells(11).Value = ""
                                DataGridView1.Rows(e.RowIndex).Cells(11).ReadOnly = True

                            Case Is = "B - Construction"
                                DataGridView1.Rows(e.RowIndex).Cells(9).Value = ""
                                DataGridView1.Rows(e.RowIndex).Cells(9).ReadOnly = True

                                DataGridView1.Rows(e.RowIndex).Cells(11).Value = ""
                                DataGridView1.Rows(e.RowIndex).Cells(11).ReadOnly = True

                            Case Is = "D - Fixed Value"
                                DataGridView1.Rows(e.RowIndex).Cells(9).Value = ""
                                DataGridView1.Rows(e.RowIndex).Cells(9).ReadOnly = True

                                DataGridView1.Rows(e.RowIndex).Cells(11).ReadOnly = False'valor fijo

                            Case Is = "E - Simple validation"
                                DataGridView1.Rows(e.RowIndex).Cells(9).Value = ""
                                DataGridView1.Rows(e.RowIndex).Cells(9).ReadOnly = True

                                DataGridView1.Rows(e.RowIndex).Cells(11).Value = ""
                                DataGridView1.Rows(e.RowIndex).Cells(11).ReadOnly = True

                        End Select


                    Case Is = 5
                        'DataType
                        Select Case DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                            Case Is = "No selection"
                                DataGridView1.Rows(e.RowIndex).Cells(6).Value = ""
                                DataGridView1.Rows(e.RowIndex).Cells(6).ReadOnly = True

                            Case Is = "Number"
                                DataGridView1.Rows(e.RowIndex).Cells(6).Value = "5"
                                DataGridView1.Rows(e.RowIndex).Cells(6).ReadOnly = False

                            Case Is = "Date" 'en este solo se pude 8 o 10
                                DataGridView1.Rows(e.RowIndex).Cells(6).Value = "8"
                                DataGridView1.Rows(e.RowIndex).Cells(6).ReadOnly = False

                            Case Is = "Text"
                                DataGridView1.Rows(e.RowIndex).Cells(6).Value = "10"
                                DataGridView1.Rows(e.RowIndex).Cells(6).ReadOnly = False

                            Case Is = "Email"
                                DataGridView1.Rows(e.RowIndex).Cells(6).Value = "100"
                                DataGridView1.Rows(e.RowIndex).Cells(6).ReadOnly = False

                            Case Is = "Decimal"
                                DataGridView1.Rows(e.RowIndex).Cells(6).Value = "15.2"
                                DataGridView1.Rows(e.RowIndex).Cells(6).ReadOnly = False

                            Case Is = "Indicator"
                                DataGridView1.Rows(e.RowIndex).Cells(6).Value = "1"
                                DataGridView1.Rows(e.RowIndex).Cells(6).ReadOnly = True 'siempre es 1!!

                            Case Is = "Time" 'aqui solo se puede 6 u 8
                                DataGridView1.Rows(e.RowIndex).Cells(6).Value = "6"
                                DataGridView1.Rows(e.RowIndex).Cells(6).ReadOnly = False '

                        End Select


                    Case Is = 6
                        'Maxs Chars
                        If DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = "" Then Exit Sub

                        If DataGridView1.Rows(e.RowIndex).Cells(5).Value = "No selection" Then
                            MsgBox("Please select first the Data Type!!", vbCritical, TitBox)
                            DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = ""
                            Exit Sub
                        End If

                        If IsNumeric(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) = False Then
                            MsgBox("Only numbers allowed!!!", vbCritical, TitBox)
                            DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = ""
                            Exit Sub
                        End If

                        cadeX = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value

                        Select Case DataGridView1.Rows(e.RowIndex).Cells(5).Value
                            Case Is = "Number"
                                If cadeX.Contains(".") = True Then
                                    MsgBox("Only integers allowed!!!", vbCritical, TitBox)
                                    DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = Math.Truncate(CDec(cadeX))
                                    Exit Sub
                                End If

                            Case Is = "Date"
                                If cadeX.Contains(".") = True Then
                                    MsgBox("Only integers allowed!!!", vbCritical, TitBox)
                                    DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = ""
                                    Exit Sub
                                End If

                                If cadeX = "8" Or cadeX = "10" Then
                                    'ok!
                                Else
                                    MsgBox("It's only allowed 8 or 10 characters allowed for Date format, YYYYMMDD or YYYY.MM.DD", vbCritical, TitBox)
                                    DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = "8"
                                End If

                            Case Is = "Text"
                                If cadeX = "" Then
                                    MsgBox("You should type at least 1 character length for Text field", vbCritical, TitBox)
                                    DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = "1"
                                    Exit Sub
                                End If

                            Case Is = "Email"
                                If cadeX = "" Then
                                    MsgBox("You should type at least 50 character length for Email field", vbCritical, TitBox)
                                    DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = "50"
                                    Exit Sub
                                End If

                            Case Is = "Decimal"
                                If cadeX.Contains(".") = False Then
                                    MsgBox("This is a decimal field, you should put a integer and decimal part for the number in this format> Example: '10.3'", vbCritical, TitBox)
                                    DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = "15.2"
                                    Exit Sub
                                End If

                            Case Is = "Indicator"
                                'de okis


                            Case Is = "Time"
                                If cadeX.Contains(".") = True Then
                                    MsgBox("Only integers allowed!!!", vbCritical, TitBox)
                                    DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = ""
                                    Exit Sub
                                End If

                                If cadeX = "6" Or cadeX = "8" Then
                                    'ok!
                                Else
                                    MsgBox("It's only allowed 6 or 8 characters allowed for Time format, HHMMSS or HH:MM:SS", vbCritical, TitBox)
                                    DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = "6"
                                End If

                        End Select


                    Case Is = 7
                        'U/LCase
                        'de okis


                    Case Is = 8
                        'Blanks
                        'de okis


                    Case Is = 9
                        'catalog Name
                        xObj = Split(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, " ")
                        ' lo buscamos en la
                        enCuentra = filtCat.Rows.Find(CStr(xObj(UBound(xObj))))
                        If IsNothing(enCuentra) = True Then
                            DataGridView1.Rows(e.RowIndex).Cells(10).Value = ""
                        Else
                            z = filtCat.Rows.IndexOf(enCuentra)
                            DataGridView1.Rows(e.RowIndex).Cells(10).Value = CStr(filtCat.Rows(z).Item(0))
                        End If

                    Case Is = 10
                        'de okis,catalog code


                    Case Is = 11
                        'valor fijo en caso de que elija Fixed Value en filling rule


                    Case Is = 12
                        'No Rep

                    Case Is = 13
                        'allowed chars


                    Case Is = 14, 15, 16, 17, 18
                        'Conditionals

                    Case Is = 19
                        'Conditional Rule


                    Case Is = 20
                        'Conditional Value


                    Case Is = 21
                        'Construction Rule


                End Select

        End Select

    End Sub

    Private Sub DataGridView1_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles DataGridView1.RowsAdded

        Select Case ToolStripComboBox1.SelectedIndex
            Case Is = 0

            Case Is = 1

            Case Is = 2

            Case Is = 3
                'records!

            Case Is = 4
                'templates

        End Select

        'DataGridView1.Rows(e.RowIndex).Cells(0)

    End Sub

    Private Sub DataGridView1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing

        Select Case ToolStripComboBox1.SelectedIndex
            Case Is = 0

            Case Is = 1

            Case Is = 2

            Case Is = 3

            Case Is = 4
                'templates
                Select Case DataGridView1.CurrentCell.ColumnIndex
                    Case Is = 9
                        'catalog Name
                        Dim autoText As TextBox = CType(e.Control, TextBox)
                        If (autoText IsNot Nothing) Then
                            autoText.AutoCompleteCustomSource = FuenteCatalogos
                            autoText.AutoCompleteMode = AutoCompleteMode.Suggest
                            autoText.AutoCompleteSource = AutoCompleteSource.CustomSource
                        End If

                    Case Is = 10
                        'otros


                End Select


        End Select
    End Sub

    Private Sub ReloadCatPro()

        CatSimple.Tables(0).Rows.Clear()
        'limpiamos y agregamos!
        Dim xObj As Object = Nothing
        'Dim enCuentra As DataRow
        'Dim z As Long
        For i = 0 To catDs.Tables.Count - 1
            xObj = Split(catDs.Tables(i).TableName, "#")

            CatSimple.Tables(0).Rows.Add({CStr(xObj(1)), CStr(xObj(0)), catDs.Tables(i).ExtendedProperties.Item("CatalogName")})

        Next

    End Sub

    Private Sub ReloadCatSimple()

        CatSimple.Tables(0).Rows.Clear()
        'limpiamos y agregamos!
        Dim xObj As Object = Nothing
        'Dim enCuentra As DataRow
        'Dim z As Long
        For i = 0 To catDs.Tables.Count - 1
            xObj = Split(catDs.Tables(i).TableName, "#")

            'enCuentra = catDs.Tables(i).Rows.Find("CatalogName")
            'If IsNothing(enCuentra) = True Then
            '    Continue For
            'End If

            'z = catDs.Tables(i).Rows.IndexOf(enCuentra)

            CatSimple.Tables(0).Rows.Add({CStr(xObj(1)), CStr(xObj(0)), catDs.Tables(i).Columns(0).ColumnName})

        Next

    End Sub

    Private Async Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick

        Dim pathBuild As String = ""
        Dim xDs As New DataSet
        xDs.Tables.Add()
        xDs.Tables(0).Columns.Add()
        xDs.Tables(0).Columns.Add()

        Select Case ToolStripComboBox1.SelectedIndex
            Case Is = 0

            Case Is = 1

            Case Is = 2


            Case Is = 3
                'records!!

                If e.RowIndex < 0 Then Exit Sub

                Dim elDeta As String = ""
                Select Case e.ColumnIndex
                    Case Is = DataGridView1.Columns.Count - 1
                        'detalles
                        elDeta = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Tag
                        MsgBox(elDeta, vbInformation, TitBox)

                    Case Else



                End Select



            Case Is = 4
                'templates!
                If puedoEditar = False Then
                    MsgBox("Sorry you are not allowed to make any changes to this template at this time. Please make sure you have proper access to this object and no other users are editing.", vbExclamation, TitBox)
                    Exit Sub
                End If

                If e.ColumnIndex >= 14 And e.ColumnIndex <= 21 Then
                    'condition table!
                    'hace falta un botón para eliminar la condicional anterior!!
                    'solo si selecciono Conditional se habilita!

                    If DataGridView1.Rows(e.RowIndex).Cells(3).Value <> "Conditional" Then
                        MsgBox("Please select 'Conditional' option in MOC column for set conditionals!", vbCritical, TitBox)
                        Exit Sub
                    End If

                    'Ahorita lo vamos a quitar, pero lo vamos a dejar cuando vayamos go live!!:
                    'pasa si esta vacío o es N/A
                    If DataGridView1.Rows(e.RowIndex).Cells(14).Value = "N/A" Or DataGridView1.Rows(e.RowIndex).Cells(14).Value = "" Then

                    Else
                        MsgBox("You have to delete first this dependencie!!", vbCritical, TitBox)
                        Exit Sub
                    End If

                    Form3.elEnfoque = "D" 'dependencia directa
                    Form3.xtraDs = tempDs
                    Form3.yTraDs = depeDs
                    Form3.depeTemplate = objetoSelek ' elNode.Parent.Name
                    Form3.depeTabla = tableSelek ' elNode.Name

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

                    Form3.depeCampo = DataGridView1.Rows(e.RowIndex).Cells(0).Value

                    'dep/objeto/tabla
                    Form3.elPath = "dependencies/" & objetoSelek & "/" & tableSelek

                    Form3.huboExito = False

                    Form3.ShowDialog()

                    If Form3.huboExito = False Then Exit Sub

                    'AQUII, hay que crear el object name siempre, antes de iniciar!!!
                    pathBuild = RaizFire & "/" & "dependencies"
                    pathBuild = pathBuild & "/" & Form3.depeTemplate

                    Await HazPutEnFbSimple(pathBuild, "ObjectName", objetoName)

                    pathBuild = RaizFire & "/" & "dependencies"
                    pathBuild = pathBuild & "/" & Form3.depeTemplate
                    pathBuild = pathBuild & "/" & Form3.depeTabla
                    pathBuild = pathBuild & "/" & Form3.resDepFieldCode

                    xDs.Tables(0).Rows.Add({"MyName", Form3.resDepFieldName}) '0
                    xDs.Tables(0).Rows.Add({"Object", Form3.resConTempCode}) '1
                    xDs.Tables(0).Rows.Add({"Module", Form3.resConTempModule}) '2
                    xDs.Tables(0).Rows.Add({"TableCode", Form3.resConTableCode}) '3
                    xDs.Tables(0).Rows.Add({"TableName", Form3.resConTableName}) '4
                    xDs.Tables(0).Rows.Add({"FieldCode", Form3.resConFieldCode}) '5
                    xDs.Tables(0).Rows.Add({"FieldName", Form3.resConFieldName}) '6

                    xDs.Tables(0).Rows.Add({"ConditionalType", Form3.resConType}) '7
                    xDs.Tables(0).Rows.Add({"ConditionalRule", Form3.resConRule}) '8
                    xDs.Tables(0).Rows.Add({"ConditionalValue", Form3.resConVal}) '9
                    xDs.Tables(0).Rows.Add({"MatchingFields", Form3.resMachFields}) '10
                    xDs.Tables(0).Rows.Add({"ConditionalScope", Form3.resScope}) '11

                    Await HazPutEnFireBase(pathBuild, xDs)

                    'aqui le ponemos la dependencia en las columnas!
                    DataGridView1.Rows(e.RowIndex).Cells(14).Value = Form3.resConTempCode & ">" & Form3.resConTableCode & ">" & Form3.resConFieldCode
                    DataGridView1.Rows(e.RowIndex).Cells(15).Value = Form3.resConTempCode
                    DataGridView1.Rows(e.RowIndex).Cells(16).Value = Form3.resConTableCode
                    DataGridView1.Rows(e.RowIndex).Cells(17).Value = Form3.resConFieldCode
                    DataGridView1.Rows(e.RowIndex).Cells(18).Value = Form3.resConType
                    DataGridView1.Rows(e.RowIndex).Cells(19).Value = Form3.resConRule
                    DataGridView1.Rows(e.RowIndex).Cells(20).Value = Form3.resConVal
                    DataGridView1.Rows(e.RowIndex).Cells(21).Value = Form3.resMachFields

                Else
                    '
                    If e.ColumnIndex = 23 Then
                        'construction rule!
                        If DataGridView1.Rows(e.RowIndex).Cells(4).Value <> "B - Construction" Then
                            MsgBox("Please select filling rule 'B - Construction' for setting this rule", vbCritical, TitBox)
                            Exit Sub
                        End If

                        If catDs.Tables.Count = 0 Then
                            usaDataset.Tables(0).Rows.Clear()
                            usaDataset.Tables(0).Rows.Add({"catpro"})
                            catDs.Tables.Clear()
                            catDs = Await PullUrlWs(usaDataset, "catpro")
                            ReloadCatPro()
                        End If


                        Form7.allDepes = depeDs
                        Form7.allTemps = tempDs
                        Form7.TablaDatos = tempDs.Tables(tablaNombre)
                        Form7.ConsModule = moduloSelek
                        Form7.ConsObject = objetoSelek
                        Form7.ConsTable = tableSelek
                        Form7.ConsField = DataGridView1.Rows(e.RowIndex).Cells(0).Value
                        Form7.BuildDs = ConstruDs
                        Form7.MisCatas = catDs
                        Form7.maxCharLimit = CInt(DataGridView1.Rows(e.RowIndex).Cells(6).Value)

                        For i = 0 To ConstruDs.Tables.Count - 1
                            If ConstruDs.Tables(i).TableName = objetoSelek Then
                                Form7.posiTabla = i
                                Exit For
                            End If
                        Next

                        Form7.ShowDialog()

                        'hacer reload de la info!
                        usaDataset.Tables(0).Rows.Clear()
                        usaDataset.Tables(0).Rows.Add({"construction"})
                        ConstruDs.Tables.Clear()
                        ConstruDs = Await PullUrlWs(usaDataset, "construction")

                    Else

                        If e.ColumnIndex = 9 Or e.ColumnIndex = 10 Then
                            'catalogos!
                            'que tome el tag y que muestre la info en un box?
                            'ok 
                            If DataGridView1.Rows(e.RowIndex).Cells(4).Value <> "A - From Catalog" Then
                                MsgBox("Please select filling rule 'A - From Catalog' for setting this rule", vbCritical, TitBox)
                                Exit Sub
                            End If

                            'El Tag va a tener:
                            'muestra la regla completa:
                            'gb:gb0001:FIELDMATCH:MATNR#A-BUKRS#B
                            'None>nuevo bind

                            'DataGridView1.Rows(e.rowindex).Cells(e.columnindex).tag
                            'gb:gb0001:FIELDMATCH:MATNR#A-BUKRS#B

                            Dim miTag As Object = DataGridView1.Rows(e.RowIndex).Cells(10).Tag

                            If miTag = "None" Or IsNothing(miTag) = True Then
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

                                If CStr(xObj(0)) <> "" And CStr(xObj(1)) <> "" Then
                                    Form11.toyLeyendo = True
                                    Form11.resCatModule = CStr(xObj(0))
                                    Form11.resCatCode = CStr(xObj(1))
                                    Form11.resCatMF = CStr(xObj(2))
                                    Form11.resCatMC = CStr(xObj(3))
                                Else
                                    'esto se va a quitar despues!!
                                    Form11.toyLeyendo = False
                                    Form11.resCatCode = ""
                                    Form11.resCatModule = ""
                                    Form11.resCatMC = ""
                                    Form11.resCatMF = ""
                                End If

                            End If

                            Form11.tablaTemp = tempDs.Tables(tablaNombre)
                            Form11.objetoModule = moduloSelek
                            Form11.objetoCode = objetoSelek
                            Form11.TablaCode = tableSelek
                            Form11.elCampoCode = CStr(DataGridView1.Rows(e.RowIndex).Cells(0).Value)
                            'Form11.Modulo1 = "gb"
                            'Form11.Modulo2 = moduloSelek
                            Form11.ultiCats = catDs
                            Form11.elEnfoque = "T" 'de templates
                            Form11.ShowDialog()

                            If Form11.huboExito = False Then Exit Sub

                            'se hace reload de catds y se re-escribe en el renglon!

                            DataGridView1.Rows(e.RowIndex).Cells(9).Value = Form11.resCatName
                            DataGridView1.Rows(e.RowIndex).Cells(10).Value = Form11.resCatCode
                            DataGridView1.Rows(e.RowIndex).Cells(10).Tag = Form11.resCatModule & ":" & Form11.resCatCode & ":" & Form11.resCatMF & ":" & Form11.resCatMC
                            'gb:gb0001:FIELDMATCH:MATNR#A-BUKRS#B

                            usaDataset.Tables(0).Rows.Clear()
                            usaDataset.Tables(0).Rows.Add({"catpro"})
                            catDs.Tables.Clear()
                            catDs = Await PullUrlWs(usaDataset, "catpro")
                            ReloadCatPro()

                        Else



                        End If

                    End If


                End If


        End Select

    End Sub

    Private Async Sub ToolStripButton13_Click(sender As Object, e As EventArgs) Handles ToolStripButton13.Click
        'boton para evaluar los registros
        Dim xObj As Object = Nothing

        Select Case ToolStripComboBox1.SelectedIndex
            Case Is = 0


            Case Is = 1


            Case Is = 2


            Case Is = 3
                'records!!
                'obtener la tabla de tempds con el detalle de la tabla, compañía,

                Dim yDs As New DataSet
                'Dim xtraFilt As New DataTable

                For i = 0 To filtDepe.Rows.Count - 1

                    yDs.Tables.Clear()

                    usaDataset.Tables(0).Rows.Clear()
                    usaDataset.Tables(0).Rows.Add({"records"})
                    usaDataset.Tables(0).Rows.Add({elNode.Parent.Parent.Name})
                    usaDataset.Tables(0).Rows.Add({CStr(filtDepe.Rows(i).Item(7))})
                    usaDataset.Tables(0).Rows.Add({CStr(filtDepe.Rows(i).Item(8))})
                    usaDataset.Tables(0).Rows.Add({"records"})

                    yDs = Await PullUrlWs(usaDataset, "records")

                    If yDs.Tables.Count = 0 Then Continue For

                    If multiDepe.Tables.IndexOf(CStr(filtDepe.Rows(i).Item(7)) & "#" & CStr(filtDepe.Rows(i).Item(8)) & "#" & CStr(filtDepe.Rows(i).Item(4))) < 0 Then
                        'la agregamos
                        multiDepe.Tables.Add(yDs.Tables(0).Copy())
                        multiDepe.Tables(multiDepe.Tables.Count - 1).TableName = CStr(filtDepe.Rows(i).Item(7)) & "#" & CStr(filtDepe.Rows(i).Item(8)) & "#" & CStr(filtDepe.Rows(i).Item(4))
                        'multiDepe.Tables(multiDepe.Tables.Count - 1).ExtendedProperties.Add("ObjectName", "")
                        'multiDepe.Tables(multiDepe.Tables.Count - 1).ExtendedProperties.Add("TableName", "")
                    End If

                    'xtraFilt.Columns.Clear()
                    'xtraFilt.Rows.Clear()
                    'xtraFilt = yDs.Tables(0).Clone()

                    'Dim resultados() As DataRow = yDs.Tables(0).Select("DepTableCode = '" & CStr(filtDepe.Rows(i).Item(8)) & "'")
                    'For Each row As DataRow In resultados
                    '    filtDepe.ImportRow(row)
                    'Next

                    If xDepeDs.Tables.IndexOf(CStr(filtDepe.Rows(i).Item(7)) & "#" & CStr(filtDepe.Rows(i).Item(8)) & "#" & CStr(filtDepe.Rows(i).Item(4))) >= 0 Then Continue For

                    Dim YavePrim(1) As DataColumn
                    Dim Yaves As New DataColumn()
                    Yaves.ColumnName = "Key"
                    YavePrim(0) = Yaves

                    'crear estructura de tablas de cada dependencia
                    'vas a agregar tantas columnas dependa matching fields!, si matching fields es None, entonces
                    'solo tomo la columna

                    xDepeDs.Tables.Add(CStr(filtDepe.Rows(i).Item(7)) & "#" & CStr(filtDepe.Rows(i).Item(8)) & "#" & CStr(filtDepe.Rows(i).Item(4)))
                    xDepeDs.Tables(xDepeDs.Tables.Count - 1).Columns.Add(Yaves)
                    xDepeDs.Tables(xDepeDs.Tables.Count - 1).PrimaryKey = YavePrim

                    Dim buscalo As DataRow
                    For j = 0 To yDs.Tables(0).Rows.Count - 1
                        buscalo = xDepeDs.Tables(xDepeDs.Tables.Count - 1).Rows.Find(CStr(yDs.Tables(0).Rows(j).Item(CStr(filtDepe.Rows(i).Item(4)))))
                        If IsNothing(buscalo) = False Then Continue For
                        xDepeDs.Tables(xDepeDs.Tables.Count - 1).Rows.Add({CStr(yDs.Tables(0).Rows(j).Item(CStr(filtDepe.Rows(i).Item(4))))})
                    Next

                Next


                Dim unYve As String = ""
                Dim enCuentra As DataRow
                Dim z As Long
                For i = 0 To filtDepe.Rows.Count - 1
                    unYve = filtDepe.Rows(i).Item(0) & "#" & filtDepe.Rows(i).Item(2)
                    enCuentra = ValidaDt.Rows.Find(CStr(unYve))
                    If IsNothing(enCuentra) = True Then Continue For
                    z = ValidaDt.Rows.IndexOf(enCuentra)
                    ValidaDt.Rows(z).Item(19) = filtDepe.Rows(i).Item(7) & ">" & filtDepe.Rows(i).Item(8) & ">" & filtDepe.Rows(i).Item(4)
                    ValidaDt.Rows(z).Item(20) = filtDepe.Rows(i).Item(7) 'objeto condicional
                    ValidaDt.Rows(z).Item(21) = filtDepe.Rows(i).Item(8) 'conditional table
                    ValidaDt.Rows(z).Item(22) = filtDepe.Rows(i).Item(4) 'conditional field
                    ValidaDt.Rows(z).Item(23) = filtDepe.Rows(i).Item(10) 'dependencie
                    ValidaDt.Rows(z).Item(24) = filtDepe.Rows(i).Item(11)
                    ValidaDt.Rows(z).Item(25) = filtDepe.Rows(i).Item(12)
                    ValidaDt.Rows(z).Item(27) = filtDepe.Rows(i).Item(13) 'matching fields
                    ValidaDt.Rows(z).Item(28) = filtDepe.Rows(i).Item(14) 'Scope
                Next


                'Jalamos los datos para las reglas de construcción!
                usaDataset.Tables(0).Rows.Clear()
                usaDataset.Tables(0).Rows.Add({"construction"})

                BuildBobDs.Tables.Clear()
                BuildBobDs = Await PullUrlWs(usaDataset, "construction")

                TableBuild.Columns.Clear()
                TableBuild.Rows.Clear()

                Dim uniTabla As New DataTable
                Dim TabCons As New DataSet

                If BuildBobDs.Tables.Count > 0 Then
                    Dim posBuild As Integer = 0
                    posBuild = BuildBobDs.Tables.IndexOf(objetoSelek)
                    If posBuild >= 0 Then
                        TableBuild = BuildBobDs.Tables(posBuild).Clone()
                        Dim resulty() As DataRow = BuildBobDs.Tables(posBuild).Select("TableCode = '" & tableSelek & "'")
                        For Each row As DataRow In resulty
                            TableBuild.ImportRow(row)
                        Next

                        uniTabla = TableBuild.DefaultView.ToTable(True, "FieldCode")

                        For i = 0 To uniTabla.Rows.Count - 1
                            Dim construTabla As New DataTable

                            construTabla = TableBuild.Clone()

                            Dim resu() As DataRow = BuildBobDs.Tables(posBuild).Select("TableCode = '" & tableSelek & "' AND FieldCode='" & CStr(uniTabla.Rows(i).Item(0)) & "'")
                            For Each row As DataRow In resu
                                construTabla.ImportRow(row)
                            Next

                            TabCons.Tables.Add(construTabla)

                            TabCons.Tables(TabCons.Tables.Count - 1).TableName = CStr(uniTabla.Rows(i).Item(0))

                        Next


                        For i = 0 To TabCons.Tables.Count - 1
                            'revisamos que YA este la tabla de dependencias agregada a tabcons, en caso de reglas externas
                            For j = 0 To TabCons.Tables(i).Rows.Count - 1
                                If TabCons.Tables(i).Rows(j).Item(5) = "E - External object" Then
                                    'lo tomamos el valor y lo explosionamos
                                    'mm:md21:md21-0001:MATNR:MAKTX#MTBN-MAKNT#MATH:MATCH
                                    xObj = Split(TabCons.Tables(i).Rows(j).Item(6), ":")

                                    If UBound(xObj) <> 5 Then Continue For

                                    yDs.Tables.Clear()

                                    usaDataset.Tables(0).Rows.Clear()
                                    usaDataset.Tables(0).Rows.Add({"records"})
                                    usaDataset.Tables(0).Rows.Add({elNode.Parent.Parent.Name})
                                    usaDataset.Tables(0).Rows.Add({CStr(xObj(1))})
                                    usaDataset.Tables(0).Rows.Add({CStr(xObj(2))})
                                    usaDataset.Tables(0).Rows.Add({"records"})

                                    yDs = Await PullUrlWs(usaDataset, "records")

                                    If yDs.Tables.Count = 0 Then Continue For

                                    If multiDepe.Tables.IndexOf(CStr(xObj(1)) & "#" & CStr(xObj(2)) & "#" & CStr(xObj(3))) < 0 Then
                                        'la agregamos
                                        multiDepe.Tables.Add(yDs.Tables(0).Copy())
                                        multiDepe.Tables(multiDepe.Tables.Count - 1).TableName = CStr(xObj(1)) & "#" & CStr(xObj(2)) & "#" & CStr(xObj(3))
                                        'multiDepe.Tables(multiDepe.Tables.Count - 1).ExtendedProperties.Add("ObjectName", "")
                                        'multiDepe.Tables(multiDepe.Tables.Count - 1).ExtendedProperties.Add("TableName", "")
                                    End If


                                End If

                            Next


                        Next

                    End If

                Else
                    TableBuild.Columns.Add()
                End If




                'Aquii, jalamos de Construds el template y filtramos por tabla si es que tiene, una tabla general de construccion!
                'BuildBobDs
                'jalamos de varias 

                'Aquii, un ciclo para jalar los registros de las tablas dependientes

                'un renglon


                DataGridView1.ShowCellToolTips = True
                Dim rowRep As Integer = 0
                'Dim y1, y2 As Boolean
                Dim esMandatorio As Boolean = False
                Dim esOpcional As Boolean = False
                Dim continuoEvaluacion As Boolean = False
                Dim anchoAValidar As Integer = 0
                Dim unError As String = ""
                Dim valEvaluar As String = ""
                Dim unComodin As String = ""
                Dim anchoTru As Object
                Dim findAnother As DataRow
                Dim posOtro As Integer = 0
                Dim esYave As Boolean = False
                Dim hayYave As Boolean = False
                Dim concaYave As String = ""
                Dim k As Integer = 0
                Dim elRowOk As Boolean = False
                Dim xLlena As New DataTable
                Dim yaveInte As String = ""
                Dim mixKeys As String = ""
                Dim cuenOk As Integer = 0
                Dim cuenBad As Integer = 0
                Dim cuenAll As Integer = 0
                Dim expText As String = ""
                'Dim xObj As Object = Nothing
                Dim yObj As Object = Nothing
                Dim zObj As Object = Nothing
                Dim mixNombreTab As String = ""
                Dim buskEst As String = ""
                Dim valBus As String = ""
                Dim valAjus As String = ""

                'Dim enCuentra As DataRow
                'Randomize()
                'Dim value As Integer = CInt(Int((6 * Rnd()) + 1))
                ReDim anchoTru(0 To DataGridView1.Columns.Count - 4)

                Dim iAveprim(1) As DataColumn
                Dim kEys As New DataColumn()
                kEys.ColumnName = "KeyField"
                iAveprim(0) = kEys

                Dim xDt As New DataTable
                'creamos una columna con campo llave, que sea la concatenación de varios!
                'los empezamos a agregar, si Llega a haber un error al agregar, señalamos que h
                xDt.Columns.Add(kEys)
                xDt.CaseSensitive = True
                xDt.PrimaryKey = iAveprim

                hayYave = False

                Dim w As Integer = 0


                'Dim YavePrimaria(1) As DataColumn
                'Dim Yavez As New DataColumn()
                'Yavez.ColumnName = "Unique"
                'YavePrimaria(0) = Yavez
                'uniTabla.Columns.Add(Yavez)
                Dim mixPobl As String = ""

                Await SigoVivo()
                Module5.AgregaTiempo()
                toyTrabajando = True

                ToolStripLabel1.Visible = True

                DoneEvaluacion = False

                For i = 0 To DataGridView1.Rows.Count - 2

                    Application.DoEvents()

                    ToolStripLabel1.Text = "Evaluating..." & CStr(i + 1) & " / " & DataGridView1.Rows.Count - 1

                    'System.Threading.Thread.Sleep(0)
                    DataGridView1.Rows(i).Tag = "NO OK"
                    'se pone en cero!
                    For j = 0 To DataGridView1.Columns.Count - 4
                        anchoTru(j) = False
                    Next

                    elRowOk = True
                    concaYave = ""
                    k = 0
                    For j = 0 To DataGridView1.Columns.Count - 4

                        'If DataGridView1.Columns(j).name
                        LimpiaCeldaDeError(i, j)
                        esYave = False

                        If IsNothing(DataGridView1.Rows(i).Cells(j).Value) = True Then
                            valEvaluar = ""
                        Else
                            If IsDBNull(DataGridView1.Rows(i).Cells(j).Value) = True Then
                                valEvaluar = ""
                            Else
                                valEvaluar = DataGridView1.Rows(i).Cells(j).Value
                            End If
                        End If

                        'valEvaluar = DataGridView1.Rows(i).Cells(j).Value
                        esMandatorio = False
                        esOpcional = False
                        continuoEvaluacion = False

                        unYve = tableSelek & "#" & DataGridView1.Columns(j).Name
                        enCuentra = ValidaDt.Rows.Find(CStr(unYve))
                        If IsNothing(enCuentra) = True Then
                            Continue For 'no se puede validar!
                        End If
                        z = ValidaDt.Rows.IndexOf(enCuentra)

                        Select Case ValidaDt.Rows(z).Item(13)
                            Case Is = "None", ""
                                'no se hace nada
                            Case Is = "Left"
                                valEvaluar = LTrim(valEvaluar)

                            Case Is = "Right"
                                valEvaluar = RTrim(valEvaluar)

                            Case Is = "Both"
                                valEvaluar = Trim(valEvaluar)

                        End Select



                        If ValidaDt.Rows(z).Item(5) = "X" Then
                            esMandatorio = True
                            esYave = True
                            k = k + 1

                            If k = 1 Then
                                concaYave = valEvaluar
                            Else
                                concaYave = concaYave & "#" & valEvaluar
                            End If

                        Else
                            'se evalua!

                            Select Case ValidaDt.Rows(z).Item(8)

                                Case Is = "No selection", ""


                                Case Is = "Mandatory"
                                    esMandatorio = True

                                Case Is = "Optional"
                                    esMandatorio = True
                                    esOpcional = True

                                Case Is = "Conditional"
                                    'esMandatorio = True
                                    'ojoo, aqui falta!
                                    'debe evaluarse si cumple la condicion, local, interna ó externa!
                                    'conditional
                                    'recuerda, local es verificar una columna o serie de columnas hermanas,
                                    'interna es en una tabla hermana
                                    'externa es un objeto y tabla externa!
                                    'en el condicional se va a poner
                                    'desde aqui podria continuar si NO cumple
                                    'Si cumple la condicion, entonces valida con la regla que se haya puesto
                                    'Si cumple la condicion, debe existir la informacion en la otra hoja

                                    xObj = Nothing
                                    Dim valPrufz As String = ""
                                    Dim misRows() As DataRow
                                    Dim busCad As String = ""
                                    Dim valCheck As String = ""
                                    'Nombre de la tabla condicional a buscar:
                                    mixNombreTab = ValidaDt.Rows(z).Item(20) & "#" & ValidaDt.Rows(z).Item(21) & "#" & ValidaDt.Rows(z).Item(22)

                                    'Esto es independiente de la opcion de arriba!
                                    If CStr(ValidaDt.Rows(z).Item(28)) = "To Condition" Then
                                        valPrufz = CStr(ValidaDt.Rows(z).Item(25))
                                    Else
                                        valPrufz = valEvaluar
                                    End If

                                    'si es local, de la misma tabla, entonces se evalúa vs el datagrid!!
                                    'DEBE tener un matching field cuando es vs una validación de objeto externo!?
                                    buskEst = ""

                                    Select Case CStr(ValidaDt.Rows(z).Item(23))
                                        Case Is = "Local"
                                            'Solo si es local debería validar SIEMPRE vs su mismo renglon y ya!
                                            If CStr(ValidaDt.Rows(z).Item(22)) = "" Then
                                                PintaCeldaDeError(i, j, "The conditional matching field(s) is not set!!, please review with the module responsible of defining condition rules!")
                                                Continue For
                                            Else
                                                valCheck = DataGridView1.Rows(i).Cells(CStr(ValidaDt.Rows(z).Item(22))).Value
                                            End If

                                            'SIEMPRE va a tener NONE en matching fields!


                                        Case Else 'Internal o External

                                            If multiDepe.Tables.IndexOf(mixNombreTab) < 0 Then
                                                PintaCeldaDeWarning(i, j, "Warning!!, There is no data to validate either if this field is right o wrong. Conditional path not found: " & CStr(ValidaDt.Rows(z).Item(20)) & " > " & CStr(ValidaDt.Rows(z).Item(21)) & " > " & CStr(ValidaDt.Rows(z).Item(22)))
                                                'ValidaDt.Rows(z).Item(20) & "#" & ValidaDt.Rows(z).Item(21) & "#" & ValidaDt.Rows(z).Item(22)
                                                Continue For
                                            End If

                                            'es como si fuera vs un ToBeValue de una columna externa de otro template!
                                            If ValidaDt.Rows(z).Item(27) <> "None" Then
                                                'es Interno ó externo!!
                                                'tiene matching fields!

                                                xObj = Split(CStr(ValidaDt.Rows(z).Item(27)), "-")

                                                'buskEst = ""
                                                mixPobl = ""
                                                busCad = ""

                                                For w = 0 To UBound(xObj)

                                                    'hacer split de los campos
                                                    'Primero el hijo, luego el condicional
                                                    'el primero es de la tabla actual, el segundo de la tabla condicional
                                                    If w <> 0 Then buskEst = buskEst & " AND "

                                                    If w <> 0 Then mixPobl = mixPobl & " > "

                                                    If w <> 0 Then busCad = busCad & " > "

                                                    yObj = Nothing
                                                    yObj = Split(CStr(xObj(w)), "#")

                                                    mixPobl = mixPobl & CStr(yObj(1))

                                                    valBus = DataGridView1.Rows(i).Cells(CStr(yObj(0))).Value

                                                    busCad = busCad & valBus

                                                    buskEst = buskEst & CStr(yObj(1)) 'siempre la primera posición
                                                    buskEst = buskEst & "='" & valBus & "'"

                                                Next

                                                misRows = multiDepe.Tables(mixNombreTab).Select(buskEst)

                                                If CStr(ValidaDt.Rows(z).Item(28)) = "To Condition" Then
                                                    'solo para condicionar la evaluación del valor!
                                                    If misRows.Count = 0 Then
                                                        'como NO existe la condición NO se evalua!, todo OK!
                                                        anchoTru(j) = True
                                                        Continue For
                                                    End If

                                                Else
                                                    'para aplicar la regla!
                                                    If misRows.Count = 0 Then
                                                        'NO existe la dependencia!!
                                                        PintaCeldaDeError(i, j, "The conditional matching field(s) > " & mixPobl & " : " & busCad & " , were not found on the conditional object path: " & ValidaDt.Rows(z).Item(20) & ">" & ValidaDt.Rows(z).Item(21) & " , please review!")
                                                        Continue For
                                                    End If

                                                End If

                                                Dim miTabF As DataTable
                                                miTabF = multiDepe.Tables(mixNombreTab).Clone()
                                                For Each row In misRows
                                                    miTabF.ImportRow(row)
                                                Next

                                                If miTabF.Rows.Count = 1 Then
                                                    valCheck = miTabF.Rows(0).Item(CStr(ValidaDt.Rows(z).Item(22)))
                                                Else
                                                    valCheck = "Multiple"
                                                End If

                                            Else
                                                'Directo vs ToBeValue
                                                valCheck = "Multiple"
                                                'Conditional>OR>Value
                                                'Hay de 2, Tobe Value, para condicionar o para aplicar

                                                'Si es para condicionar estamos diciendo que con que 1 valor de las opciones seleccionadas
                                                'cumpla la regla se aplica la regla exterior, se debe usar valprufz

                                                'Si es para aplicar estamos diciendo que valevaluar debe hacer match con la regla seleccionada
                                                'dentro de la columna de valores

                                            End If


                                    End Select

                                    'Cuando sea vs evaluación NO se puede seleccionar OR, NULL y NOT NULL

                                    Select Case ValidaDt.Rows(z).Item(24)

                                        Case Is = "OR"
                                            xObj = Nothing
                                            xObj = Split(valPrufz, "#")
                                            Dim siMachea As Boolean = False
                                            For w = 0 To UBound(xObj)
                                                'evaluar si alguno matchea!
                                                If valCheck = xObj(w) Then
                                                    siMachea = True
                                                    Exit For
                                                End If
                                            Next

                                            'si condiciona el valor
                                            If siMachea = True Then
                                                If CStr(ValidaDt.Rows(z).Item(28)) = "To Condition" Then
                                                    'aplica la condición, si se evalua para adelante!
                                                    esMandatorio = True
                                                End If
                                                'si NO machea NO es mandatorio!, fin!
                                            Else
                                                esOpcional = True
                                            End If

                                        Case Is = "NULL" 'va se mandatorio si el campo SI es vacío!
                                            'mitabf.
                                            If valCheck = "" Then esMandatorio = True Else esOpcional = True
                                                'esMandatorio = True
                                                'PintaCeldaDeError(i, j, "The conditional final matching field > " & mixPobl & " > " & CStr(ValidaDt.Rows(z).Item(22)) & " : " & busCad & " > " & valCheck & " , must be null were not found on the conditional object path: " & ValidaDt.Rows(z).Item(20) & ">" & ValidaDt.Rows(z).Item(21) & ">" & ValidaDt.Rows(z).Item(22) & " , please review!")

                                        Case Is = "NOTNULL" 'va ser mandatorio si el campo NO es vacío
                                            If valCheck <> "" Then esMandatorio = True Else esOpcional = True


                                        Case Is = "MATCHS" 'va ser mandatorio si el campo matchea con el valor condicional!

                                            If valCheck = "Multiple" Then

                                                'Aqui falta!!
                                                If buskEst <> "" Then buskEst = buskEst & " AND "

                                                buskEst = buskEst & CStr(ValidaDt.Rows(z).Item(22)) & "='" & valPrufz & "'"

                                                Dim xtraRows() As DataRow = multiDepe.Tables(mixNombreTab).Select(buskEst)
                                                If xtraRows.Length = 0 Then
                                                    PintaCeldaDeError(i, j, "The conditional matching field(s) > " & mixPobl & " : " & busCad & " , were not found on the conditional object path: " & ValidaDt.Rows(z).Item(20) & ">" & ValidaDt.Rows(z).Item(21) & " > " & CStr(ValidaDt.Rows(z).Item(22)) & " , please review!")
                                                    Continue For
                                                End If

                                            Else
                                                If valCheck = valPrufz Then
                                                    'si macheo!
                                                    If CStr(ValidaDt.Rows(z).Item(28)) = "To Condition" Then
                                                        esMandatorio = True
                                                    Else
                                                        'se evalua!!
                                                        'la regla esta OK!
                                                        'continuo!
                                                        anchoTru(j) = True
                                                        Continue For
                                                    End If
                                                Else
                                                    If CStr(ValidaDt.Rows(z).Item(28)) = "To Condition" Then
                                                        'NO es mandatorio, se queda vacío!
                                                        esMandatorio = False
                                                        esOpcional = True
                                                    Else
                                                        PintaCeldaDeError(i, j, "The conditional final matching field > " & mixPobl & " > " & CStr(ValidaDt.Rows(z).Item(22)) & " : " & busCad & " > " & valCheck & " , does not match to the required value: " & valPrufz & " , review conditional object path: " & ValidaDt.Rows(z).Item(20) & ">" & ValidaDt.Rows(z).Item(21) & ">" & ValidaDt.Rows(z).Item(22) & " ")
                                                        Continue For
                                                    End If

                                                End If
                                            End If



                                        Case Is = "STARTWITH"

                                            If valCheck = "Multiple" Then
                                                If buskEst <> "" Then buskEst = buskEst & " AND "
                                                buskEst = buskEst & CStr(ValidaDt.Rows(z).Item(22)) & " LIKE '" & valPrufz & "%'"
                                                Dim xtraRows() As DataRow = multiDepe.Tables(mixNombreTab).Select(buskEst)
                                                If xtraRows.Length = 0 Then
                                                    PintaCeldaDeError(i, j, "The conditional matching field(s) > " & mixPobl & " : " & busCad & " , were not found on the conditional object path: " & ValidaDt.Rows(z).Item(20) & ">" & ValidaDt.Rows(z).Item(21) & " > " & CStr(ValidaDt.Rows(z).Item(22)) & " , please review!")
                                                    Continue For
                                                End If
                                            Else
                                                If valCheck.StartsWith(valPrufz) = True Then
                                                    'cumple la condición, 
                                                    If CStr(ValidaDt.Rows(z).Item(28)) = "To Condition" Then
                                                        esMandatorio = True
                                                    Else
                                                        anchoTru(j) = True
                                                        Continue For
                                                    End If

                                                Else
                                                    If CStr(ValidaDt.Rows(z).Item(28)) = "To Condition" Then
                                                        esMandatorio = False
                                                        esOpcional = True
                                                    Else
                                                        PintaCeldaDeError(i, j, "The conditional final matching field > " & mixPobl & " > " & CStr(ValidaDt.Rows(z).Item(22)) & " : " & busCad & " > " & valCheck & " , does not starts with the required value: " & valPrufz & " , review conditional object path: " & ValidaDt.Rows(z).Item(20) & ">" & ValidaDt.Rows(z).Item(21) & ">" & ValidaDt.Rows(z).Item(22) & " ")
                                                        Continue For
                                                    End If

                                                End If
                                            End If



                                        Case Is = "ENDWITH"

                                            If valCheck = "Multiple" Then
                                                If buskEst <> "" Then buskEst = buskEst & " AND "
                                                buskEst = buskEst & CStr(ValidaDt.Rows(z).Item(22)) & " LIKE '%" & valPrufz & "'"
                                                Dim xtraRows() As DataRow = multiDepe.Tables(mixNombreTab).Select(buskEst)
                                                If xtraRows.Length = 0 Then
                                                    PintaCeldaDeError(i, j, "The conditional matching field(s) > " & mixPobl & " : " & busCad & " , were not found on the conditional object path: " & ValidaDt.Rows(z).Item(20) & ">" & ValidaDt.Rows(z).Item(21) & " > " & CStr(ValidaDt.Rows(z).Item(22)) & " , please review!")
                                                    Continue For
                                                End If
                                            Else
                                                If valCheck.EndsWith(valPrufz) = True Then

                                                    If CStr(ValidaDt.Rows(z).Item(28)) = "To Condition" Then
                                                        esMandatorio = True
                                                    Else
                                                        anchoTru(j) = True
                                                        Continue For
                                                    End If

                                                Else
                                                    If CStr(ValidaDt.Rows(z).Item(28)) = "To Condition" Then
                                                        esMandatorio = False
                                                        esOpcional = True
                                                    Else
                                                        PintaCeldaDeError(i, j, "The conditional final matching field > " & mixPobl & " > " & CStr(ValidaDt.Rows(z).Item(22)) & " : " & busCad & " > " & valCheck & " , does not ends with the required value: " & valPrufz & " , review conditional object path: " & ValidaDt.Rows(z).Item(20) & ">" & ValidaDt.Rows(z).Item(21) & ">" & ValidaDt.Rows(z).Item(22) & " ")
                                                        Continue For
                                                    End If

                                                End If
                                            End If




                                        Case Is = "CONTAINS"

                                            If valCheck = "Multiple" Then
                                                If buskEst <> "" Then buskEst = buskEst & " AND "
                                                buskEst = buskEst & CStr(ValidaDt.Rows(z).Item(22)) & " LIKE '%" & valPrufz & "%'"
                                                Dim xtraRows() As DataRow = multiDepe.Tables(mixNombreTab).Select(buskEst)
                                                If xtraRows.Length = 0 Then
                                                    PintaCeldaDeError(i, j, "The conditional matching field(s) > " & mixPobl & " : " & busCad & " , were not found on the conditional object path: " & ValidaDt.Rows(z).Item(20) & ">" & ValidaDt.Rows(z).Item(21) & " > " & CStr(ValidaDt.Rows(z).Item(22)) & " , please review!")
                                                    Continue For
                                                End If
                                            Else
                                                If valCheck.Contains(valPrufz) = True Then
                                                    If CStr(ValidaDt.Rows(z).Item(28)) = "To Condition" Then
                                                        esMandatorio = True
                                                    Else
                                                        anchoTru(j) = True
                                                        Continue For
                                                    End If
                                                Else
                                                    If CStr(ValidaDt.Rows(z).Item(28)) = "To Condition" Then
                                                        esMandatorio = False
                                                        esOpcional = True
                                                    Else
                                                        PintaCeldaDeError(i, j, "The conditional final matching field > " & mixPobl & " > " & CStr(ValidaDt.Rows(z).Item(22)) & " : " & busCad & " > " & valCheck & " , does not CONTAINS the required value: " & valPrufz & " , review conditional object path: " & ValidaDt.Rows(z).Item(20) & ">" & ValidaDt.Rows(z).Item(21) & ">" & ValidaDt.Rows(z).Item(22) & " ")
                                                        Continue For
                                                    End If
                                                End If
                                            End If



                                    End Select



                                    'multidepe va a contener las tablas para validar los datos de todos 
                                    'los objetos dependientes, dependiendo de lo que se valide es la consulta que se
                                    'hace, se usa un tabla de ayuda para hacer el select correspondiente!
                                    'esto es para evaluar si se evalua lo que puso el usuario!!

                                    'va ser una regla que va a condicionar para que SI se evalúe o NO el valor
                                    'ó es la regla que se va a aplicar para evaluar el valor per se!?
                                    'Condición para evaluar el valor, ó 


                            End Select

                        End If

                        If esOpcional = True Then
                            If valEvaluar = "" Then
                                anchoTru(j) = True
                                continuoEvaluacion = False
                            Else
                                continuoEvaluacion = True
                            End If
                        Else
                            continuoEvaluacion = True
                        End If

                        If continuoEvaluacion = False Then Continue For


                        If esMandatorio = True Then

                            If valEvaluar = "" Then
                                If CStr(ValidaDt.Rows(z).Item(10)) = "Indicator" Then
                                    PintaCeldaDeWarning(i, j, "This field is mandatory to set as indicator (True or False = 'X' or 'blank'), currently is set to False. Just review this the correct setting")
                                    anchoTru(j) = True
                                    Continue For
                                Else
                                    'Se pone errors!!
                                    PintaCeldaDeError(i, j, "This field is mandatory!, please fill it accordingly!!")
                                    Continue For
                                End If
                            End If

                        End If

                        'If esMandatorio = True And valEvaluar = "" Then
                        '    PintaCeldaDeError(i, j, "This field is mandatory!, please fill it accordingly!!")
                        '    Continue For
                        'End If

                        'aplica!
                        'si es campo llave, es mandatorio!!
                        'OJOO, al final se debe evaluar la concatenación de los campos llaves donde aplique!!
                        'si es decimal no va asi!!
                        'anchoAValidar

                        If CStr(ValidaDt.Rows(z).Item(11)) = "" Then
                            anchoAValidar = 0
                        Else
                            If CStr(ValidaDt.Rows(z).Item(11)).Contains(".") Then
                                xObj = Split(CStr(ValidaDt.Rows(z).Item(11)), ".")
                                anchoAValidar = CInt(xObj(0)) + CInt(xObj(1)) + 1
                            Else
                                anchoAValidar = CInt(ValidaDt.Rows(z).Item(11))
                            End If
                        End If

                        If isLengthOk(valEvaluar, anchoAValidar, unError) = False Then
                            PintaCeldaDeError(i, j, unError)
                            Continue For
                        End If

                        Select Case CStr(ValidaDt.Rows(z).Item(10))'data type

                            Case Is = "No selection", ""



                            Case Is = "Number"
                                If isIntegerNumberOk(valEvaluar, unError) = False Then
                                    PintaCeldaDeError(i, j, unError)
                                    Continue For
                                End If
                                DataGridView1.Rows(i).Cells(j).Value = valEvaluar


                            Case Is = "Date"
                                Select Case anchoAValidar
                                    Case Is = 10
                                        If isFieldinDateFormat(valEvaluar, unComodin, unError) = False Then
                                            PintaCeldaDeError(i, j, unError)
                                            Continue For
                                        End If
                                        'se lo corregimos!
                                        DataGridView1.Rows(i).Cells(j).Value = unComodin

                                    Case Is = 8
                                        If isDate8FormatOk(valEvaluar, unError) = False Then
                                            PintaCeldaDeError(i, j, unError)
                                            Continue For
                                        End If

                                End Select


                            Case Is = "Text"
                                '?

                            Case Is = "Email"
                                If isEmailValid(valEvaluar, unError) = False Then
                                    PintaCeldaDeError(i, j, unError)
                                    Continue For
                                End If

                            Case Is = "Decimal"
                                valAjus = ""
                                If isDecimalOk(valEvaluar, CDec(ValidaDt.Rows(z).Item(11)), unError, valAjus) = False Then
                                    PintaCeldaDeError(i, j, unError)
                                    Continue For
                                End If
                                DataGridView1.Rows(i).Cells(j).Value = valAjus

                            Case Is = "Indicator"
                                If isIndicatorFieldOk(valEvaluar, unError) = False Then
                                    PintaCeldaDeError(i, j, unError)
                                    Continue For
                                End If

                            Case Is = "Time"
                                If isFieldinTimeFormat(valEvaluar, CInt(ValidaDt.Rows(z).Item(11)), unError, unComodin) = False Then
                                    PintaCeldaDeError(i, j, unError)
                                    Continue For
                                End If
                                DataGridView1.Rows(i).Cells(j).Value = unComodin

                        End Select


                        Select Case CStr(ValidaDt.Rows(z).Item(9))

                            Case Is = "A - From Catalog"
                                'CStr(ValidaDt.Rows(z).Item(14))'tiene posicion del catálogo, con 
                                'este buscamos y ubicamos la tabla a buscar!
                                'tm#gb0001
                                '
                                'ValidaDt.Rows(z).Item(14)' catalog code
                                'ValidaDt.Rows(z).Item(29)'catalog module
                                'ValidaDt.Rows(z).Item(30)'match field
                                'ValidaDt.Rows(z).Item(31)'match conditions
                                unComodin = CStr(ValidaDt.Rows(z).Item(29)) & "#" & ValidaDt.Rows(z).Item(14)
                                If catDs.Tables.IndexOf(unComodin) < 0 Then
                                    anchoTru(j) = True 'se pone OK, pero hay warnings!
                                    PintaCeldaDeWarning(i, j, "Unable to found catalog related: " & CStr(ValidaDt.Rows(z).Item(29)) & " > " & ValidaDt.Rows(z).Item(14) & " > " & ValidaDt.Rows(z).Item(15))
                                    Continue For
                                Else

                                    Dim yTabla As New DataTable
                                    yTabla = catDs.Tables(unComodin).Clone()

                                    Dim rowX() As DataRow
                                    Dim mixCade As String = ""
                                    Dim cadCOmp As String = ""
                                    yObj = Nothing
                                    zObj = Nothing

                                    If ValidaDt.Rows(z).Item(31) = "" Or ValidaDt.Rows(z).Item(31) = "None" Then
                                        'match directo a la columna, sin filtros
                                        If CStr(ValidaDt.Rows(z).Item(30)).Contains(">") = True Or CStr(ValidaDt.Rows(z).Item(30)).Contains("<") = True Then
                                            'es entre 2 numeros
                                            'OJOO, SI ESTA BIEEN, solo falta obligar al usuario a requerir un match condicional!!
                                            'Aqui si o SI debe haber un match field!
                                            '1000>2000
                                            '1000<2000
                                            'Entonces este caso se supondría imposible!
                                            'este caso nunca va a entrar!
                                            If CStr(ValidaDt.Rows(z).Item(30)).Contains(">") = True Then
                                                yObj = Split(CStr(ValidaDt.Rows(z).Item(30)), ">")
                                                cadCOmp = ">"
                                            Else
                                                yObj = Split(CStr(ValidaDt.Rows(z).Item(30)), "<")
                                                cadCOmp = "<"
                                            End If

                                            rowX = catDs.Tables(unComodin).Select(mixCade)

                                            For Each row As DataRow In rowX
                                                yTabla.ImportRow(row)
                                            Next

                                            'Validación extra!!
                                            If yTabla.Rows.Count = 0 Then
                                                PintaCeldaDeError(i, j, "The value " & valEvaluar & " does not exists on the " & CStr(ValidaDt.Rows(z).Item(15)) & " catalog")
                                                Continue For
                                            End If


                                        Else
                                            'Comparación directa!
                                            rowX = catDs.Tables(unComodin).Select(CStr(ValidaDt.Rows(z).Item(30)) & " = '" & valEvaluar & "'")
                                            For Each row As DataRow In rowX
                                                yTabla.ImportRow(row)
                                            Next
                                            'Directo
                                            If yTabla.Rows.Count = 0 Then
                                                PintaCeldaDeError(i, j, "The value " & valEvaluar & " does not exists on the " & CStr(ValidaDt.Rows(z).Item(15)) & " catalog")
                                                Continue For
                                            End If

                                        End If

                                    Else
                                        mixPobl = ""
                                        yObj = Split(CStr(ValidaDt.Rows(z).Item(31)), "-")
                                        mixCade = ""
                                        For w = 0 To UBound(yObj)
                                            zObj = Split(CStr(yObj(w)), "#")
                                            '0 para el grid, 1 para el catalogo
                                            If mixCade <> "" Then mixCade = mixCade & " AND "
                                            If mixPobl <> "" Then mixPobl = mixPobl & " > "
                                            mixPobl = mixPobl & CStr(zObj(1))

                                            If CStr(zObj(0)) = "Company Code" Then
                                                'hace match vs el nodo!
                                                mixCade = mixCade & CStr(zObj(1)) & " = '" & compaSelekted & "'"
                                            Else
                                                mixCade = mixCade & CStr(zObj(1)) & " = '" & DataGridView1.Rows(i).Cells(CStr(zObj(0))).Value & "'"
                                            End If

                                        Next

                                        'match con filtros adicionales
                                        If CStr(ValidaDt.Rows(z).Item(30)).Contains(">") = True Or CStr(ValidaDt.Rows(z).Item(30)).Contains("<") = True Then
                                            'es entre 2 numeros
                                            'AQUI FALTA la Concatenacion adicional para

                                            rowX = catDs.Tables(unComodin).Select(mixCade)

                                            For Each row As DataRow In rowX
                                                yTabla.ImportRow(row)
                                            Next
                                            'validacion extra!

                                            If yTabla.Rows.Count = 0 Then
                                                PintaCeldaDeWarning(i, j, "Unable yo define if the value " & valEvaluar & " is in the correct range due to missing match with catalog traceability: " & CStr(ValidaDt.Rows(z).Item(15)) & " " & mixPobl)
                                                Continue For
                                            End If

                                            '0,1
                                            '2000>x>1000
                                            '1000<x<2000

                                            If CStr(ValidaDt.Rows(z).Item(30)).Contains(">") = True Then
                                                yObj = Split(CStr(ValidaDt.Rows(z).Item(30)), ">")
                                                cadCOmp = ">"
                                            Else
                                                yObj = Split(CStr(ValidaDt.Rows(z).Item(30)), "<")
                                                cadCOmp = "<"
                                            End If

                                            If cadCOmp = ">" Then
                                                If CDec(valEvaluar) > CDec(yObj(1)) Then

                                                    If CDec(valEvaluar) < CDec(yObj(0)) Then
                                                        'Ok!
                                                    Else
                                                        'error!
                                                        PintaCeldaDeError(i, j, "The value " & valEvaluar & " must be between " & CStr(yObj(1)) & " and " & CStr(yObj(0)))
                                                        Continue For
                                                    End If

                                                Else
                                                    PintaCeldaDeError(i, j, "The value " & valEvaluar & " must be between " & CStr(yObj(1)) & " and " & CStr(yObj(0)))
                                                    Continue For
                                                End If
                                            Else
                                                If CDec(valEvaluar) > CDec(yObj(0)) Then
                                                    If CDec(valEvaluar) < CDec(yObj(1)) Then
                                                        'ok!
                                                    Else
                                                        PintaCeldaDeError(i, j, "The value " & valEvaluar & " must be between " & CStr(yObj(0)) & " and " & CStr(yObj(1)))
                                                        Continue For
                                                    End If
                                                Else
                                                    PintaCeldaDeError(i, j, "The value " & valEvaluar & " must be between " & CStr(yObj(0)) & " and " & CStr(yObj(1)))
                                                    Continue For
                                                End If


                                            End If

                                        Else
                                            'match ultimo
                                            mixCade = mixCade & " AND " & CStr(ValidaDt.Rows(z).Item(30)) & " = '" & valEvaluar & "'"
                                            rowX = catDs.Tables(unComodin).Select(mixCade)

                                            For Each row As DataRow In rowX
                                                yTabla.ImportRow(row)
                                            Next

                                            If yTabla.Rows.Count = 0 Then
                                                PintaCeldaDeError(i, j, "The value " & valEvaluar & " does not exists on the " & CStr(ValidaDt.Rows(z).Item(15)) & " catalog  and fields " & mixPobl)
                                                Continue For
                                            End If

                                        End If

                                    End If


                                    'unComodin = Microsoft.VisualBasic.Left(CStr(ValidaDt.Rows(z).Item(14)), 2) & "#" & CStr(ValidaDt.Rows(z).Item(14))
                                    'findAnother = catDs.Tables(unComodin).Rows.Find(valEvaluar)
                                    'If IsNothing(findAnother) = True Then
                                    '    PintaCeldaDeError(i, j, "The value " & valEvaluar & " does not exists on the " & CStr(ValidaDt.Rows(z).Item(15)) & " catalog")
                                    '    Continue For
                                    'End If

                                End If



                            Case Is = "B - Construction"
                                'este falta!, llamar a la validacion de construccion!
                                'construTabla.Columns.Clear()
                                'construTabla.Rows.Clear()
                                Dim posBi As Integer = 0
                                posBi = TabCons.Tables.IndexOf(CStr(ValidaDt.Rows(z).Item(3)))
                                'sacamos la tabla y jalamos la info!
                                Dim meQuedolaRegla As Boolean = False

                                If posBi >= 0 Then

                                    Dim tab1 As DataTable = TabCons.Tables(posBi).DefaultView.ToTable(True, "RuleKey")
                                    Dim miObjeto As Object = Nothing
                                    Dim poSiclo As Integer = 1
                                    Dim numChars As Integer = 0
                                    Dim cadFix As String = ""
                                    Dim cadLeft As String = ""
                                    Dim partCad As String = ""
                                    'Dim partCad As String = ""

                                    Dim cadError As String = ""
                                    Dim longAcum As Integer = 0
                                    Dim charResta As Integer = 0

                                    Dim pedX As String = ""
                                    For w = 0 To tab1.Rows.Count - 1
                                        Dim xTabla As New DataTable
                                        xTabla = TabCons.Tables(posBi).Clone()

                                        Dim result() As DataRow = TabCons.Tables(posBi).Select("RuleKey = '" & CStr(tab1.Rows(w).Item(0)) & "'", "Consec ASC")

                                        For Each row As DataRow In result
                                            xTabla.ImportRow(row)
                                        Next

                                        poSiclo = 1
                                        cadFix = ""
                                        cadLeft = valEvaluar
                                        meQuedolaRegla = True
                                        longAcum = 0
                                        charResta = 0
                                        'aqui verificamos de una!
                                        For r = 0 To xTabla.Rows.Count - 1

                                            numChars = CInt(xTabla.Rows(r).Item(7))
                                            miObjeto = Nothing
                                            poSiclo = 1

                                            Select Case CStr(xTabla.Rows(r).Item(5))

                                                Case Is = "A - From Catalog"

                                                    xObj = Split(CStr(xTabla.Rows(r).Item(6)), ":")

                                                    If UBound(xObj) <> 3 Then

                                                        meQuedolaRegla = False

                                                        cadError = "Unable to find the proper catalog to find this matching rule: " & CStr(xTabla.Rows(r).Item(6)) & " , please review with dev team"

                                                        Exit For

                                                    End If

                                                    mixNombreTab = CStr(xObj(0)) & "#" & CStr(xObj(1))

                                                    If catDs.Tables.IndexOf(mixNombreTab) < 0 Then
                                                        meQuedolaRegla = False

                                                        cadError = "Unable to find the proper catalog to find this matching rule: " & CStr(xTabla.Rows(r).Item(6)) & " , please review with dev team"

                                                        Exit For
                                                    End If

                                                    'regla puede quedar:
                                                    '1. Si hay caracter de concatenación al final, se hace split, y se toma el primer valor SIEMPRE
                                                    '2. A ese primer valor se le quitan los espacios en blanco al inicio y final con trim
                                                    '3. Se compara vs el ancho de caracteres y vs el catalogo, si todo OK, continua
                                                    '4. Se recalcula cadleft!

                                                    buskEst = ""

                                                    If CStr(xObj(3)) <> "None" And CStr(xObj(3)) <> "" Then
                                                        'se requiere validación match adicional!
                                                        'del campo mio al campo del objeto externo
                                                        yObj = Split(CStr(xObj(3)), "-")
                                                        For p = 0 To UBound(yObj)

                                                            If p <> 0 Then buskEst = buskEst & " AND "

                                                            zObj = Split(CStr(yObj(p)), "#")

                                                            buskEst = buskEst & CStr(zObj(1)) & " = '"

                                                            buskEst = buskEst & DataGridView1.Rows(i).Cells(CStr(zObj(0))).Value & "'"

                                                        Next

                                                    End If

                                                    If buskEst <> "" Then buskEst = buskEst & " AND "


                                                    'otra forma:
                                                    pedX = ""
                                                    If CStr(xTabla.Rows(r).Item(8)) <> "" Then
                                                        'tiene caracter de concatenación, se hace split

                                                        If cadLeft.Contains(CStr(xTabla.Rows(r).Item(8))) = False Then

                                                            cadError = "Couldn't found the concatenation character required by rule: '" & CStr(xTabla.Rows(r).Item(8)) & "', please review!"

                                                            meQuedolaRegla = False

                                                            Exit For

                                                        End If

                                                        pedX = CStr(xTabla.Rows(r).Item(8))
                                                        miObjeto = Split(cadLeft, CStr(xTabla.Rows(r).Item(8)))
                                                        partCad = CStr(miObjeto(0))
                                                        partCad = partCad.Trim() 'se quitan los espacios en bblanco!
                                                    Else
                                                        'se toma el ancho de caracteres
                                                        If numChars <= cadLeft.Length Then
                                                            partCad = cadLeft.Substring(0, numChars) 'otra forma para lo mismo!
                                                        Else
                                                            If r = xTabla.Rows.Count - 1 Then
                                                                partCad = cadLeft
                                                            Else
                                                                'solo si NO es el último paso
                                                                If TengoCataMatch(catDs.Tables(mixNombreTab), cadLeft, buskEst, CStr(xObj(2)), partCad) = True Then

                                                                Else
                                                                    cadError = "The is no matches on the catalog: '" & catDs.Tables(mixNombreTab).ExtendedProperties.Item("CatalogName") & "' for the value '" & cadLeft & "', please check!"

                                                                    meQuedolaRegla = False

                                                                    Exit For
                                                                End If

                                                            End If
                                                        End If

                                                        'partCad = Mid(cadLeft, poSiclo, poSiclo + numChars - 1)
                                                    End If

                                                    If partCad.Length > numChars Then
                                                        'salimos
                                                        cadError = "The length for " & partCad & " doesn't meet the current required that is as maximum " & numChars & " characters length!"

                                                        meQuedolaRegla = False

                                                        Exit For

                                                    End If

                                                    buskEst = buskEst & CStr(xObj(2)) & " = '" & partCad & "'"

                                                    Dim fatRows() As DataRow
                                                    fatRows = catDs.Tables(mixNombreTab).Select(buskEst)

                                                    If fatRows.Length = 0 Then
                                                        '
                                                        cadError = "The is no matches on the catalog: '" & catDs.Tables(mixNombreTab).ExtendedProperties.Item("CatalogName") & "' for the value '" & partCad & "', please check!"

                                                        meQuedolaRegla = False

                                                        Exit For

                                                    End If

                                                    cadLeft = cadLeft.Substring(partCad.Length + pedX.Length)

                                                    longAcum = longAcum + partCad.Length + pedX.Length
                                                    charResta = valEvaluar.Length - longAcum


                                                Case Is = "B - Free Text"

                                                    pedX = ""

                                                    If r = xTabla.Rows.Count - 1 Then

                                                        partCad = cadLeft

                                                        If Len(partCad) <= CInt(numChars) Then

                                                            'ok

                                                            cadFix = cadFix & partCad

                                                            poSiclo = poSiclo + Len(partCad)

                                                            cadLeft = Mid(cadLeft, poSiclo, Len(cadLeft) - (poSiclo - 1))

                                                        Else

                                                            'error

                                                            cadError = "The length of " & partCad & " exceeds the limit of " & CInt(numChars) & " characters length!"

                                                            meQuedolaRegla = False

                                                            Exit For

                                                        End If

                                                    Else

                                                        pedX = ""

                                                        If CStr(xTabla.Rows(r).Item(8)) <> "" Then
                                                            'tiene caracter de concatenación, se hace split

                                                            If cadLeft.Contains(CStr(xTabla.Rows(r).Item(8))) = False Then

                                                                cadError = "Couldn't found the concatenation character required by rule: '" & CStr(xTabla.Rows(r).Item(8)) & "', please review!"

                                                                meQuedolaRegla = False

                                                                Exit For

                                                            End If

                                                            pedX = CStr(xTabla.Rows(r).Item(8))
                                                            miObjeto = Split(cadLeft, CStr(xTabla.Rows(r).Item(8)))
                                                            partCad = CStr(miObjeto(0))
                                                            partCad = partCad.Trim() 'se quitan los espacios en bblanco!
                                                        Else
                                                            'se toma el ancho de caracteres
                                                            'ir hacia adelante, tomar la 
                                                            'hay un caracter de concatenación missing!
                                                            If numChars <= cadLeft.Length Then
                                                                'se puede tomar
                                                                partCad = cadLeft.Substring(0, numChars)
                                                            Else
                                                                'NO se puede definir la cantidad de caracteres a tomar
                                                                cadError = "Unable to define the number of characters due to missing concatenation character please review with dev team or consultants!"

                                                                meQuedolaRegla = False

                                                                Exit For

                                                            End If
                                                            'otra forma para lo mismo!
                                                            'partCad = Mid(cadLeft, poSiclo, poSiclo + numChars - 1)
                                                        End If

                                                        If partCad.Length > numChars Then
                                                            'salimos
                                                            cadError = "The length for " & partCad & " doesn't meet the current required that is as maximum " & numChars & " characters length!"

                                                            meQuedolaRegla = False

                                                            Exit For

                                                        End If

                                                        'el valor que quea de partcad esta bien porque es free text!!

                                                        cadLeft = cadLeft.Substring(partCad.Length + pedX.Length)



                                                    End If



                                                Case Is = "C - Running number"

                                                    pedX = ""

                                                    If r = xTabla.Rows.Count - 1 Then

                                                        partCad = cadLeft

                                                        If IsNumeric(partCad) = True Then

                                                            'es numero

                                                            If partCad.Length = numChars Then
                                                                'cumple el largo de numeros
                                                                'debe cuadrarlo a la totalidad de caracteres que puso
                                                                poSiclo = poSiclo + Len(partCad) + Len(CStr(xTabla.Rows(r).Item(8)))

                                                                cadLeft = Mid(cadLeft, poSiclo, Len(cadLeft) - (poSiclo - 1))

                                                            Else
                                                                'NO lo cumple
                                                                cadError = "This part of the construction rule exceeds the maximum character limit : '" & CInt(numChars) & "', please check!"

                                                                meQuedolaRegla = False

                                                                Exit For

                                                            End If


                                                        Else

                                                            'No es numero!

                                                            cadError = "This part of the construction rule must be a valid running number, curently you have: '" & partCad & "', please check!"

                                                            meQuedolaRegla = False

                                                            Exit For

                                                        End If

                                                    Else

                                                        partCad = cadLeft.Substring(0, numChars)

                                                        If CStr(xTabla.Rows(r).Item(8)) <> "" Then

                                                            pedX = cadLeft.Substring(numChars, CStr(xTabla.Rows(r).Item(8)).Length)

                                                            If pedX <> CStr(xTabla.Rows(r).Item(8)) Then

                                                                'NO coincide!

                                                                cadError = "The concatenation character doesn't match the required by rule: '" & CStr(xTabla.Rows(r).Item(8)) & "', please check!"

                                                                meQuedolaRegla = False

                                                                Exit For

                                                            End If

                                                        End If


                                                        If IsNumeric(partCad) = True Then

                                                            If partCad.Length = numChars Then
                                                                'cumple el largo de numeros
                                                                'debe cuadrarlo a la totalidad de caracteres que puso
                                                                poSiclo = poSiclo + Len(partCad) + Len(CStr(xTabla.Rows(r).Item(8)))

                                                                cadLeft = Mid(cadLeft, poSiclo, Len(cadLeft) - (poSiclo - 1))

                                                            Else
                                                                'NO lo cumple
                                                                cadError = "This part of the construction rule exceeds the maximum character limit : '" & CInt(numChars) & "', please check!"

                                                                meQuedolaRegla = False

                                                                Exit For

                                                            End If

                                                        Else
                                                            cadError = "This part of the construction rule must be a valid running number, curently you have: '" & partCad & "', please check!"

                                                            meQuedolaRegla = False

                                                            Exit For

                                                        End If



                                                    End If



                                                Case Is = "D - Fixed Value"
                                                    'value
                                                    partCad = Mid(cadLeft, poSiclo, poSiclo + Len(CStr(xTabla.Rows(r).Item(6))) - 1) '

                                                    'esto se evalua!

                                                    If partCad = CStr(xTabla.Rows(r).Item(6)) Then

                                                        'ok!

                                                        If CStr(xTabla.Rows(r).Item(8)) = "" Then

                                                            cadLeft = Mid(cadLeft, poSiclo + Len(CStr(xTabla.Rows(r).Item(6))) + 1, Len(cadLeft) - (poSiclo + Len(CStr(xTabla.Rows(r).Item(6))) + 1))

                                                            poSiclo = poSiclo + Len(CStr(xTabla.Rows(r).Item(6)))

                                                        Else

                                                            partCad = Mid(cadLeft, poSiclo + Len(CStr(xTabla.Rows(r).Item(6))), Len(CStr(xTabla.Rows(r).Item(8))))

                                                            If partCad = CStr(xTabla.Rows(r).Item(8)) Then

                                                                'esta OK, fue igual al caracter próximo!

                                                                cadFix = cadFix & CStr(xTabla.Rows(r).Item(8))

                                                                poSiclo = poSiclo + Len(CStr(xTabla.Rows(r).Item(6))) + Len(CStr(xTabla.Rows(r).Item(8)))

                                                                cadLeft = Mid(cadLeft, poSiclo, Len(cadLeft) - (poSiclo - 1))


                                                            Else

                                                                'NO esta OK!

                                                                cadError = "The concatenation character doesn't match the required by rule: '" & CStr(xTabla.Rows(r).Item(8)) & "', please check!"

                                                                meQuedolaRegla = False

                                                                Exit For

                                                            End If

                                                        End If

                                                    Else

                                                        'No ok

                                                        'NO esta OK!
                                                        cadError = "The required Fixed value doesn't match the rule: '" & CStr(xTabla.Rows(r).Item(8)) & "', please check!"

                                                        meQuedolaRegla = False

                                                        Exit For

                                                    End If


                                                Case Is = "E - External object"

                                                    pedX = ""

                                                    xObj = Split(CStr(xTabla.Rows(r).Item(6)), ":")
                                                    'desde aqui, si el template code y el table code es el mismo, entonces va sobre su mismo renglon!!
                                                    'mm:md21:md21-0001:MATNR:MAKTX#MTBN-MAKNT#MATH:MATCH
                                                    mixNombreTab = CStr(xObj(1)) & "#" & CStr(xObj(2)) & "#" & CStr(xObj(3))

                                                    buskEst = ""

                                                    If CStr(xObj(4)) <> "None" And CStr(xObj(4)) <> "" Then
                                                        'se requiere validación match adicional!
                                                        'del campo mio al campo del objeto externo
                                                        yObj = Split(CStr(xObj(4)), "-")
                                                        For p = 0 To UBound(yObj)

                                                            If p <> 0 Then buskEst = buskEst & " AND "

                                                            zObj = Split(CStr(yObj(p)), "#")

                                                            buskEst = buskEst & CStr(zObj(1)) & " = '"

                                                            buskEst = buskEst & DataGridView1.Rows(i).Cells(CStr(zObj(0))).Value & "'"

                                                        Next

                                                    End If

                                                    If buskEst <> "" Then buskEst = buskEst & " AND "

                                                    If CStr(xTabla.Rows(r).Item(8)) <> "" Then
                                                        'se hace split!

                                                        If cadLeft.Contains(CStr(xTabla.Rows(r).Item(8))) = False Then

                                                            cadError = "Couldn't found the concatenation character required by rule: '" & CStr(xTabla.Rows(r).Item(8)) & "', please review!"

                                                            meQuedolaRegla = False

                                                            Exit For

                                                        End If

                                                        pedX = CStr(xTabla.Rows(r).Item(8))
                                                        miObjeto = Split(cadLeft, CStr(xTabla.Rows(r).Item(8)))
                                                        partCad = CStr(miObjeto(0))
                                                        partCad = partCad.Trim() 'se quitan los espacios en bblanco!

                                                    Else

                                                        If r = xTabla.Rows.Count - 1 Then
                                                            partCad = cadLeft
                                                        Else
                                                            'solo si NO es el último paso
                                                            If CStr(xObj(1)) = objetoSelek And CStr(xObj(2)) = tableSelek Then
                                                                'que alguna parte del objeto cuadre
                                                                'Aquii falta una funcion para buscar dentro de cadleft un valor que le cuadre al datagridview.cells.
                                                                'DataGridView1.Rows(i).Cells(CStr(xObj(3))).Value'esta es la 
                                                                'cadleft es la cadena en la que se va a buscar
                                                                'Datagridview es ciclo de cadena a buscar!

                                                                If HayCadenainsideCadena(cadLeft, DataGridView1.Rows(i).Cells(CStr(xObj(3))).Value, partCad) = False Then

                                                                    cadError = "The is no match on the referenced column: '" & DataGridView1.Columns(CStr(xObj(3))).HeaderText & "' inside the string: '" & cadLeft & "', please check!, there should be a match with this string: " & DataGridView1.Rows(i).Cells(CStr(xObj(3))).Value

                                                                    meQuedolaRegla = False

                                                                    Exit For

                                                                End If


                                                            Else

                                                                If TengoCataMatch(multiDepe.Tables(mixNombreTab), cadLeft, buskEst, CStr(xObj(3)), partCad) = False Then

                                                                    cadError = "The is no matches on the External object>: '" & CStr(xObj(1)) & " > " & CStr(xObj(2)) & " > " & CStr(xObj(3)) & " for the value '" & cadLeft & "', please check!"

                                                                    meQuedolaRegla = False

                                                                    Exit For

                                                                End If

                                                            End If

                                                        End If

                                                    End If


                                                    If CStr(xObj(1)) = objetoSelek And CStr(xObj(2)) = tableSelek Then
                                                        'es local, misma template, misma hoja
                                                        'se supone que SIEMPRE sera con None en match fields
                                                        If partCad <> DataGridView1.Rows(i).Cells(CStr(xObj(3))).Value Then

                                                            cadError = "There is no matchs on the external object required > " & CStr(xObj(1)) & " > " & CStr(xObj(2)) & " > " & CStr(xObj(3))

                                                            meQuedolaRegla = False

                                                            Exit For

                                                        End If

                                                    Else
                                                        '

                                                        If multiDepe.Tables.IndexOf(mixNombreTab) < 0 Then

                                                            cadError = "The required object to evaluate this field is not available or has not yet builded: '" & CStr(xObj(1)) & " > " & CStr(xObj(2)) & " > " & CStr(xObj(3)) & ""

                                                            meQuedolaRegla = False

                                                            Exit For

                                                        End If

                                                        buskEst = buskEst & CStr(xObj(3)) & " = '" & partCad & "'"

                                                        Dim fatRows() As DataRow
                                                        fatRows = multiDepe.Tables(mixNombreTab).Select(buskEst)

                                                        If fatRows.Length = 0 Then
                                                            '
                                                            cadError = "There is no matchs on the external object required > " & CStr(xObj(1)) & " > " & CStr(xObj(2)) & " > " & CStr(xObj(3))

                                                            meQuedolaRegla = False

                                                            Exit For

                                                        End If

                                                    End If

                                                    'si llega aqui entonces SI paso, y debe redefinirse cadleft



                                                    'cadLeft = Mid(cadLeft, poSiclo + numChars + pedX.Length, Len(cadLeft) - (poSiclo - 1))
                                                    cadLeft = cadLeft.Substring(partCad.Length + pedX.Length)
                                                    'Dim xTabF As DataTable
                                                    'xTabF = multiDepe.Tables(mixNombreTab).Clone()
                                                    'For Each row In fatRows
                                                    '    xTabF.ImportRow(row)
                                                    'Next




                                            End Select

                                        Next

                                        'meQuedolaRegla = True 'este no va!

                                        If meQuedolaRegla = True Then Exit For

                                    Next


                                    If meQuedolaRegla = False Then
                                        PintaCeldaDeError(i, j, cadError)
                                        Continue For
                                    End If

                                End If


                            Case Is = "D - Fixed Value"
                                If valEvaluar <> CStr(ValidaDt.Rows(z).Item(16)) Then
                                    PintaCeldaDeError(i, j, "This is a fixed value, and must be " & CStr(ValidaDt.Rows(z).Item(16)))
                                    Continue For
                                End If

                            Case Is = "E - Simple validation"
                                'nada, porque ya se valido arriba!


                            Case Is = "F - Against conditions"


                        End Select


                        If CStr(ValidaDt.Rows(z).Item(18)) <> "" Then
                            'if non allowed charss, tiene algo, debe validar que la cadena que se puso NO contenga caracteres inválidos!
                            If ContainsInvalidChars(valEvaluar, CStr(ValidaDt.Rows(z).Item(18))) = True Then
                                PintaCeldaDeError(i, j, "This field is not allowed to contain any of the following characters: " & CStr(ValidaDt.Rows(z).Item(18)) & " , please review!")
                                Continue For
                            End If
                        End If


                        Select Case CStr(ValidaDt.Rows(z).Item(12))
                            Case Is = "None", ""
                                'nada!

                            Case Is = "UCASE"
                                valEvaluar = valEvaluar.ToUpper()
                                DataGridView1.Rows(i).Cells(j).Value = valEvaluar

                            Case Is = "lcase"
                                valEvaluar = valEvaluar.ToLower()
                                DataGridView1.Rows(i).Cells(j).Value = valEvaluar

                        End Select

                        anchoTru(j) = True

                    Next

                    If k > 0 Then hayYave = True Else hayYave = False

                    For j = 0 To DataGridView1.Columns.Count - 4
                        If anchoTru(j) = False Then
                            elRowOk = False
                            Exit For
                        End If
                    Next

                    If elRowOk = True Then
                        DataGridView1.Rows(i).Tag = "Ok!"
                        DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 1 - 2).Value = "Ok!"
                        DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 1 - 2).Style.ForeColor = Color.DarkGreen
                        DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 1 - 2).Style.Font = New System.Drawing.Font("Calibri", 10, FontStyle.Bold)
                    Else
                        DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 1 - 2).Value = "NO OK"
                        DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 1 - 2).Style.ForeColor = Color.Crimson
                        DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 1 - 2).Style.Font = New System.Drawing.Font("Calibri", 10, FontStyle.Bold)
                    End If

                    If hayYave = True Then
                        Try
                            xDt.Rows.Add({concaYave})
                            DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 2).Value = "Ok. No repeteability detected on key field"
                            DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 2).Style.ForeColor = Color.DarkGreen
                            DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 2).Style.Font = New System.Drawing.Font("Calibri", 10, FontStyle.Bold)
                            DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 2).Tag = "Ok. No repeteability detected on key field"
                        Catch ex As Exception

                            'PintaCeldaDeError(i, j, "")
                            findAnother = xDt.Rows.Find(concaYave)
                            posOtro = xDt.Rows.IndexOf(findAnother)
                            If posOtro >= 0 Then
                                'aqui
                                'la celda de repetibilidad
                                DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 2).Value = "NO OK. Key field repeteability detected on row " & posOtro
                                DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 2).Style.ForeColor = Color.Crimson
                                DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 2).Style.Font = New System.Drawing.Font("Calibri", 10, FontStyle.Bold)
                                DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 2).Tag = "NO OK. Key field repeteability detected on row " & posOtro
                            End If
                        End Try
                    Else

                        DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 2).Value = "Ok. No repeteability is configured for this template"
                        DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 2).Style.ForeColor = Color.Black
                        DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 2).Style.Font = New System.Drawing.Font("Calibri", 10, FontStyle.Bold)
                        DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 2).Tag = "Ok. No repeteability is configured for this template"
                    End If

                    'se evalua!
                    If filtRelas.Rows.Count = 0 Then
                        'se pone que OK!
                        DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 1).Value = "No internal relations detected"
                        DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 1).Style.ForeColor = Color.DarkGreen
                        DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 1).Style.Font = New System.Drawing.Font("Calibri", 10, FontStyle.Bold)
                        DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 1).Tag = "No internal relations detected"
                        Continue For
                    Else
                        cuenAll = 0
                        cuenOk = 0
                        cuenBad = 0

                        expText = ""

                        For j = 0 To relaUniqRul.Rows.Count - 1

                            yaveInte = ""
                            mixKeys = ""
                            xLlena.Clear()
                            xLlena = filtRelas.Clone()
                            Dim resy() As DataRow = filtRelas.Select("RuleCode='" & CStr(relaUniqRul.Rows(j).Item(0)) & "'", "Consec ASC")
                            For Each row As DataRow In resy
                                xLlena.ImportRow(row)
                            Next

                            For w = 0 To xLlena.Rows.Count - 1
                                If yaveInte = "" Then
                                    yaveInte = DataGridView1.Rows(i).Cells(CStr(xLlena.Rows(w).Item(4))).Value
                                Else
                                    yaveInte = yaveInte & "#" & DataGridView1.Rows(i).Cells(CStr(xLlena.Rows(w).Item(4))).Value
                                End If

                                If mixKeys = "" Then
                                    mixKeys = CStr(xLlena.Rows(w).Item(4))
                                Else
                                    mixKeys = mixKeys & " > " & CStr(xLlena.Rows(w).Item(4))
                                End If

                            Next

                            'ciclamos por la regla para obtener la concatenacion del datagrid!
                            'vemos si la contiene!
                            enCuentra = internRelation.Tables(CStr(relaUniqRul.Rows(j).Item(0))).Rows.Find(yaveInte)
                            unComodin = internRelation.Tables(CStr(relaUniqRul.Rows(j).Item(0))).ExtendedProperties.Item("TableCode")

                            If expText <> "" Then expText = expText & vbCrLf
                            cuenAll = cuenAll + 1

                            If IsNothing(enCuentra) = True Then
                                'NO esta!
                                cuenBad = cuenBad + 1
                                expText = expText & "1 Relation could NOT be verified correctly > " & "Relation with > " & unComodin & " > " & mixKeys & " to value > " & yaveInte & " not verified on sibling table. Please review"

                            Else
                                'SISTA!
                                cuenOk = cuenOk + 1
                                expText = expText & "1 Relation verified correctly > " & "Relation with > " & unComodin & " > " & mixKeys & " to value > " & yaveInte & " verified on sibling table!"

                            End If

                        Next

                        DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 1).Tag = expText

                        If cuenOk = cuenAll Then
                            DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 1).Value = "All relations on sibling(s) table(s) verified correctly!"
                            DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 1).Style.ForeColor = Color.DarkGreen
                            DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 1).Style.Font = New System.Drawing.Font("Calibri", 10, FontStyle.Bold)
                        Else
                            If cuenBad = cuenAll Then
                                DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 1).Value = "NO OK. The relations on sibling(s) table(s) could NOT verified, please review"
                                DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 1).Style.ForeColor = Color.Crimson
                                DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 1).Style.Font = New System.Drawing.Font("Calibri", 10, FontStyle.Bold)
                            Else
                                DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 1).Value = "Warning. Some relations could not be verified. Double click on this cell for details "
                                DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 1).Style.ForeColor = Color.Goldenrod
                                DataGridView1.Rows(i).Cells(DataGridView1.Columns.Count - 1).Style.Font = New System.Drawing.Font("Calibri", 10, FontStyle.Bold)

                            End If
                        End If

                    End If

                Next

                ToolStripLabel1.Text = "Ready"
                ToolStripLabel1.Visible = False
                DoneEvaluacion = True
                toyTrabajando = False


            Case Is = 4



            Case Is = 5


        End Select



    End Sub

    Private Sub PintaCeldaDeWarning(ByVal rowInd As Integer, ByVal colInd As Integer, ByVal xError As String)
        DataGridView1.Rows(rowInd).Cells(colInd).Style.BackColor = Color.FromArgb(250, 173, 20)
        DataGridView1.Rows(rowInd).Cells(colInd).Style.ForeColor = Color.White
        DataGridView1.Rows(rowInd).Cells(colInd).ToolTipText = xError
        DataGridView1.Rows(rowInd).Cells(colInd).Tag = xError.Replace(",", "-")
    End Sub

    Private Sub PintaCeldaDeError(ByVal rowInd As Integer, ByVal colInd As Integer, ByVal xError As String)
        DataGridView1.Rows(rowInd).Cells(colInd).Style.BackColor = Color.Crimson
        DataGridView1.Rows(rowInd).Cells(colInd).Style.ForeColor = Color.White
        DataGridView1.Rows(rowInd).Cells(colInd).ToolTipText = xError
        DataGridView1.Rows(rowInd).Cells(colInd).Tag = xError.Replace(",", "-")
    End Sub

    Private Sub LimpiaCeldaDeError(ByVal rowInd As Integer, ByVal colInd As Integer)

        DataGridView1.Rows(rowInd).Cells(colInd).Style.BackColor = Color.White
        DataGridView1.Rows(rowInd).Cells(colInd).Style.ForeColor = Color.Black
        DataGridView1.Rows(rowInd).Cells(colInd).ToolTipText = ""
        DataGridView1.Rows(rowInd).Cells(colInd).Tag = "Ok!"
    End Sub

    Private Async Sub DataGridView1_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEnter

        If estoyAgregandoRows = True Then Exit Sub

        If rowEstaba > DataGridView1.Rows.Count - 1 Then Exit Sub

        If e.RowIndex < 0 Then Exit Sub

        If rowEstaba = e.RowIndex Then Exit Sub

        If rowEstaba < 0 Then
            rowEstaba = e.RowIndex
            Exit Sub
        End If

        Dim xObj As Object
        Dim unPaths As String
        Dim cadFind As String = ""
        Dim moduSel As Object = Nothing
        Dim tabCode As String
        Dim tabNombre As String = ""
        Dim xResp As String = ""

        Select Case ToolStripComboBox1.SelectedIndex

            Case Is = 0
                'nada

            Case Is = 1



            Case Is = 2



            Case Is = 3



            Case Is = 4
                'templates!

                xObj = Split(elNode.FullPath, "\")

                If xObj.Length = 1 Or xObj.Length = 2 Then

                    Exit Sub
                End If

                writeDs.Tables(0).Rows.Clear() 'tabla de los niveles
                writeDs.Tables(1).Rows.Clear() 'nuevos
                writeDs.Tables(2).Rows.Clear() 'deleted
                writeDs.Tables(3).Rows.Clear() 'current

                writeDs.Tables(0).Rows.Add({"templates"}) 'nivel 1

                unPaths = RaizFire
                unPaths = unPaths & "/templates"

                Select Case xObj.Length
                    Case Is = 2
                        cadFind = elNode.Name & "#" & elNode.Tag
                        writeDs.Tables(0).Rows.Add({elNode.Name}) 'nivel 2

                    Case Is = 3
                        moduSel = Split(elNode.Parent.Tag, "#")
                        tabCode = elNode.Name
                        tabNombre = elNode.Tag
                        cadFind = elNode.Parent.Name & "#" & elNode.Parent.Tag
                        writeDs.Tables(0).Rows.Add({elNode.Parent.Name})
                        writeDs.Tables(0).Rows.Add({tabCode})

                        unPaths = unPaths & "/" & elNode.Parent.Name
                        unPaths = unPaths & "/" & tabCode

                    Case Is = 4
                        moduSel = Split(elNode.Parent.Parent.Tag, "#")
                        tabCode = elNode.Parent.Name
                        tabNombre = elNode.Parent.Tag
                        cadFind = elNode.Parent.Parent.Name & "#" & elNode.Parent.Parent.Tag
                        writeDs.Tables(0).Rows.Add({elNode.Parent.Parent.Name})
                        writeDs.Tables(0).Rows.Add({tabCode})

                        unPaths = unPaths & "/" & elNode.Parent.Parent.Name
                        unPaths = unPaths & "/" & tabCode

                End Select

                Dim enCuentra As DataRow
                enCuentra = ModuPermit.Tables(0).Rows.Find(CStr(moduSel(1)).ToUpper())
                If IsNothing(enCuentra) = True Then
                    Exit Sub
                End If

                If puedoEditar = False Then Exit Sub 'NO puede editar!

                If IsNothing(DataGridView1.Rows(rowEstaba).Cells(0).Value) = True Then Exit Sub
                If DataGridView1.Rows(rowEstaba).Cells(0).Value = "" Then Exit Sub

                Call TomaInfoPara1RowTemplate(DataGridView1, writeDs, 3, 0, 3, 13, rowEstaba)

                'tomamos la info!!, 
                'Cursor.Current = Cursors.WaitCursor
                ToolStripLabel1.Visible = True
                ToolStripLabel1.Text = "Updating row..." & CStr(rowEstaba + 1)

                Await HazPutEnFbSimple(unPaths, "TableName", tabNombre)

                xResp = Await HazPutEnFireBasePathYColumnas(unPaths, writeDs.Tables(3), 0)

                Await SigoVivo()
                Module5.AgregaTiempo()

                ToolStripLabel1.Text = "Ready"
                ToolStripLabel1.Visible = False


            Case Is = 5



        End Select


        rowEstaba = e.RowIndex

    End Sub

    Private Async Sub ToolStripButton14_Click(sender As Object, e As EventArgs) Handles ToolStripButton14.Click

        Dim x As Integer = 0
        Dim elCamino As String = ""
        Dim unResp As String = ""
        Select Case ToolStripComboBox1.SelectedIndex
            Case Is = 0

            Case Is = 1


            Case Is = 2


            Case Is = 3


            Case Is = 4
                'templates!
                'primero, debe estar seleccionado el combo de conditional!

                Dim enCuentra As DataRow
                enCuentra = ModuPermit.Tables(0).Rows.Find(moduloSelek.ToUpper())
                If IsNothing(enCuentra) = True Then
                    MsgBox("Sorry you are not allowed to add catalogs on the selected module", vbCritical, TitBox)
                    Exit Sub
                End If

                If puedoEditar = False Then
                    MsgBox("Sorry, you are not allowed to edit this template at this time, please try again later or when it becomes available.", vbExclamation, TitBox)
                    Exit Sub
                End If

                Await SigoVivo()
                Module5.AgregaTiempo()

                If DataGridView1.CurrentCell.RowIndex < 0 Then Exit Sub

                If IsNothing(DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(0).Value) = True Then
                    MsgBox("Please define a field code first!", vbCritical, TitBox)
                    Exit Sub
                End If

                If DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(0).Value = "" Then
                    MsgBox("Please define a field code first!", vbCritical, TitBox)
                    Exit Sub
                End If

                If DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(3).Value <> "Conditional" Then
                    MsgBox("Please select 'Conditional' type in the MOC column!", vbCritical, TitBox)
                    Exit Sub
                End If

                x = MsgBox("ATTENTION" & vbCrLf & "Are you sure you want to delete this dependencie?. This action can not be undone!", vbQuestion + vbYesNo, TitBox)

                If x <> 6 Then Exit Sub

                unResp = Await BorraDepen(objetoSelek, tableSelek, CStr(DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(0).Value))

                If unResp = "Ok" Then
                    'hacer reload!
                    usaDataset.Tables(0).Rows.Clear()
                    usaDataset.Tables(0).Rows.Add({"dependencies"})
                    depeDs.Tables.Clear()
                    depeDs = Await PullUrlWs(usaDataset, "dependencies")

                    usaDataset.Tables(0).Rows.Clear()
                    usaDataset.Tables(0).Rows.Add({"templates"})
                    tempDs.Tables.Clear()
                    tempDs = Await PullUrlWs(usaDataset, "templates")

                    DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(14).Value = ""
                    DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(15).Value = ""
                    DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(16).Value = ""
                    DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(17).Value = ""
                    DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(18).Value = ""
                    DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(19).Value = ""
                    DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(20).Value = ""

                End If



            Case Is = 5





        End Select

    End Sub

    Public Async Function BorraCampo(ByVal elObjeto As String, ByVal laTabla As String, ByVal elCampo As String) As Task(Of String)

        'OJO, con las implicaciones de borrar un campo!!
        Dim elCamino As String = ""
        elCamino = RaizFire
        elCamino = elCamino & "/templates/" & elObjeto & "/" & laTabla

        Return Await HazDeleteEnFbSimple(elCamino, elCampo)

    End Function

    Private Async Function BorraDepen(ByVal elObjeto As String, ByVal laTabla As String, ByVal elCampo As String) As Task(Of String)
        'im laResp As String = ""
        Dim elCamino As String = ""
        elCamino = RaizFire
        elCamino = elCamino & "/dependencies/" & elObjeto & "/" & laTabla

        Return Await HazDeleteEnFbSimple(elCamino, elCampo)

    End Function

    Private Async Sub ToolStripButton15_Click(sender As Object, e As EventArgs) Handles ToolStripButton15.Click





        Await SigoVivo()
        Module5.AgregaTiempo()

        Form8.xTempDs = tempDs
        Form8.queObjetoEs = objetoSelek
        Form8.relaDs = relaTionDs

        Form8.ShowDialog()

        usaDataset.Tables(0).Rows.Clear()
        usaDataset.Tables(0).Rows.Add({"relations"})
        relaTionDs.Tables.Clear()
        relaTionDs = Await PullUrlWs(usaDataset, "relations")

    End Sub

    Private Async Sub JalaRegistrosHermanos(ByVal laCompany As String)
        'OJO, en este SOLO se van a jalar los registros y se va a construir una columna yave con la concatenación de las N columnas yave que lo conformen!!
        'Con 
        'la ParentTable te va a decir la info!
        'para comparar s edebe filtrar por ChildTableCode, y por rulecode
        'se deben sacar las rules unicas!
        'la reglaUnica
        'se puede filtrar primero por childTableCode y luego sacar los Únicos de parentTableCode y con eso hacer la búsqueda!
        'si llegara a no tener entonces no hay nada que validar!
        internRecords.Tables.Clear()
        internRelation.Tables.Clear()

        Dim poSi As Integer = 0
        poSi = relaTionDs.Tables.IndexOf(objetoSelek)

        filtRelas.Clear()
        filtRelas.Columns.Clear()
        filtRelas.Rows.Clear()

        If poSi >= 0 Then

            filtRelas = relaTionDs.Tables(poSi).Clone()

            Dim miFil As New DataTable
            Dim xDs As New DataSet
            Dim result() As DataRow = relaTionDs.Tables(poSi).Select("ChildTableCode='" & tableSelek & "'")
            Dim uniDt As New DataTable
            'Dim uniqRul As New DataTable
            relaUniqRul.Clear()

            For Each row As DataRow In result
                filtRelas.ImportRow(row)
            Next

            uniDt = filtRelas.DefaultView.ToTable(True, "ParentTableCode")
            relaUniqRul = filtRelas.DefaultView.ToTable(True, "RuleCode")

            For i = 0 To uniDt.Rows.Count - 1
                'obtenemos los registros únicos!
                'y creamos la tabla única!

                'y jalamos la info!
                xDs.Tables.Clear()
                usaDataset.Tables(0).Rows.Clear()
                usaDataset.Tables(0).Rows.Add({"records"})
                usaDataset.Tables(0).Rows.Add({laCompany})
                usaDataset.Tables(0).Rows.Add({objetoSelek})
                usaDataset.Tables(0).Rows.Add({CStr(uniDt.Rows(i).Item(0))})
                usaDataset.Tables(0).Rows.Add({"records"})
                xDs = Await PullUrlWs(usaDataset, "records")

                'se construye
                If xDs.Tables.Count > 0 Then
                    'se copia!
                    internRecords.Tables.Add(xDs.Tables(0).Copy())
                    internRecords.Tables(internRecords.Tables.Count - 1).TableName = CStr(uniDt.Rows(i).Item(0))
                Else
                    internRecords.Tables.Add(CStr(uniDt.Rows(i).Item(0)))
                End If

            Next

            'Dim tablaPapName As String = ""
            Dim tablaPapa As String = ""
            Dim posPapa As Integer = 0
            Dim enCuentra As DataRow
            'filtramos vs todas las ruleCode y ordenamos de menor a mayor, 
            For i = 0 To relaUniqRul.Rows.Count - 1

                tablaPapa = ""

                miFil.Clear()
                miFil = filtRelas.Clone()
                Dim resy() As DataRow = filtRelas.Select("RuleCode='" & CStr(relaUniqRul.Rows(i).Item(0)) & "' AND ParentField<>''", "Consec ASC")
                For Each row As DataRow In resy
                    miFil.ImportRow(row)
                    tablaPapa = row.Item(3)
                Next

                'vamos a construir el mismo numero de tablas a la par al mismo numero de reglas construidas!
                internRelation.Tables.Add(CStr(relaUniqRul.Rows(i).Item(0)))
                'le construimos su campo yave!
                Call AsignaYavePrimariaATabla(internRelation.Tables(internRelation.Tables.Count - 1), CStr(relaUniqRul.Rows(i).Item(0)), True)

                If tablaPapa = "" Then Continue For

                posPapa = internRecords.Tables.IndexOf(tablaPapa)
                internRelation.Tables(internRelation.Tables.Count - 1).ExtendedProperties.Add("TableCode", tablaPapa)

                If posPapa >= 0 Then
                    'ciclamos por toda la tabla y concatenamos y agregamos a la tabla de internrelation!
                    Dim Conca As String = ""
                    For j = 0 To internRecords.Tables(posPapa).Rows.Count - 1
                        Conca = ""
                        'sacamos el campo yave!
                        For w = 0 To miFil.Rows.Count - 1
                            'se toma el valor
                            If Conca = "" Then
                                Conca = internRecords.Tables(posPapa).Rows(j).Item(CStr(miFil.Rows(w).Item(6)))
                            Else
                                Conca = Conca & "#" & internRecords.Tables(posPapa).Rows(j).Item(CStr(miFil.Rows(w).Item(6)))
                            End If
                        Next

                        enCuentra = internRelation.Tables(internRelation.Tables.Count - 1).Rows.Find(CStr(Conca))

                        If IsNothing(enCuentra) = False Then Continue For 'significa que ya estaba en la coleccion!

                        internRelation.Tables(internRelation.Tables.Count - 1).Rows.Add({Conca})

                    Next

                End If

            Next


        End If



    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseDoubleClick


        Select Case ToolStripComboBox1.SelectedIndex
            Case Is = 0
                'nada

            Case Is = 1
                'catalogos!

            Case Is = 2


            Case Is = 3
                'records!


            Case Is = 4
                'templates!

            Case Is = 5


        End Select


    End Sub

    Private Async Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click



    End Sub

    Private Async Sub ToolStripButton16_Click(sender As Object, e As EventArgs) Handles ToolStripButton16.Click

        Dim xObj As Object = Nothing
        Dim moduCode As String = ""
        Dim mitablaCode As String = ""
        Select Case ToolStripComboBox1.SelectedIndex
            Case Is = 0
                'nada!

            Case Is = 1
                'catalogos!
                If puedoEditar = False Then
                    MsgBox("Sorry, you are not allowed to edit this catalog at this time, please try again later or wait for it to be available to edit.", vbExclamation, TitBox)
                    Exit Sub
                End If


                If puSSyCat < 0 Then
                    MsgBox("Please select a catalog first!!", vbCritical, TitBox)
                    Exit Sub
                End If

                'moduCode = elNode.Parent.Name
                'mitablaCode = elNode.Name
                Dim enCuentra As DataRow
                enCuentra = ModuPermit.Tables(0).Rows.Find(CStr(catDs.Tables(cataNombre).ExtendedProperties.Item("ModuleCode")).ToUpper())
                If IsNothing(enCuentra) = True Then
                    MsgBox("Sorry you are not allowed to add catalogs on the selected module", vbCritical, TitBox)
                    Exit Sub
                End If

                Await SigoVivo()
                Module5.AgregaTiempo()

                'copiamos la ESTRUCTURA de información!
                TabColumnas.Rows.Clear()
                niuTabcols.Rows.Clear()
                For i = 2 To catDs.Tables(cataNombre).Columns.Count - 2 'desde la columna yave +1 hasta la ultima -1, quitando el campo de Descripción Final!
                    'solo desde la columna de los datos
                    TabColumnas.Rows.Add()
                    TabColumnas.Rows(TabColumnas.Rows.Count - 1).Item(0) = catDs.Tables(cataNombre).Columns(i).ColumnName
                    TabColumnas.Rows(TabColumnas.Rows.Count - 1).Item(1) = catDs.Tables(cataNombre).Columns(i).ExtendedProperties.Item("FieldName")
                    TabColumnas.Rows(TabColumnas.Rows.Count - 1).Item(2) = catDs.Tables(cataNombre).Columns(i).ExtendedProperties.Item("Description")
                    TabColumnas.Rows(TabColumnas.Rows.Count - 1).Item(3) = catDs.Tables(cataNombre).Columns(i).ExtendedProperties.Item("Position")
                    TabColumnas.Rows(TabColumnas.Rows.Count - 1).Item(4) = catDs.Tables(cataNombre).Columns(i).ExtendedProperties.Item("isText")
                Next

                Form10.moduObjeto = elNode.Parent.Name
                Form10.tablaCodigo = elNode.Name
                Form10.huboCambio = False
                Form10.ShowDialog()

                If Form10.huboCambio = False Then Exit Sub

                'Aquii, se debe hacer reload de esa tabla ó reload completo de catds, y luego re-escribir la estructura del template
                usaDataset.Tables(0).Rows.Clear()
                usaDataset.Tables(0).Rows.Add({"catpro"})
                catDs.Tables.Clear()
                catDs = Await PullUrlWs(usaDataset, "catpro")
                'aquii falta el sub para cargar los titulos del datagridview!
                ReloadCatPro()
                ReloadTempStruct()
                'por último se hace el reload de la estructura!


            Case Is = 2
                'depe

            Case Is = 3
                'record

            Case Is = 4
                'templates

            Case Is = 5
                'tree dep


        End Select

    End Sub

    Private Sub ReloadTempStruct()

        If cataNombre = "" Then Exit Sub

        DataGridView1.DataSource = Nothing
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()

        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AllowUserToDeleteRows = False

        For i = 0 To DataGridView1.Columns.Count - 1
            DataGridView1.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        Next

        'ahora primero agregamos las columnas!
        For j = 2 To catDs.Tables(cataNombre).Columns.Count - 1
            DataGridView1.Columns.Add(catDs.Tables(cataNombre).Columns(j).ColumnName, catDs.Tables(cataNombre).Columns(j).ExtendedProperties.Item("FieldName"))
            DataGridView1.Columns(DataGridView1.Columns.Count - 1).DataPropertyName = catDs.Tables(cataNombre).Columns(j).ColumnName
        Next

        For i = 0 To catDs.Tables(cataNombre).Rows.Count - 1
            DataGridView1.Rows.Add()
            DataGridView1.Rows(DataGridView1.Rows.Count - 1).HeaderCell.Value = CStr(i + 1)

            For j = 0 To DataGridView1.Columns.Count - 1
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(j).Value = catDs.Tables(cataNombre).Rows(i).Item(DataGridView1.Columns(j).Name)
            Next

            'For j = 0 To catDs.Tables(Posi).Columns.Count - 1
            'DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(j).Value = catDs.Tables(Posi).Rows(i).Item(j)
            'Next

            'DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = catDs.Tables(Posi).Rows(j).Item(0)
            'DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = catDs.Tables(Posi).Rows(j).Item(1)
        Next

        DataGridView1.RowHeadersWidth = 70

        For i = 0 To DataGridView1.Columns.Count - 1
            DataGridView1.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next

        DataGridView1.AllowUserToAddRows = puedoEditar ' True
        DataGridView1.AllowUserToDeleteRows = puedoEditar ' True

    End Sub

    Private Async Sub ToolStripButton17_Click(sender As Object, e As EventArgs) Handles ToolStripButton17.Click
        'boton para hacer refresh de los registros guardados!
        If CategSelected <> 3 Then Exit Sub
        'bajar info de firebase
        If IsNothing(elNode) = True Then
            MsgBox("Please select a valid node!", vbCritical, TitBox)
            Exit Sub
        End If

        Dim xObj As Object
        xObj = Split(elNode.FullPath, "\")

        If xObj.Length < 4 Then
            MsgBox("Please select a node at a table level!", vbCritical, TitBox)
            Exit Sub
        End If

        Dim elCamino As String = RaizFire
        elCamino = elCamino & "/records/" & elNode.Parent.Parent.Name 'compania
        elCamino = elCamino & "/" & elNode.Parent.Name 'objeto
        elCamino = elCamino & "/" & elNode.Name 'la tabla
        elCamino = elCamino & "/records"

        usaDataset.Tables(0).Rows.Clear()
        usaDataset.Tables(0).Rows.Add({"records"})
        usaDataset.Tables(0).Rows.Add({elNode.Parent.Parent.Name})
        usaDataset.Tables(0).Rows.Add({elNode.Parent.Name})
        usaDataset.Tables(0).Rows.Add({elNode.Name})
        usaDataset.Tables(0).Rows.Add({"records"})

        Dim xDs As New DataSet
        xDs = Await PullUrlWs(usaDataset, "records")

        DataGridView1.DataSource = Nothing
        DataGridView1.Rows.Clear()

        DataGridView1.AllowUserToAddRows = False
        DataGridView1.AllowUserToDeleteRows = False

        If xDs.Tables.Count > 0 Then
            Dim k As Integer = 0
            For i = 0 To xDs.Tables(0).Rows.Count - 1

                DataGridView1.Rows.Add()
                DataGridView1.Rows(DataGridView1.Rows.Count - 1).HeaderCell.Value = CStr(i + 1)
                For j = 0 To xDs.Tables(0).Columns.Count - 1
                    DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(xDs.Tables(0).Columns(j).ColumnName).Value = xDs.Tables(0).Rows(i).Item(j)
                Next

            Next

        Else
            MsgBox("No records found for this object!!", vbInformation, TitBox)
        End If

        DataGridView1.AllowUserToAddRows = puedoEditar ' True
        DataGridView1.AllowUserToDeleteRows = puedoEditar ' True
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        If ToolStripTextBox1.Text = "" Then
            MsgBox("Please type some info to search", vbCritical, TitBox)
            Exit Sub
        End If

        elQbusca = ToolStripTextBox1.Text
        buscaEnlower = ToolStripTextBox1.Text.ToLower()

        Select Case ToolStripComboBox1.SelectedIndex
            Case Is = 0
                'Dim NodoMadre As TreeNode
                'NodoMadre = TreeView1.Nodes(0)

                'euReka = False
                'oNode = Nothing
                'Call buSkaXNodo(NodoMadre)

            Case Is = 1, 2
                For i = 0 To TreeView1.Nodes.Count - 1
                    euReka = False
                    oNode = Nothing
                    Call buSkaXNodo(TreeView1.Nodes(i))
                    If euReka = True Then Exit For
                Next

        End Select


        If euReka = False Then
            MsgBox("Node not found!!", vbCritical, TitBox)
            Exit Sub
        End If

        oNode.BackColor = Color.Khaki
        oNode.EnsureVisible()
        oNode.Expand()
    End Sub

    Private Sub ToolStripTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ToolStripTextBox1.KeyDown
        If e.KeyCode = Keys.Return Then
            ToolStripButton3.PerformClick()
        End If
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        'save all

    End Sub

    Private Sub ToolStripButton9_CheckedChanged(sender As Object, e As EventArgs) Handles ToolStripButton9.CheckedChanged

        If ToolStripComboBox1.SelectedIndex <> 3 Then Exit Sub

        If DataGridView1.DataSource Is Nothing = False Then
            'tiene algo en el binding
            Dim currencyManager1 As CurrencyManager = CType(BindingContext(DataGridView1.DataSource), CurrencyManager)
            currencyManager1.SuspendBinding()

            If ToolStripButton9.Checked = True Then

                For i = 0 To DataGridView1.Rows.Count - 2
                    If IsNothing(DataGridView1.Rows(i).Tag) = True Then Continue For
                    If DataGridView1.Rows(i).Tag = "" Then Continue For
                    If DataGridView1.Rows(i).Tag = "NO OK" Then
                        DataGridView1.Rows(i).Visible = True
                    Else
                        DataGridView1.Rows(i).Visible = False
                    End If
                Next

            Else
                If ToolStripButton9.Checked = False Then
                    For i = 0 To DataGridView1.Rows.Count - 2
                        If IsNothing(DataGridView1.Rows(i).Tag) = True Then Continue For
                        If DataGridView1.Rows(i).Tag = "" Then Continue For
                        DataGridView1.Rows(i).Visible = True
                    Next
                End If
            End If

            currencyManager1.ResumeBinding()

        Else
            'NO tiene nada!
            For i = 0 To DataGridView1.Rows.Count - 2
                If IsNothing(DataGridView1.Rows(i).Tag) = True Then Continue For
                If DataGridView1.Rows(i).Tag = "" Then Continue For
                If DataGridView1.Rows(i).Tag = "NO OK" Then
                    DataGridView1.Rows(i).Visible = True
                Else
                    DataGridView1.Rows(i).Visible = False
                End If
            Next

        End If



    End Sub

    Private Sub DataGridView1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit

        Select Case ToolStripComboBox1.SelectedIndex
            Case Is = 0 'nada


            Case Is = 1
                'catalogos
                If estoyAgregandoRows = True Then Exit Sub

                If ToolStripComboBox1.SelectedIndex <> 1 Then Exit Sub 'si no son catalogos no hago nada

                'https://stackoverflow.com/questions/20439217/add-a-row-after-setting-datasource-to-datagridview



            Case Is = 2


            Case Is = 3


            Case Is = 4


            Case Is = 5



        End Select


    End Sub

    Private Async Sub ToolStripButton18_Click(sender As Object, e As EventArgs) Handles ToolStripButton18.Click
        'boton para eliminar el campo X
        Select Case ToolStripComboBox1.SelectedIndex
            Case Is = 0

            Case Is = 1


            Case Is = 2


            Case Is = 3
                'Records!!
                If compaSelekted = "" Then
                    MsgBox("Please select a company first!!", vbExclamation, TitBox)
                    Exit Sub
                End If

                If objetoSelek = "" Or tableSelek = "" Then
                    MsgBox("Please select a table first!!", vbExclamation, TitBox)
                    Exit Sub
                End If

                If puedoEditar = False Then
                    MsgBox("Sorry, you are not allowed to edit this template at this time, please try again later or when it becomes available.", vbExclamation, TitBox)
                    Exit Sub
                End If

                'https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.datagridview.selectionmode?view=windowsdesktop-6.0


                If DataGridView1.SelectedRows.Count = 0 Then
                    MsgBox("Please select at least 1 row to delete!", vbExclamation, TitBox)
                    Exit Sub
                End If


                Dim X As Integer
                X = MsgBox("Are you sure you want to delete this records?!" & vbCrLf & "This action can not be undone!", vbExclamation + vbYesNo, TitBox)

                If X <> 6 Then Exit Sub

                Await SigoVivo()
                Module5.AgregaTiempo()
                'Dim SelectedIndexes As List(Of Integer) = DataGridView1.SelectedRows.Cast(Of DataRow)().Select(Function(view) RecDt.Rows.IndexOf(view)).ToList()
                Dim campoYav As String = "" ' DataGridView1.SelectedRows(0).DataBoundItem

                Dim unPath As String

                'OJOO, esto también funciono pero es para borrar 1 registro!!
                ToolStripLabel1.Visible = True

                For i = 0 To DataGridView1.SelectedRows.Count - 1
                    Dim unDr As DataRowView = DataGridView1.SelectedRows(i).DataBoundItem
                    campoYav = unDr.Row.Item(0)
                    ToolStripLabel1.Text = "Deleting row..." & CStr(i + 1)
                    'se elimina y se van borrando aqui también!
                    unPath = RaizFire
                    unPath = unPath & "/" & "records"
                    unPath = unPath & "/" & compaSelekted
                    unPath = unPath & "/" & objetoSelek
                    unPath = unPath & "/" & tableSelek
                    unPath = unPath & "/" & "records"
                    Await HazDeleteEnFbSimple(unPath, campoYav)

                Next

                ToolStripLabel1.Text = "Ready"
                ToolStripLabel1.Visible = False

                MuestraRecords()

                Exit Sub


                'Dim xTipo As String = ""
                'xTipo = bs.DataSource.GetType().ToString
                Dim bs As New BindingSource
                bs.DataSource = DataGridView1.DataSource

                Dim fixDt As DataRow
                fixDt = CType(bs.DataSource.Current.Row, DataRow)

                'campoYav = fixDt.Item(0).
                'Dim SelectedIndexes As List(Of Integer) = DataGridView1.SelectedRows.Cast(Of DataRow)().Select(Function(view) RecDt.Rows.IndexOf(view)).ToList()
                Dim miDv As DataView
                miDv = CType(bs.DataSource.DataSource.DefaultView, DataView)

                Dim xDr As DataRowView = miDv(bs.DataSource.Position)

                Dim alGi As Integer = 0
                campoYav = fixDt.Item(0)
                campoYav = xDr.Item(0)

                'se elimina en la base y de paso en 

                'alGi = RecDt.Rows(0).Item(0)

            Case Is = 4
                'templates!!

                Dim enCuentra As DataRow
                enCuentra = ModuPermit.Tables(0).Rows.Find(moduloSelek.ToUpper())
                If IsNothing(enCuentra) = True Then
                    MsgBox("Sorry you are not allowed to delete fields on the selected template", vbCritical, TitBox)
                    Exit Sub
                End If

                If PussyTemp < 0 Then
                    MsgBox("Please select a table first!!", vbExclamation, TitBox)
                    Exit Sub
                End If

                If objetoSelek = "" Or tableSelek = "" Then
                    MsgBox("Please select a table first!!", vbExclamation, TitBox)
                    Exit Sub
                End If

                If puedoEditar = False Then
                    MsgBox("Sorry, you are not allowed to edit this template at this time, please try again later or when it becomes available.", vbExclamation, TitBox)
                    Exit Sub
                End If

                Await SigoVivo()
                Module5.AgregaTiempo()

                If DataGridView1.CurrentCell.RowIndex < 0 Then Exit Sub

                If IsNothing(DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(0).Value) = True Then
                    MsgBox("Please define a field code first!", vbCritical, TitBox)
                    Exit Sub
                End If

                If DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(0).Value = "" Then
                    MsgBox("Please define a field code first!", vbCritical, TitBox)
                    Exit Sub
                End If

                Dim X As Integer = 0
                X = MsgBox("ATTENTION!!" & vbCrLf & "Are you sure you want to delete the field: " & CStr(DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(0).Value) & vbCrLf & "This action can not be undone!!", vbQuestion + vbYesNo, TitBox)
                If X <> 6 Then Exit Sub

                ToolStripLabel1.Visible = True
                ToolStripLabel1.Text = "Updating positions..."

                DataGridView1.AllowUserToAddRows = False
                DataGridView1.AllowUserToDeleteRows = False

                Dim resp As String = Await BorraCampo(objetoSelek, tableSelek, CStr(DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(0).Value))
                If resp = "Ok" Then
                    'hay que re-asignar las posiciones y letras!!
                    Dim posiBack As Integer = DataGridView1.CurrentCell.RowIndex
                    Dim xDt As New DataTable
                    xDt.Columns.Add("FieldCode", GetType(String))
                    xDt.Columns.Add("Position", GetType(String))
                    xDt.Columns.Add("Letter", GetType(String))

                    DataGridView1.Rows.RemoveAt(DataGridView1.CurrentCell.RowIndex)
                    For i = 0 To DataGridView1.Rows.Count - 1
                        'se crea una matriz para hacer el put
                        DataGridView1.Rows(i).HeaderCell.Value = CStr(i + 1)
                    Next

                    For i = 0 To DataGridView1.Rows.Count - 1
                        xDt.Rows.Add({CStr(DataGridView1.Rows(i).Cells(0).Value), CStr(i + 1), LetraNumero.Tables(0).Rows(i).Item(1)})
                    Next

                    Dim elCamino As String
                    elCamino = RaizFire
                    elCamino = elCamino & "/" & "templates"
                    elCamino = elCamino & "/" & objetoSelek
                    elCamino = elCamino & "/" & tableSelek

                    resp = Await HazPutEnFireBasePathYColumnas(elCamino, xDt, 0)

                    usaDataset.Tables(0).Rows.Clear()
                    usaDataset.Tables(0).Rows.Add({"templates"})
                    tempDs.Tables.Clear()
                    tempDs = Await PullUrlWs(usaDataset, "templates")

                    MsgBox("Update done!", vbInformation, TitBox)

                End If

                ToolStripLabel1.Text = "Ready"
                ToolStripLabel1.Visible = False

                DataGridView1.AllowUserToAddRows = True

            Case Is = 5



        End Select


    End Sub

    Private Async Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        If RoleUsuario = "Editor" Then Module5.TerminaTime()

        Await DejoDeUsar()

    End Sub

    Private Async Function DejoDeUsar() As Task(Of String)

        Dim elRet As String = ""
        If inUseDt.ExtendedProperties.Item("inEdit") = True Then
            If inUseDt.ExtendedProperties.Item("lastPath") <> "" Then

                elRet = Await HazDeleteEnFbSimple(inUseDt.ExtendedProperties.Item("lastPath"), "")
                'OJO, tmb cuando cambie de nodo o se salga del catalogo!
                inUseDt.ExtendedProperties.Item("inEdit") = False
            End If
        End If

        Return elRet

    End Function

    Public Async Function siPuedoEditar(ByVal unCamino As String) As Task(Of Boolean)

        Dim siPuedo As Boolean = False

        If editDs.Tables.Count = 0 Then
            'puedoEditar = False 'ocurrió un error!
        Else
            If editDs.Tables(0).Rows.Count = 0 Then
                'puedoEditar = False
            Else
                If editDs.Tables(0).Rows(0).Item(4) = "X" Then
                    'primero verificamos el último momento!
                    'si soy yo mismo entonces solo posteo el nuevo momento actual, caso contrario,
                    'valido vs el tiempo
                    If editDs.Tables(0).Rows(0).Item(2) = UsuarioCorreo Then
                        'solo posteo el nuevo time!
                        'igual lo borro!
                        Await HazDeleteEnFbSimple(unCamino, "")
                        'Await HazPutEnFbSimple(unCamino & "/" & editDs.Tables(0).ExtendedProperties.Item("Key"), "LastUsed", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                        siPuedo = True
                    Else
                        Dim diff As Long
                        'fechAdd = DateAdd(DateInterval.Minute, 6, editDs.Tables(0).Rows(0).Item(0))
                        diff = DateDiff(DateInterval.Minute, CDate(editDs.Tables(0).Rows(0).Item(0)), DateTime.Now)
                        If diff > 11 Then
                            'lleva mas de 10 min fuera!, se patea ó se
                            'pero primero lo pateo!
                            'aunque sea YO mismo, me autoborro

                            Await HazDeleteEnFbSimple(unCamino, "")

                            siPuedo = True

                        End If

                    End If

                Else
                    'NO existia este seria el caso mas comun
                    siPuedo = True
                End If
            End If

        End If

        Return siPuedo

    End Function

    Private Sub NoPuedesEditar(ByVal miDs As DataSet, ByRef elLabel As ToolStripLabel)

        elLabel.ForeColor = Color.Crimson
        If miDs.Tables.Count = 0 Then
            elLabel.Text = "You can't edit this object"
        Else
            If miDs.Tables(0).Rows.Count = 0 Then
                elLabel.Text = "You can't edit this object"
            Else
                elLabel.Text = miDs.Tables(0).Rows(0).Item(3) & " is editing this object"
            End If
        End If

    End Sub

    Private Async Sub HandleElapsed(ByVal remainingTime As TimeSpan)
        'ToolStripLabel21.Text = "Tiempo restante: " & String.Format("{0}:{1:d2}:{2:d2}", remainingTime.Hours, remainingTime.Minutes, remainingTime.Seconds)
        Label3.Text = String.Format("{0}:{1:d2}:{2:d2}", remainingTime.Hours, remainingTime.Minutes, remainingTime.Seconds)
        If remainingTime.Minutes > 1 Then
            'ToolStripLabel21.ForeColor = Color.DarkGreen
            Label3.ForeColor = Color.DarkGreen
        Else
            'ToolStripLabel21.ForeColor = Color.Red
            If toyTrabajando = True Then
                Await SigoVivo()
                Module5.AgregaTiempo()
            End If
            Label3.ForeColor = Color.Red
        End If
    End Sub

    Private Sub HandleComplete(ByVal message As String)

        'If Application.OpenForms().OfType(Of Form16).Any Then

        'Form16.Close()

        'End If

        Me.Close()

    End Sub

    Private Async Function SigoVivo() As Task(Of String)
        Dim elReg As String = "fail"
        If inUseDt.ExtendedProperties.Item("Key") = "" Then Return elReg
        If inUseDt.ExtendedProperties.Item("lastPath") = "" Then Return elReg

        elReg = Await HazPutEnFbSimple(inUseDt.ExtendedProperties.Item("lastPath") & "/" & inUseDt.ExtendedProperties.Item("Key"), "LastUsed", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        'If editDs.Tables.Count = 0 Then Return elReg
        'If IsNothing(editDs.Tables(0).ExtendedProperties.Item("Key")) = True Then Return elReg
        'If IsNothing(editDs.Tables(0).ExtendedProperties.Item("lastPath")) = True Then Return elReg

        'If editDs.Tables(0).ExtendedProperties.Item("Key") = "" Then Return elReg

        'If editDs.Tables(0).ExtendedProperties.Item("lastPath") = "" Then Return elReg

        'elReg = Await HazPutEnFbSimple(editDs.Tables(0).ExtendedProperties.Item("lastPath") & "/" & editDs.Tables(0).ExtendedProperties.Item("Key"), "LastUsed", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

        Return elReg

    End Function

    Private Async Sub ToolStripButton21_Click(sender As Object, e As EventArgs) Handles ToolStripButton21.Click

        If ToolStripComboBox1.SelectedIndex <> 4 Then Exit Sub

        If PussyTemp < 0 Then
            MsgBox("Please select a table first!!", vbExclamation, TitBox)
            Exit Sub
        End If

        If objetoSelek = "" Or tableSelek = "" Then
            MsgBox("Please select a table first!!", vbExclamation, TitBox)
            Exit Sub
        End If

        If puedoEditar = False Then
            MsgBox("Sorry you are not allowed to edit this template at this time!, please try again later!", vbExclamation, TitBox)
            Exit Sub
        End If

        Await SigoVivo()
        Module5.AgregaTiempo()

        Dim unCamino As String
        unCamino = RaizFire
        unCamino = unCamino & "/" & "templates"
        unCamino = unCamino & "/" & objetoSelek
        unCamino = unCamino & "/" & tableSelek

        Form2.keyValue = ""
        Form2.tabValue = ""
        Form2.elCamino = "templates" & "/" & objetoSelek & "/" & tableSelek
        Form2.queOpcion = 7
        Form2.huboExito = False
        Form2.pathLabel = unCamino

        Form2.ShowDialog()

        If Form2.huboExito = False Then Exit Sub
        'guardamos el field code primero, y luego lo agregamos al dgv
        'reload del template unitario?!

        elNode.Nodes.Add(Form2.keyValue, Form2.keyValue & " / " & Form2.tabValue, 8, 8)

        unCamino = RaizFire
        unCamino = unCamino & "/" & "templates"
        unCamino = unCamino & "/" & objetoSelek

        Dim dt As New DataTable
        dt = Await PullDtFireBase(unCamino, "tempunit", objetoSelek)
        tempDs.Tables.RemoveAt(PussyTemp)
        tempDs.Tables.Add(dt)
        PussyTemp = tempDs.Tables.Count - 1

        DataGridView1.Rows.Add()
        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(0).Value = Form2.keyValue
        DataGridView1.Rows(DataGridView1.Rows.Count - 1).Cells(1).Value = Form2.tabValue
        DataGridView1.Rows(DataGridView1.Rows.Count - 1).HeaderCell.Value = CStr(DataGridView1.Rows.Count)

    End Sub

    Private Async Sub ToolStripButton19_Click(sender As Object, e As EventArgs) Handles ToolStripButton19.Click
        'move field down
        If IsNothing(DataGridView1.CurrentCell) = True Then
            MsgBox("Please select a row with data to move up or down!", vbCritical, TitBox)
            Exit Sub
        End If

        Dim posXold As Integer = DataGridView1.CurrentCell.RowIndex
        If posXold + 1 > DataGridView1.Rows.Count - 1 Then
            MsgBox("This field is already on bottom!!", vbExclamation, TitBox)
            Exit Sub
        End If

        Dim xRow As New DataGridViewRow
        xRow = DataGridView1.Rows(posXold).Clone()

        For i = 0 To xRow.Cells.Count - 1
            xRow.Cells(i).Value = DataGridView1.Rows(posXold).Cells(i).Value
        Next

        estoyAgregandoRows = True

        DataGridView1.Rows.InsertRange(posXold + 2, xRow)
        DataGridView1.Rows.RemoveAt(posXold)

        For i = 0 To DataGridView1.Rows.Count - 1
            DataGridView1.Rows(i).HeaderCell.Value = CStr(i + 1)
        Next

        Dim xDt As New DataTable
        xDt.Columns.Add("FieldCode", GetType(String))
        xDt.Columns.Add("Position", GetType(String))
        xDt.Columns.Add("Letter", GetType(String))

        xDt.Rows.Add({CStr(DataGridView1.Rows(posXold).Cells(0).Value), CStr(posXold + 1), LetraNumero.Tables(0).Rows(posXold).Item(1)})
        xDt.Rows.Add({CStr(DataGridView1.Rows(posXold + 1).Cells(0).Value), CStr(posXold + 1 + 1), LetraNumero.Tables(0).Rows(posXold + 1).Item(1)})

        Dim elCamino As String
        elCamino = RaizFire
        elCamino = elCamino & "/" & "templates"
        elCamino = elCamino & "/" & objetoSelek
        elCamino = elCamino & "/" & tableSelek

        Await HazPutEnFireBasePathYColumnas(elCamino, xDt, 0)

        DataGridView1.CurrentCell = DataGridView1.Rows(posXold).Cells(0)

        rowEstaba = posXold

        estoyAgregandoRows = False

    End Sub

    Private Async Sub ToolStripButton20_Click(sender As Object, e As EventArgs) Handles ToolStripButton20.Click
        'move field up
        If IsNothing(DataGridView1.CurrentCell) = True Then
            MsgBox("Please select a row with data to move up or down!", vbCritical, TitBox)
            Exit Sub
        End If

        Dim posXold As Integer = DataGridView1.CurrentCell.RowIndex

        If posXold - 1 < 0 Then
            MsgBox("This field is already on top!!", vbExclamation, TitBox)
            Exit Sub
        End If

        Dim xRow As New DataGridViewRow
        xRow = DataGridView1.Rows(posXold).Clone()

        For i = 0 To xRow.Cells.Count - 1
            xRow.Cells(i).Value = DataGridView1.Rows(posXold).Cells(i).Value
        Next

        estoyAgregandoRows = True

        DataGridView1.Rows.InsertRange(posXold - 1, xRow)

        DataGridView1.Rows.RemoveAt(posXold + 1)

        For i = 0 To DataGridView1.Rows.Count - 1
            DataGridView1.Rows(i).HeaderCell.Value = CStr(i + 1)
        Next

        Dim xDt As New DataTable
        xDt.Columns.Add("FieldCode", GetType(String))
        xDt.Columns.Add("Position", GetType(String))
        xDt.Columns.Add("Letter", GetType(String))

        xDt.Rows.Add({CStr(DataGridView1.Rows(posXold - 1).Cells(0).Value), CStr(posXold - 1 + 1), LetraNumero.Tables(0).Rows(posXold - 1).Item(1)})
        xDt.Rows.Add({CStr(DataGridView1.Rows(posXold).Cells(0).Value), CStr(posXold + 1), LetraNumero.Tables(0).Rows(posXold).Item(1)})

        Dim elCamino As String
        elCamino = RaizFire
        elCamino = elCamino & "/" & "templates"
        elCamino = elCamino & "/" & objetoSelek
        elCamino = elCamino & "/" & tableSelek

        Await HazPutEnFireBasePathYColumnas(elCamino, xDt, 0)

        DataGridView1.CurrentCell = DataGridView1.Rows(posXold - 1).Cells(0)

        rowEstaba = posXold - 1

        estoyAgregandoRows = False


    End Sub

    Private Sub ToolStripButton22_Click(sender As Object, e As EventArgs) Handles ToolStripButton22.Click

        If ToolStripComboBox1.SelectedIndex <> 4 Then Exit Sub

        If PussyTemp < 0 Then
            MsgBox("Please select a table first!!", vbExclamation, TitBox)
            Exit Sub
        End If

        If objetoSelek = "" Or tableSelek = "" Then
            MsgBox("Please select a table first!!", vbExclamation, TitBox)
            Exit Sub
        End If

        If puedoEditar = False Then
            MsgBox("Sorry you are not allowed to edit this template at this time!, please try again later!", vbExclamation, TitBox)
            Exit Sub
        End If

        If DataGridView1.CurrentCell.RowIndex < 0 Then Exit Sub

        If IsNothing(DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(0).Value) = True Then
            MsgBox("Please define a field code first!", vbCritical, TitBox)
            Exit Sub
        End If

        If DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(0).Value = "" Then
            MsgBox("Please define a field code first!", vbCritical, TitBox)
            Exit Sub
        End If

        'Editar el campo, Solo puede 
        Dim elPa As String = ""
        elPa = RaizFire
        elPa = elPa & "/" & "templates"
        elPa = elPa & "/" & objetoSelek
        elPa = elPa & "/" & tableSelek

        Form2.pathLabel = elPa
        Form2.elCamino = "templates/" & objetoSelek & "/" & tableSelek
        Form2.keyValue = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(0).Value
        Form2.tabValue = DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(1).Value
        Form2.queOpcion = 8
        Form2.huboExito = False

        Form2.ShowDialog()

        If Form2.huboExito = False Then Exit Sub

        DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(0).Value = Form2.keyValue
        DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex).Cells(1).Value = Form2.tabValue


    End Sub

    Private Sub ToolStripButton23_Click(sender As Object, e As EventArgs) Handles ToolStripButton23.Click

        Select Case ToolStripComboBox1.SelectedIndex
            Case Is = 0


            Case Is = 1


            Case Is = 2


            Case Is = 3
                'records
                If puedoEditar = False Then
                    MsgBox("Sorry you are not allowed to modify this template!!", vbCritical, TitBox)
                    Exit Sub
                End If

                DataGridView1.DataSource = Nothing
                DataGridView1.Rows.Clear()

            Case Is = 4
                'templates


        End Select

    End Sub

    Private Sub ToolStripButton24_Click(sender As Object, e As EventArgs) Handles ToolStripButton24.Click

        If ToolStripComboBox1.SelectedIndex <> 3 Then Exit Sub

        If DoneEvaluacion = False Then
            MsgBox("Please evaluate first the records of this object!!", vbExclamation, TitBox)
            Exit Sub
        End If

        If DataGridView1.Rows.Count - 1 < 0 Then
            MsgBox("Deploy some information first!!", vbCritical, TitBox)
            Exit Sub
        End If

        Dim lCol As Integer = 1

        'If ToolStripComboBox1.SelectedIndex = 3 Then lCol = 4

        Dim diLOg As New SaveFileDialog
        Dim oFileName As String
        oFileName = ""
        diLOg.InitialDirectory = Environment.SpecialFolder.MyDocuments
        diLOg.Filter = "Excel CSV File (*.csv)|*.csv" '"Excel File (*.xls)|*.xls"
        'diLOg.Filter = "Excel File (*.xls)|*.xls|Todos los archivos (*.*)|*.*"
        If diLOg.ShowDialog = System.Windows.Forms.DialogResult.OK Then

            Call ExportErrors(DataGridView1, diLOg.FileName, 1)

        End If

    End Sub
End Class
