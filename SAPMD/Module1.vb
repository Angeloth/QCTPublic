Imports Firebase.Database
Imports Firebase.Database.Query
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography

Module Module1
    'Modulo para realizar las escrituras y lecturas de FireBase!
    Public Const TitBox As String = "QDT"
    Public Const RaizFire As String = "https://radiant-rig-337421-default-rtdb.firebaseio.com" 'sustituir por el de la empresa
    Public RoleUsuario As String
    Public ModuPermit As New DataSet
    Public LetraNumero As New DataSet
    Public matAlpha(27) As String

    Public UsuarioNombre As String
    Public UsuarioCorreo As String
    Public UsuarioRole As String
    Public UsuarioApellido As String
    Public UsuarioModulos As String

    Public Async Function HazPutEnFbSimple(ByVal elPath As String, ByVal laYave As String, ByVal elDato As String) As Task(Of String)

        Dim elRegreso As String = ""
        Dim client = New FirebaseClient(elPath)

        'laYave = elSet.Rows(i).Item(posYave) '0,1
        Dim miDato As String = """" & elDato & """"

        Try

            Await client.Child(laYave).PutAsync(miDato).ConfigureAwait(False)
            elRegreso = "Ok"
        Catch ex As Exception
            elRegreso = "Error adding this data > " & laYave & " : " & elDato
        End Try

        Return elRegreso

    End Function

    Public Async Function HazDeleteEnFbSimple(ByVal elPath As String, ByVal Hijo As String) As Task(Of String)

        Dim elRet As String = ""
        Dim client = New FirebaseClient(elPath)

        Try
            Await client.Child(Hijo).DeleteAsync().ConfigureAwait(False)
            elRet = "Ok"
        Catch ex As Exception
            elRet = ex.Message
        End Try

        Return elRet

    End Function

    Public Async Function HazPostEnFireBaseConPathYColumnas(ByVal elPath As String, ByVal elSet As DataTable, ByVal elChild As String, ByVal colYave As Integer) As Task(Of String)

        If elSet.Rows.Count = 0 Then Return "No records to write!"

        Dim client = New FirebaseClient(elPath)
        Dim laYave As String = ""
        Dim miDato As String = ""
        Dim elHijo As String = ""

        Dim elRegreso As String = ""
        Dim cuentaErrores As Long = 0
        Dim cuentaOk As Long = 0
        Dim errLog As String = ""

        Dim cadJson As String = ""

        Dim elDato As String = ""
        elDato = "{" & vbCrLf
        elDato = elDato & """MX"":""Mexico"""
        elDato = elDato & "," & vbCrLf
        elDato = elDato & """GT"":""Guatemala"""
        elDato = elDato & vbCrLf & "}" 'este funciona pero agrega hijos!!

        Dim k As Integer = 0

        For i = 0 To elSet.Rows.Count - 1

            'elHijo = elSet.Rows(i).Item(colYave)

            cadJson = "{" & vbCrLf
            k = 0
            For j = 0 To elSet.Columns.Count - 1

                If j = colYave Then Continue For

                If k > 0 Then cadJson = cadJson & "," & vbCrLf

                miDato = """" & elSet.Rows(i).Item(j) & """"

                cadJson = cadJson & """" & elSet.Columns(j).ColumnName & """:" & miDato

                k = k + 1

            Next

            cadJson = cadJson & vbCrLf & "}"

            Try

                Await client.Child(elChild).PostAsync(cadJson)
                cuentaOk = cuentaOk + 1
            Catch ex As Exception

                cuentaErrores = cuentaErrores + 1
                errLog = errLog & "Error writting on node: " & elPath & " > " & laYave & " > " & miDato & " : " & ex.Message & vbCrLf

            End Try


        Next


        If cuentaOk > 0 Then
            elRegreso = elRegreso & cuentaOk & " records added!" & vbCrLf & vbCrLf
        End If

        If cuentaErrores > 0 Then
            elRegreso = elRegreso & cuentaErrores & " Errors on writting..." & vbCrLf
            elRegreso = elRegreso & errLog
        End If

        Return elRegreso

    End Function


    Public Async Function HazPutEnFireBasePathYColumnas(ByVal elPath As String, ByVal elSet As DataTable, ByVal colYave As Integer) As Task(Of String)

        If elSet.Rows.Count = 0 Then Return "No records to write!"

        Dim client = New FirebaseClient(elPath)
        Dim laYave As String = ""
        Dim miDato As String = ""
        Dim elHijo As String = ""

        Dim elRegreso As String = ""
        Dim cuentaErrores As Long = 0
        Dim cuentaOk As Long = 0
        Dim errLog As String = ""
        For i = 0 To elSet.Rows.Count - 1

            elHijo = elSet.Rows(i).Item(colYave)

            For j = 0 To elSet.Columns.Count - 1

                If j = colYave Then Continue For

                laYave = elSet.Columns(j).ColumnName '0,1
                miDato = """" & elSet.Rows(i).Item(j) & """"

                Try

                    Await client.Child(elHijo).Child(laYave).PutAsync(miDato).ConfigureAwait(False)
                    cuentaOk = cuentaOk + 1
                Catch ex As Exception

                    cuentaErrores = cuentaErrores + 1
                    errLog = errLog & "Error writting on node: " & elPath & " > " & laYave & " > " & miDato & " : " & ex.Message & vbCrLf

                End Try

            Next

        Next


        If cuentaOk > 0 Then
            elRegreso = elRegreso & cuentaOk & " records added!" & vbCrLf & vbCrLf
        End If

        If cuentaErrores > 0 Then
            elRegreso = elRegreso & cuentaErrores & " Errors on writting..." & vbCrLf
            elRegreso = elRegreso & errLog
        End If

        Return elRegreso

    End Function

    Public Async Function HazUpdateEnFireBaseConYave(ByVal elPath As String, ByVal posYave As Integer, ByVal elSet As DataTable) As Task(Of String)
        If elSet.Rows.Count = 0 Then Return "No records to write!"
        Dim client = New FirebaseClient(elPath)
        Dim laYave As String = ""
        Dim laKey As String = ""
        Dim miDato As String = ""
        Dim elRegreso As String = ""
        Dim cuentaErrores As Long = 0
        Dim cuentaOk As Long = 0
        Dim errLog As String = ""

        For i = 0 To elSet.Rows.Count - 1
            laYave = elSet.Rows(i).Item(posYave) '0,1

            ' .Rows(i).Item(posKey)

            For j = 0 To elSet.Columns.Count - 1

                If j = posYave Then Continue For

                laKey = elSet.Columns(j).ColumnName
                miDato = """" & elSet.Rows(i).Item(j) & """"

                Try

                    Await client.Child(laYave).Child(laKey).PutAsync(miDato).ConfigureAwait(False)
                    cuentaOk = cuentaOk + 1
                Catch ex As Exception
                    cuentaErrores = cuentaErrores + 1
                    errLog = errLog & "Error writting on node: " & elPath & " > " & laYave & " > " & miDato & " : " & ex.Message & vbCrLf
                    'Err = ex.Message
                End Try

            Next

        Next

        If cuentaOk > 0 Then
            elRegreso = elRegreso & cuentaOk & " records added!" & vbCrLf & vbCrLf
        End If

        If cuentaErrores > 0 Then
            elRegreso = elRegreso & cuentaErrores & " Errors on writting..." & vbCrLf
            elRegreso = elRegreso & errLog
        End If

        Return elRegreso

    End Function

    Public Async Function HazPutEnFireBaseConPath(ByVal elPath As String, ByVal elSet As DataTable, ByVal posYave As Integer, ByVal posDato As Integer) As Task(Of String)

        'La primera tabla traiga el numero de niveles a escribir!!
        'solo es una columna, el numero de renglones es el numero de niveles, y el valor del renglon es el nodo!

        'la SEGUNDA tabla que tenga los valores finales, ciclado a pair/value
        'cuando se escriba en un nodo, se va recibir una tabla
        'El titulo de la columna sera el nombre del campo(Key) y el valor del renglon sera el valor (Value)
        'Column1    Column2     Column3
        'Value1     Value2      Value3
        'If elSet.Tables.Count < 2 Then Return "No Tables found!"
        'If elSet.Rows.Count = 0 Then Return "No-Tables"
        If elSet.Rows.Count = 0 Then Return "No records to write!"

        Dim client = New FirebaseClient(elPath)

        Dim laYave As String = ""
        Dim miDato As String = ""
        Dim elRegreso As String = ""
        Dim cuentaErrores As Long = 0
        Dim cuentaOk As Long = 0
        Dim errLog As String = ""

        For i = 0 To elSet.Rows.Count - 1
            laYave = elSet.Rows(i).Item(posYave) '0,1
            miDato = """" & elSet.Rows(i).Item(posDato) & """"
            Try

                Await client.Child(laYave).PutAsync(miDato).ConfigureAwait(False)
                cuentaOk = cuentaOk + 1
            Catch ex As Exception
                cuentaErrores = cuentaErrores + 1
                errLog = errLog & "Error writting on node: " & elPath & " > " & laYave & " > " & miDato & " : " & ex.Message & vbCrLf
                'Err = ex.Message
            End Try
        Next

        If cuentaOk > 0 Then
            elRegreso = elRegreso & cuentaOk & " records added!" & vbCrLf & vbCrLf
        End If

        If cuentaErrores > 0 Then
            elRegreso = elRegreso & cuentaErrores & " Errors on writting..." & vbCrLf
            elRegreso = elRegreso & errLog
        End If

        Return elRegreso

    End Function

    Public Async Function HazDeleteAFireBase(ByVal unaRaiz As String, ByVal TabDel As DataTable) As Task(Of String)

        Dim client = New FirebaseClient(unaRaiz)
        'nodo a eliminarts!
        Dim cuentaOks As Long = 0
        Dim cuentaErr As Long = 0
        Dim errLog As String = ""
        Dim elRegreso As String = ""
        For i = 0 To TabDel.Rows.Count - 1

            Try
                Await client.Child(CStr(TabDel.Rows(i).Item(0))).DeleteAsync().ConfigureAwait(False)
                cuentaOks = cuentaOks + 1
            Catch ex As Exception
                cuentaErr = cuentaErr + 1
                errLog = errLog & "Error deleting on node: " & unaRaiz & " > " & CStr(TabDel.Rows(i).Item(0)) & " : " & ex.Message & vbCrLf
            End Try

        Next

        If cuentaOks > 0 Then
            elRegreso = elRegreso & cuentaOks & " record(s) deleted!" & vbCrLf
        End If

        If cuentaErr > 0 Then
            elRegreso = elRegreso & cuentaErr & " error(s) found on deleting items:" & vbCrLf & vbCrLf
            elRegreso = elRegreso & errLog
        End If

        Return elRegreso

    End Function

    Public Async Function HazCambiosAFireBase(ByVal elSet As DataSet, ByVal laOpcion As Integer) As Task(Of String)

        Dim xPath As String = RaizFire
        Dim i As Integer = 0
        Dim k As Integer = 0
        Dim Posi As Integer = 0
        Dim cadResult As String = "Procedure result: " & vbCrLf
        Dim ySet As New DataSet
        Dim wSet As New DataSet
        Dim xSet As New DataSet
        Dim zSet As New DataSet

        Select Case laOpcion
            Case Is = 1
                'catalogos!
                For i = 0 To elSet.Tables(0).Rows.Count - 1
                    xPath = xPath & "/" & CStr(elSet.Tables(0).Rows(i).Item(0))
                Next
                '1 nuevos
                '2 antiguos
                If elSet.Tables(1).Rows.Count > 0 Then
                    cadResult = cadResult & Await HazPostEnFireBaseConPathYColumnas(xPath, elSet.Tables(1), "", 2) 'Await HazPutEnFireBaseConPath(xPath, elSet.Tables(1), 0, 1)
                    cadResult = cadResult & vbCrLf
                End If

                If elSet.Tables(2).Rows.Count > 0 Then
                    For i = 0 To elSet.Tables(2).Rows.Count - 1
                        cadResult = cadResult & Await HazDeleteEnFbSimple(xPath, elSet.Tables(2).Rows(i).Item(2))
                    Next
                End If

                If elSet.Tables(4).Rows.Count > 0 Then
                    cadResult = cadResult & Await HazUpdateEnFireBaseConYave(xPath, 2, elSet.Tables(4))
                End If


            Case Is = 2
                'dependencias

            Case Is = 3
                'records

            Case Is = 4
                'templates!

                'se debe transformar la info, 
                'sacar datos únicos, y lueeego, tomar la info!
                'osea todas las actualizaciones por nodo de la tabla, md32-0001, y de esa a las demas!
                'igual y construir una datatable con esta info!

                'Campos!
                For i = 0 To elSet.Tables(1).Rows.Count - 1
                    'nuevos!

                    Posi = ySet.Tables.IndexOf(CStr(elSet.Tables(1).Rows(i).Item(0)))
                    If Posi < 0 Then
                        ySet.Tables.Add(CStr(elSet.Tables(1).Rows(i).Item(0)))
                        Posi = ySet.Tables.Count - 1
                        ySet.Tables(Posi).Columns.Add("Key")
                        ySet.Tables(Posi).Columns.Add("Value")
                        'ySet.Tables(Posi).Rows.Add({"TableName", CStr(elSet.Tables(1).Rows(i).Item(1))})
                        'los 4 renglones con las 4 propiedades
                        ySet.Tables(Posi).Rows.Add({"Letter", CStr(elSet.Tables(1).Rows(i).Item(4))})
                        ySet.Tables(Posi).Rows.Add({"Name", CStr(elSet.Tables(1).Rows(i).Item(1))})
                        ySet.Tables(Posi).Rows.Add({"Position", CStr(elSet.Tables(1).Rows(i).Item(3))})
                        ySet.Tables(Posi).Rows.Add({"isKey", CStr(elSet.Tables(1).Rows(i).Item(2))})
                    End If

                Next

                For i = 0 To ySet.Tables.Count - 1
                    xPath = RaizFire
                    For j = 0 To elSet.Tables(0).Rows.Count - 1
                        xPath = xPath & "/" & CStr(elSet.Tables(0).Rows(j).Item(0))
                    Next
                    xPath = xPath & "/" & ySet.Tables(i).TableName 'field code!
                    cadResult = cadResult & Await HazPutEnFireBaseConPath(xPath, ySet.Tables(i), 0, 1) & vbCrLf
                Next




                'ojo ahora los Nodos que se eliminan!!
                For i = 0 To elSet.Tables(2).Rows.Count - 1
                    Posi = xSet.Tables.IndexOf(CStr(elSet.Tables(2).Rows(i).Item(0)))
                    If Posi < 0 Then
                        xSet.Tables.Add(CStr(elSet.Tables(2).Rows(i).Item(0)))
                        Posi = xSet.Tables.Count - 1
                        xSet.Tables(Posi).Columns.Add("Key")
                        xSet.Tables(Posi).Rows.Add({CStr(elSet.Tables(2).Rows(i).Item(0))})
                    End If

                Next

                For i = 0 To xSet.Tables.Count - 1
                    xPath = RaizFire
                    For j = 0 To elSet.Tables(0).Rows.Count - 1
                        xPath = xPath & "/" & CStr(elSet.Tables(0).Rows(j).Item(0))
                    Next
                    cadResult = cadResult & Await HazDeleteAFireBase(xPath, xSet.Tables(i))
                Next


                'Y ahora las llaves que se eliminan!




        End Select

        Return cadResult

    End Function


    Public Async Function PullUrlWs(ByVal elSet As DataSet, ByVal elConcepto As String) As Task(Of DataSet)
        'la tabla a enviar solo va a ser 1
        '1 tabla con los renglones y los nodos!, se cicla hasta construir el url con el nodo en orden, y se llama
        'a la funcion jalaws
        Dim xSet As New DataSet 'en este set se va a escribir la tabla resultante!
        'If elSet.Tables.Count = 0 Then Return xSet
        'If elSet.Tables(0).Rows.Count = 0 Then Return xSet
        Dim i As Integer = 0
        Dim k As Integer = 0
        Dim noKids As Boolean = False
        Dim laCade As String = RaizFire
        Dim tabX As String = ""
        Dim moduX As String = ""
        Dim miTab As String = ""
        Dim LevKey As String = ""
        Dim rulKey As String = ""
        Dim p1 As Integer = 0
        Dim p2 As Integer = 0

        Dim userCorreo As String = ""
        Dim userName As String = ""
        Dim userApellido As String = ""
        Dim userRole As String = ""
        Dim userModules As String = ""
        Dim userPass As String = ""
        Dim userFirst As String = ""


        If elSet.Tables.Count > 0 Then
            For i = 0 To elSet.Tables(0).Rows.Count - 1
                laCade = laCade & "/" & CStr(elSet.Tables(0).Rows(i).Item(0))
            Next
        End If

        Dim client = New FirebaseClient(laCade)

        'OJO solo habra 4 nodos principales
        'catalogs:
        'global>gbaaaa>AD:Andorra
        '2 fors

        'dependencies:
        'md01>tablax>field>FieldCode:Kunnr
        '3 fors

        'records
        'us10>fi>md01>key1>KTopl:ITUS
        '4 fors

        'templates:
        'md01>tabla1>ktopl:Chart of accounts

        Try
            Dim dinos = Await client.Child("").OnceAsync(Of Object)
            'Dim fbo As FirebaseObject(Of Object)
            Select Case elConcepto
                Case Is = "catalogs"
                    For Each dino In dinos
                        'xSet.Tables.Add("")
                        Dim ser As JObject = JObject.Parse(dino.Object.ToString)
                        Dim datos As List(Of JToken) = ser.Children().ToList
                        Dim doDatos As List(Of JToken)
                        Dim deDatos As List(Of JToken)
                        For Each item As JProperty In datos
                            item.CreateReader()
                            xSet.Tables.Add(dino.Key.ToString() & "#" & item.Name)
                            'gb#gb0007

                            Dim iAveprim(1) As DataColumn
                            Dim kEys As New DataColumn()

                            kEys.ColumnName = "CatalogName"
                            iAveprim(0) = kEys

                            xSet.Tables(xSet.Tables.Count - 1).Columns.Add(kEys) '"CatalogName", GetType(String)
                            xSet.Tables(xSet.Tables.Count - 1).Columns.Add("Value", GetType(String))
                            xSet.Tables(xSet.Tables.Count - 1).Columns.Add("Unike", GetType(String)) 'Yave de firebase
                            xSet.Tables(xSet.Tables.Count - 1).CaseSensitive = True
                            xSet.Tables(xSet.Tables.Count - 1).PrimaryKey = iAveprim

                            doDatos = item.Value.Children().ToList()

                            For Each catitem As JProperty In doDatos

                                If catitem.Name = "CatalogName" Then

                                    xSet.Tables(xSet.Tables.Count - 1).Columns(0).ColumnName = catitem.Value
                                    Continue For
                                End If

                                deDatos = catitem.Value.Children.ToList()

                                tabX = ""
                                moduX = ""

                                For Each prap As JProperty In deDatos

                                    Select Case prap.Name
                                        Case Is = "Key"
                                            tabX = prap.Value

                                        Case Is = "Description"
                                            moduX = prap.Value

                                    End Select

                                Next

                                xSet.Tables(xSet.Tables.Count - 1).Rows.Add({tabX, moduX, catitem.Name})

                            Next


                        Next


                    Next


                Case Is = "catpro"
                    'nuevos catalogos

                    For Each dino In dinos

                        Dim ser As JObject = JObject.Parse(dino.Object.ToString)

                        Dim datos As List(Of JToken) = ser.Children().ToList

                        Dim doDatos As List(Of JToken)

                        Dim deDatos As List(Of JToken)

                        Dim diDatos As List(Of JToken)

                        For Each item As JProperty In datos

                            If item.Name = "ModuleName" Then

                                moduX = item.Value.ToString()

                                Continue For

                            End If

                            xSet.Tables.Add(dino.Key.ToString() & "#" & item.Name) '

                            xSet.Tables(xSet.Tables.Count - 1).ExtendedProperties.Add("ModuleName", moduX)

                            xSet.Tables(xSet.Tables.Count - 1).ExtendedProperties.Add("CatalogCode", item.Name)

                            xSet.Tables(xSet.Tables.Count - 1).ExtendedProperties.Add("CatalogName", "Undefined")

                            xSet.Tables(xSet.Tables.Count - 1).ExtendedProperties.Add("ColumnCount", 0)

                            xSet.Tables(xSet.Tables.Count - 1).ExtendedProperties.Add("ModuleCode", dino.Key.ToString())

                            'la primer columna SIEMPRE va ser el campo Yave, osea la concatenación de las demás columnas!

                            Dim iAveprim(1) As DataColumn

                            Dim kEys As New DataColumn()

                            kEys.ColumnName = "KeyField"

                            iAveprim(0) = kEys

                            xSet.Tables(xSet.Tables.Count - 1).Columns.Add(kEys) '0 Siempre!
                            xSet.Tables(xSet.Tables.Count - 1).Columns.Add("fbKey", GetType(String)) '1 'Yave de firebase!
                            xSet.Tables(xSet.Tables.Count - 1).CaseSensitive = True
                            xSet.Tables(xSet.Tables.Count - 1).PrimaryKey = iAveprim

                            doDatos = item.Value.Children().ToList()

                            For Each catitem As JProperty In doDatos

                                If catitem.Name = "CatalogName" Then

                                    xSet.Tables(xSet.Tables.Count - 1).ExtendedProperties.Item("CatalogName") = catitem.Value.ToString()

                                    Continue For

                                End If

                                deDatos = catitem.Value.Children.ToList()

                                For Each reitem As JProperty In deDatos

                                    diDatos = reitem.Value.Children.ToList()

                                    Select Case catitem.Name

                                        Case Is = "data"

                                            'agregamos las columnas

                                            xSet.Tables(xSet.Tables.Count - 1).Columns.Add()

                                            For Each lastItem As JProperty In diDatos

                                                Select Case lastItem.Name

                                                    Case Is = "FieldCode"

                                                        xSet.Tables(xSet.Tables.Count - 1).Columns(xSet.Tables(xSet.Tables.Count - 1).Columns.Count - 1).ColumnName = lastItem.Value.ToString()



                                                    Case Is = "FieldName"

                                                        xSet.Tables(xSet.Tables.Count - 1).Columns(xSet.Tables(xSet.Tables.Count - 1).Columns.Count - 1).ExtendedProperties.Add("FieldName", lastItem.Value.ToString())



                                                    Case Is = "Position"

                                                        xSet.Tables(xSet.Tables.Count - 1).Columns(xSet.Tables(xSet.Tables.Count - 1).Columns.Count - 1).ExtendedProperties.Add("Position", CInt(lastItem.Value.ToString()))



                                                    Case Is = "isText"

                                                        xSet.Tables(xSet.Tables.Count - 1).Columns(xSet.Tables(xSet.Tables.Count - 1).Columns.Count - 1).ExtendedProperties.Add("isText", lastItem.Value.ToString())



                                                    Case Is = "Description"

                                                        xSet.Tables(xSet.Tables.Count - 1).Columns(xSet.Tables(xSet.Tables.Count - 1).Columns.Count - 1).ExtendedProperties.Add("Description", lastItem.Value.ToString())

                                                End Select

                                            Next



                                        Case Is = "records"

                                            'agregamos los registros!

                                            tabX = ""

                                            xSet.Tables(xSet.Tables.Count - 1).Rows.Add({"YaveX"})

                                            For Each lastItem As JProperty In diDatos

                                                xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(lastItem.Name) = lastItem.Value.ToString()

                                                If lastItem.Name = "ZZ" Then Continue For 'NO se guarda como llave!

                                                If tabX <> "" Then tabX = tabX & "#"

                                                tabX = tabX & lastItem.Value.ToString()

                                            Next

                                            xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(0) = tabX
                                            xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(1) = reitem.Name 'yave de firebase!
                                            'el ultimo field code siempre será ZZZZ

                                            'De que forma se pueden vincular los catalogos!?

                                            '1. Directo de entre una lista de valores, siempre será de un campo yave, no repetible!. Siempre se marca primero el campo final a comparar,

                                            'si el campo final a comparar NO es yave(hay repetidos), se debe definir al condicionante de las columnas anteriores a cumplir para que se tome el valor correcto,

                                            'Puede ir ligado a un campo de SU mismo renglon de otra columna que deba compararse y al hacer match, se deba tomar el valor fulano o sutano, ó un set de valores

                                            '2. Se puede condicionar el valor a utilizar dependiendo del mandante ó combinación de otros campos anteriores (company code)

                                            'Ejemplo: Region Code!, hay repetidos, pero depende del pais, y este puede estar dentro de un campo del mismo template ó de un campo externo?

                                            '3. Puede estar vinculado a un intervalo de valores a evaluarts!

                                            'Se requiere poner en templates:

                                            'Match de campo vs: modulecode>catalogcode>fieldcode conditions:CampoMatch1#CampoMatch1-CampoMatch2#CampoMatch2, igual NO puede haber repetidos, al final se elige de un set de valores finales

                                            'Si el set final de valores es 1 renglon, entonces se coloca automáticamente!!

                                            'Una última opción es VS mandante, ya sea company code ó plant code!

                                    End Select

                                Next

                            Next

                            'A ver si con esto jala!
                            xSet.Tables(xSet.Tables.Count - 1).ExtendedProperties.Item("ColumnCount") = xSet.Tables(xSet.Tables.Count - 1).Rows.Count
                            For i = 2 To xSet.Tables(xSet.Tables.Count - 1).Columns.Count - 1

                                xSet.Tables(xSet.Tables.Count - 1).Columns(i).SetOrdinal(CInt(xSet.Tables(xSet.Tables.Count - 1).Columns(i).ExtendedProperties.Item("Position")) + 1)

                            Next

                        Next

                    Next


                Case Is = "dependencies"
                    For Each dino In dinos
                        Dim ser As JObject = JObject.Parse(dino.Object.ToString)
                        Dim datos As List(Of JToken) = ser.Children().ToList
                        Dim doDatos As List(Of JToken)
                        Dim deDatos As List(Of JToken)
                        For Each item As JProperty In datos
                            item.CreateReader()

                            'If item.Name = "ObjectModule" Then
                            '    moduX = item.Value.ToString()
                            '    Continue For
                            'End If

                            If item.Name = "ObjectName" Then
                                xSet.Tables.Add(dino.Key.ToString() & "#" & item.Value.ToString) 'md01#companymaster
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("DepTableCode", GetType(String)) '0
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("DepTableName", GetType(String)) '1
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("DepFieldCode", GetType(String)) '2
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("DepFieldName", GetType(String)) '3
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ConFieldCode", GetType(String)) '4
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ConFieldName", GetType(String)) '5
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ConFieldModule", GetType(String)) '6
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ConFieldObject", GetType(String)) '7
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ConFieldTableCode", GetType(String)) '8
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ConFieldTableName", GetType(String)) '9

                                'Esto se agregó
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ConditionalType", GetType(String)) '10
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ConditionalRule", GetType(String)) '11
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ConditionalValue", GetType(String)) '12
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("MatchingFields", GetType(String)) '13
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ConditionalScope", GetType(String)) '14
                                Continue For
                            End If

                            doDatos = item.Value.Children().ToList()

                            p1 = -1

                            For Each catitem As JProperty In doDatos

                                If catitem.Name = "TableName" Then
                                    tabX = catitem.Value.ToString()
                                    Continue For
                                End If


                                xSet.Tables(xSet.Tables.Count - 1).Rows.Add({item.Name, tabX, "", "", "", "", "", "", "", "", "", "", "", "None", "To Condition"})

                                If p1 < 0 Then p1 = xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1

                                deDatos = catitem.Value.Children.ToList()

                                xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(2) = catitem.Name

                                'k = 3
                                For Each prap As JProperty In deDatos

                                    Select Case prap.Name

                                        Case Is = "MyName"
                                            xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(3) = prap.Value

                                        Case Is = "FieldCode"
                                            xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(4) = prap.Value

                                        Case Is = "FieldName"
                                            xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(5) = prap.Value

                                        Case Is = "Module"
                                            xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(6) = prap.Value

                                        Case Is = "Object"
                                            xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(7) = prap.Value

                                        Case Is = "TableCode"
                                            xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(8) = prap.Value

                                        Case Is = "TableName"
                                            xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(9) = prap.Value

                                        Case Is = "ConditionalType"
                                            xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(10) = prap.Value

                                        Case Is = "ConditionalRule"
                                            xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(11) = prap.Value

                                        Case Is = "ConditionalValue"
                                            xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(12) = prap.Value

                                        Case Is = "MatchingFields"
                                            xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(13) = prap.Value

                                        Case Is = "ConditionalScope"
                                            xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(14) = prap.Value

                                    End Select

                                Next

                            Next

                            If p1 < 0 Then Continue For

                            For i = p1 To xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1
                                xSet.Tables(xSet.Tables.Count - 1).Rows(i).Item(1) = tabX
                            Next

                        Next

                    Next

                Case Is = "records"
                    For Each dino In dinos
                        Dim ser As JObject = JObject.Parse(dino.Object.ToString)
                        Dim datos As List(Of JToken) = ser.Children().ToList
                        Dim doDatos As List(Of JToken)

                        If xSet.Tables.Count = 0 Then
                            xSet.Tables.Add("Records")
                            For Each motem As JProperty In datos
                                'se crea!!
                                motem.CreateReader()
                                xSet.Tables(0).Columns.Add(motem.Name)
                            Next
                        End If

                        xSet.Tables(0).Rows.Add()

                        For i = 0 To xSet.Tables(0).Columns.Count - 1
                            xSet.Tables(0).Rows(xSet.Tables(0).Rows.Count - 1).Item(i) = "" 'vaciamos todo!
                        Next

                        For Each item As JProperty In datos
                            item.CreateReader()
                            k = xSet.Tables(0).Columns.IndexOf(item.Name)
                            If k >= 0 Then
                                xSet.Tables(0).Rows(xSet.Tables(0).Rows.Count - 1).Item(k) = item.Value.ToString()
                            End If

                        Next

                    Next

                Case Is = "construction"

                    For Each dino In dinos
                        Dim ser As JObject = JObject.Parse(dino.Object.ToString)
                        Dim datos As List(Of JToken) = ser.Children().ToList
                        Dim doDatos As List(Of JToken)
                        Dim deDatos As List(Of JToken)
                        Dim diDatos As List(Of JToken)
                        Dim duDatos As List(Of JToken)

                        xSet.Tables.Add(dino.Key)
                        xSet.Tables(xSet.Tables.Count - 1).Columns.Add("TableCode", GetType(String)) '0
                        xSet.Tables(xSet.Tables.Count - 1).Columns.Add("FieldCode", GetType(String)) '1
                        xSet.Tables(xSet.Tables.Count - 1).Columns.Add("RuleKey", GetType(String)) '2
                        xSet.Tables(xSet.Tables.Count - 1).Columns.Add("LevelKey", GetType(String)) '3
                        xSet.Tables(xSet.Tables.Count - 1).Columns.Add("Consec", GetType(Integer)) '4
                        xSet.Tables(xSet.Tables.Count - 1).Columns.Add("Rule", GetType(String)) '5
                        xSet.Tables(xSet.Tables.Count - 1).Columns.Add("Value", GetType(String)) '6
                        xSet.Tables(xSet.Tables.Count - 1).Columns.Add("CharLen", GetType(String)) '7
                        xSet.Tables(xSet.Tables.Count - 1).Columns.Add("Signo", GetType(String)) '8

                        For Each item As JProperty In datos

                            item.CreateReader()

                            tabX = item.Name 'table code

                            doDatos = item.Value.Children().ToList()

                            For Each catitem As JProperty In doDatos
                                'catitem.name
                                moduX = catitem.Name 'field code
                                deDatos = catitem.Value.Children().ToList()

                                For Each yavItem As JProperty In deDatos

                                    miTab = yavItem.Name 'ruleKey
                                    diDatos = yavItem.Value.Children().ToList()

                                    For Each xItem As JProperty In diDatos

                                        LevKey = xItem.Name
                                        duDatos = xItem.Value.Children().ToList()

                                        'aqui agregamos el renglon!
                                        xSet.Tables(xSet.Tables.Count - 1).Rows.Add({tabX, moduX, miTab, LevKey, 0, "", "", "", ""})

                                        For Each yItem As JProperty In duDatos

                                            Select Case yItem.Name
                                                Case Is = "CharLen"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(7) = yItem.Value.ToString()

                                                Case Is = "Consec"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(4) = CInt(yItem.Value.ToString)

                                                Case Is = "Rule"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(5) = yItem.Value.ToString()

                                                Case Is = "Signo"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(8) = yItem.Value.ToString()

                                                Case Is = "Value"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(6) = yItem.Value.ToString()

                                            End Select

                                        Next

                                    Next

                                Next

                            Next

                        Next
                    Next

                Case Is = "templates"
                    'con este vamos a sacar todos los templates!
                    p1 = -1
                    p2 = -1
                    For Each dino In dinos
                        Dim ser As JObject = JObject.Parse(dino.Object.ToString)
                        Dim datos As List(Of JToken) = ser.Children().ToList
                        Dim doDatos As List(Of JToken)
                        Dim deDatos As List(Of JToken)
                        For Each item As JProperty In datos
                            item.CreateReader()

                            If item.Name = "Module" Then
                                moduX = item.Value.ToString
                                Continue For
                            End If

                            If item.Name = "ObjectName" Then

                                Dim iAveprim(1) As DataColumn
                                Dim kEys As New DataColumn()

                                kEys.ColumnName = "KeyField"
                                iAveprim(0) = kEys

                                tabX = item.Value.ToString
                                xSet.Tables.Add(dino.Key.ToString() & "#" & tabX & "#" & moduX) 'md01#companymaster
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add(kEys) '0
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("TableCode", GetType(String)) '1
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("TableName", GetType(String)) '2
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("FieldCode", GetType(String)) '3
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("FieldName", GetType(String)) '4
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("isKey", GetType(String)) '5
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("Position", GetType(Integer)) '6
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("Letter", GetType(String)) '7

                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("MOC", GetType(String)) '8 Mandatory/Optional/Conditional
                                'xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ConditionKind", GetType(String)) '9 From same table, from other table, from other object
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("FillingRule", GetType(String)) '9 A,B,D,E,F
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("DataType", GetType(String)) '10 Tipo de dato
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("MaxChar", GetType(String)) '11 Máximo de caracteres
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ULCase", GetType(String)) '12 Upper/Lower Case
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("Blanks", GetType(String)) '13 Spaces, left, right, both
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("CatalogCode", GetType(String)) '14 Codigo de catalogo
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("CatalogName", GetType(String)) '15 Nombre de catalogo
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ValueColumn", GetType(String)) '16 Columna Valor (para valores fijos)
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("NonRep", GetType(String)) '17 No-repetibilidad?, checkbox, ó N/A en caso de ser campo llave!
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("NonAllowedChars", GetType(String)) '18 Checkbox ó texto de caracteres válidos
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ConditionalPath", GetType(String)) '19 Ruta de condicionante md36>md36-0001>FIELD
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ConditionalObject", GetType(String)) '20 Objeto condicional md36
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ConditionalTable", GetType(String)) '21 Tabla condicionante
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ConditionalField", GetType(String)) '22 Campo condicionante
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ConditionalType", GetType(String)) '23 Internal/External
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ConditionalRule", GetType(String)) '24 Regla condicionante: OR, AND, NULL, STARTWITH,ENDWITH,EXCEPT, FIXED VALUE
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ConditionalValue", GetType(String)) '25 Valor condicionante: Si el objeto condicionante es de otra hoja, que el valor exista en la otra hoja, o que se aplique las reglas de arriba: OR, STARTWITH, ENDWITH, CONTAINS,EXCEPT,FIXED
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ConstructionRule", GetType(String)) '26 Valor
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("MatchingFields", GetType(String)) '27 Valor
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ConditionalScope", GetType(String)) '28 Scope

                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("CatalogModule", GetType(String)) '29 Catalog Module
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("CatMatchField", GetType(String)) '30 Catalog Match Field
                                xSet.Tables(xSet.Tables.Count - 1).Columns.Add("CatMatchConditions", GetType(String)) '31 Catalog Match conditions

                                'Boton en el strip de add conditional rule, boton de add construction rule, 
                                'Cuando se posicione sobre el template completo, se activa el boton add sheet interdependence
                                'Mandatory/Optional/Conditional, Dependant? (Dependant from same table/Dependant from other table/Dependant from other object(este se selecciona en automatico), Regla(A,B,D,E,F), Data Type, MaxChar, UCASE, White Spaces,CatalogCode,CatalogName, ValueColumn ,Non-Repeteability CheckBox,ValidCharactersCheckBox,ConditionalPath,ConditionalObject,ConditionalTable,Conditional field(s),Conditional Rules,ConditionalValue(s)

                                'Construction Rule, Debe tener su propia interfaz-ok!
                                'Parent/Child dpendencie, debe tener su propia interfaz!
                                'Nodo: Construction/build, construction/conditional/, construction/
                                'Por definicion, el campo llave o las columnas con campo llave NO se pueden repetir!!
                                xSet.Tables(xSet.Tables.Count - 1).PrimaryKey = iAveprim
                                Continue For
                            End If

                            doDatos = item.Value.Children().ToList()

                            p1 = -1

                            For Each catitem As JProperty In doDatos

                                Select Case catitem.Name
                                    Case Is = "TableName"
                                        miTab = catitem.Value.ToString
                                        'Siempre noo!
                                        'xSet.Tables(xSet.Tables.Count - 1).Rows.Add({item.Name & "#" & catitem.Name, item.Name, miTab, catitem.Name, "", "", 0, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "None", "To Condition", "", "", ""})

                                    Case Else

                                        'xSet.Tables(xSet.Tables.Count - 1).Rows.Add({item.Name & "#" & catitem.Name, item.Name, "", catitem.Name, "", "", 0, ""})
                                        xSet.Tables(xSet.Tables.Count - 1).Rows.Add({item.Name & "#" & catitem.Name, item.Name, "", catitem.Name, "", "", 0, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "None", "To Condition", "", "", ""})

                                        deDatos = catitem.Value.Children.ToList()

                                        For Each prap As JProperty In deDatos

                                            Select Case prap.Name
                                                Case Is = "Letter"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(7) = prap.Value.ToString()

                                                Case Is = "Name"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(4) = prap.Value.ToString()

                                                Case Is = "Position"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(6) = CInt(prap.Value.ToString())

                                                Case Is = "isKey"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(5) = prap.Value.ToString()

                                                Case Is = "MOC"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(8) = prap.Value.ToString()

                                                Case Is = "FillingRule"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(9) = prap.Value.ToString()

                                                Case Is = "DataType"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(10) = prap.Value.ToString()

                                                Case Is = "MaxChar"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(11) = prap.Value.ToString()

                                                Case Is = "ULCase"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(12) = prap.Value.ToString()

                                                Case Is = "Blanks"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(13) = prap.Value.ToString()

                                                Case Is = "CatalogCode"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(14) = prap.Value.ToString()

                                                Case Is = "CatalogName"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(15) = prap.Value.ToString()

                                                Case Is = "ValueColumn"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(16) = prap.Value.ToString()

                                                Case Is = "NonRep"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(17) = prap.Value.ToString()

                                                Case Is = "NonAllowedChars"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(18) = prap.Value.ToString()

                                                Case Is = "ConditionalPath"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(19) = prap.Value.ToString()

                                                Case Is = "ConditionalObject"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(20) = prap.Value.ToString()

                                                Case Is = "ConditionalTable"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(21) = prap.Value.ToString()

                                                Case Is = "ConditionalField"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(22) = prap.Value.ToString()

                                                Case Is = "ConditionalType"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(23) = prap.Value.ToString()

                                                Case Is = "ConditionalRule"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(24) = prap.Value.ToString()

                                                Case Is = "ConditionalValue"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(25) = prap.Value.ToString()

                                                Case Is = "Construction"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(26) = prap.Value.ToString()

                                                Case Is = "MatchingFields"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(27) = prap.Value.ToString()

                                                Case Is = "ConditionalScope"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(28) = prap.Value.ToString()

                                                Case Is = "CatalogModule"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(29) = prap.Value.ToString()

                                                Case Is = "CatMatchField"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(30) = prap.Value.ToString()

                                                Case Is = "CatMatchConditions"
                                                    xSet.Tables(xSet.Tables.Count - 1).Rows(xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1).Item(31) = prap.Value.ToString()

                                            End Select

                                        Next

                                        If p1 < 0 Then p1 = xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1

                                End Select

                            Next

                            If p1 < 0 Then Continue For

                            For i = p1 To xSet.Tables(xSet.Tables.Count - 1).Rows.Count - 1
                                xSet.Tables(xSet.Tables.Count - 1).Rows(i).Item(2) = miTab
                            Next

                        Next


                    Next


                Case Is = "users"

                    Dim iAveprim(1) As DataColumn
                    Dim kEys As New DataColumn()

                    kEys.ColumnName = "Email"
                    iAveprim(0) = kEys

                    xSet.Tables.Add("Users")
                    xSet.Tables(0).Columns.Add(kEys) '0 Email
                    xSet.Tables(0).Columns.Add("Name", GetType(String)) '1 correo
                    xSet.Tables(0).Columns.Add("LastName", GetType(String)) '2 correo
                    xSet.Tables(0).Columns.Add("FirstTime", GetType(String)) '3 correo
                    xSet.Tables(0).Columns.Add("Pass", GetType(String)) '4 password
                    xSet.Tables(0).Columns.Add("Role", GetType(String)) '5 correo
                    xSet.Tables(0).Columns.Add("Modules", GetType(String)) '6 correo
                    xSet.Tables(0).Columns.Add("Key", GetType(String)) '7 correo
                    xSet.Tables(0).PrimaryKey = iAveprim

                    For Each dino In dinos

                        If dino.Key = "Name" Then Continue For

                        Dim ser As JObject = JObject.Parse(dino.Object.ToString)

                        Dim datos As List(Of JToken) = ser.Children().ToList
                        'Dim doDatos As List(Of JToken)

                        For Each item As JProperty In datos
                            item.CreateReader()

                            Select Case item.Name
                                Case Is = "Email"
                                    userCorreo = item.Value.ToString()

                                Case Is = "LastName"
                                    userApellido = item.Value.ToString()

                                Case Is = "Name"
                                    userName = item.Value.ToString()

                                Case Is = "FirstTime"
                                    userFirst = item.Value.ToString()

                                Case Is = "Modules"
                                    userModules = item.Value.ToString()

                                Case Is = "Pass"
                                    userPass = item.Value.ToString()

                                Case Is = "Role"
                                    userRole = item.Value.ToString()

                            End Select

                        Next

                        xSet.Tables(0).Rows.Add({userCorreo, userName, userApellido, userFirst, userPass, userRole, userModules, dino.Key})

                    Next


                Case Is = "modules"
                    Dim iAveprim(1) As DataColumn
                    Dim kEys As New DataColumn()

                    kEys.ColumnName = "KeyField"
                    iAveprim(0) = kEys

                    xSet.Tables.Add("Modules")
                    xSet.Tables(0).Columns.Add(kEys) '0 ModuleCode
                    xSet.Tables(0).Columns.Add("Name", GetType(String)) '1 correo
                    xSet.Tables(0).PrimaryKey = iAveprim

                    For Each dino In dinos

                        'If dino.Key = "Name" Then Continue For

                        Dim ser As JObject = JObject.Parse(dino.Object.ToString)

                        Dim datos As List(Of JToken) = ser.Children().ToList

                        For Each item As JProperty In datos
                            item.CreateReader()
                            xSet.Tables(0).Rows.Add({dino.Key, item.Value.ToString()})
                        Next

                    Next


                Case Is = "relations"

                    Dim miAncho As Integer = 0
                    Dim unRow As Integer = 0
                    Dim unCampo As String = ""
                    Dim unPapa As String = ""
                    Dim nomPapa As String = ""
                    Dim nomCampo As String = ""
                    Dim siHubo As Boolean = False

                    For Each dino In dinos
                        Dim ser As JObject = JObject.Parse(dino.Object.ToString)
                        Dim datos As List(Of JToken) = ser.Children().ToList
                        Dim doDatos As List(Of JToken)
                        Dim deDatos As List(Of JToken)
                        Dim diDatos As List(Of JToken)
                        Dim duDatos As List(Of JToken)
                        Dim xiDatos As List(Of JToken)

                        xSet.Tables.Add(dino.Key)
                        xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ChildTableCode", GetType(String)) '0
                        xSet.Tables(xSet.Tables.Count - 1).Columns.Add("RuleCode", GetType(String)) '1
                        xSet.Tables(xSet.Tables.Count - 1).Columns.Add("RuleName", GetType(String)) '2
                        xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ParentTableCode", GetType(String)) '3

                        xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ChildField", GetType(String)) '4
                        xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ChildFieldName", GetType(String)) '5
                        xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ParentField", GetType(String)) '6
                        xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ParentFieldName", GetType(String)) '7
                        xSet.Tables(xSet.Tables.Count - 1).Columns.Add("Consec", GetType(Integer)) '8
                        xSet.Tables(xSet.Tables.Count - 1).Columns.Add("NumberOfColumns", GetType(Integer)) '9

                        'xSet.Tables(xSet.Tables.Count - 1).ExtendedProperties.Add("NumberOfColumns", CInt(0))

                        For Each item As JProperty In datos
                            item.CreateReader()

                            'creamos la tabla directo!

                            doDatos = item.Value.Children().ToList()

                            moduX = ""
                            miTab = ""
                            LevKey = ""
                            rulKey = ""
                            miAncho = 0

                            For Each catitem As JProperty In doDatos

                                siHubo = False

                                moduX = catitem.Name 'table code

                                deDatos = catitem.Value.Children().ToList()

                                For Each opItem As JProperty In deDatos

                                    miTab = opItem.Name 'rule key!

                                    Select Case opItem.Name

                                        Case Is = "ParentTable"
                                            LevKey = opItem.Value.ToString()

                                        Case Is = "RuleName"
                                            rulKey = opItem.Value.ToString() 'rulename
                                                'se supone que aqui entraría para agregar los renglones!

                                        Case Is = "NumberOfColumns"
                                            miAncho = CInt(opItem.Value.ToString())
                                            'xSet.Tables(xSet.Tables.Count - 1).ExtendedProperties.Item("NumberOfColumns") = miAncho

                                        Case Is = "logic"

                                            siHubo = True

                                            duDatos = opItem.Value.Children().ToList()

                                            unRow = 0
                                            unCampo = ""
                                            unPapa = ""
                                            nomCampo = ""
                                            nomPapa = ""

                                            For Each xItemo As JProperty In duDatos

                                                xiDatos = xItemo.Value.Children().ToList()

                                                For Each yItem As JProperty In xiDatos

                                                    Select Case yItem.Name
                                                        Case Is = "Consec"
                                                            unRow = CInt(yItem.Value.ToString())

                                                        Case Is = "ChildField"
                                                            unCampo = yItem.Value.ToString()

                                                        Case Is = "ParentField"
                                                            unPapa = yItem.Value.ToString()

                                                        Case Is = "ChildFieldName"
                                                            nomCampo = yItem.Value.ToString()

                                                        Case Is = "ParentFieldName"
                                                            nomPapa = yItem.Value.ToString()

                                                    End Select

                                                Next

                                                xSet.Tables(xSet.Tables.Count - 1).Rows.Add({item.Name, moduX, rulKey, LevKey, unCampo, nomCampo, unPapa, nomPapa, unRow, miAncho})

                                            Next


                                    End Select

                                Next

                                If siHubo = False Then
                                    'agregamos un renglon!
                                    'OJO, esto aqui abajo SI estaba!
                                    xSet.Tables(xSet.Tables.Count - 1).Rows.Add({item.Name, moduX, rulKey, LevKey, "", "", "", "", 0, 0})
                                End If

                            Next

                        Next

                    Next


            End Select

        Catch ex As Exception


        End Try


        Return xSet


    End Function

    Public Async Function RecurFireBaseTree(ByVal elUrl As String) As Task(Of IReadOnlyCollection(Of FirebaseObject(Of Object)))
        Dim client = New FirebaseClient(elUrl)
        'Dim dinos As FirebaseObject(Of Object)

        Try
            Dim dinos = Await client.Child("").OnceAsync(Of Object)
            Return dinos

        Catch ex As Exception

            'dinos = Nothing
            Return Nothing
        End Try



    End Function

    Public Async Function JalaWS(ByVal elUrl As String) As Task(Of String)

        Dim laResp As String = ""

        Try
            Dim s As HttpWebRequest
            Dim enc As UTF8Encoding
            'Dim postdata As String
            Dim postdatabytes As Byte()
            s = HttpWebRequest.Create(elUrl)
            enc = New System.Text.UTF8Encoding()
            'postdata = "username=*****&password=*****&message=test+message&orig=test&number=447712345678"
            postdatabytes = enc.GetBytes(elUrl)
            s.Method = "POST"
            s.ContentType = "application/x-www-form-urlencoded"
            s.ContentLength = postdatabytes.Length

            Using stream = s.GetRequestStream()
                'stream.Write(postdatabytes, 0, postdatabytes.Length)
                Await stream.WriteAsync(postdatabytes, 0, postdatabytes.Length)
            End Using

            Using tResponse As WebResponse = s.GetResponse()
                Using dataStreamResponse As Stream = tResponse.GetResponseStream()
                    Using tReader As StreamReader = New StreamReader(dataStreamResponse)
                        Dim sResponseFromServer As String = tReader.ReadToEnd()
                        Dim str As String = sResponseFromServer

                        laResp = str

                    End Using
                End Using
            End Using

        Catch ex As Exception

            laResp = ""

        End Try


        Return laResp

    End Function

    Public Async Function Jala() As Task
        'este si funciona!!
        Dim client = New FirebaseClient(RaizFire)

        Dim dinos = Await client.Child("").OnceAsync(Of Object)

        For Each dino In dinos
            Console.WriteLine($"{dino.Key} is {dino.Object.ToString()}m high.")
        Next

    End Function

    Public Async Function Escribe() As Task

        Dim client = New FirebaseClient(RaizFire)
        Dim elDato As String = ""
        elDato = "{" & vbCrLf
        elDato = elDato & """VZ"":""Venezuela"""
        elDato = elDato & vbCrLf & "}"

        Dim miDato As String = ""
        miDato = """VZ"":""Venezuela"""

        Dim err As String = ""

        Try
            Dim dino = Await client.Child("global").Child("gb-aaaa").PostAsync(miDato)
        Catch ex As Exception
            err = ex.Message
        End Try


    End Function

    Public Async Sub hazPut()
        Dim err As String = ""
        Dim client = New FirebaseClient(RaizFire)
        'Dim quer = FirebaseQuery
        Dim elDato As String = ""
        elDato = "{" & vbCrLf
        elDato = elDato & """MX"":""Mexico"""
        elDato = elDato & vbCrLf & "}" 'este funciona pero agrega hijos!!

        Dim miDato As String = ""
        'miDato = """VZ"":""Venezuela"""

        'miDato = "{" & vbCrLf
        miDato = miDato & """Venezuela""" 'Este funciono!!, directo a root!
        'miDato = miDato & vbCrLf & "}"

        Dim cadeConv = JsonConvert.SerializeObject(elDato)
        'Dim ser As JObject = JObject.Parse("")
        'client.Child("").PutAsync("").ConfigureAwait(False)

        Try
            'Await client.Child("global").Child("gb-aaaa").PutAsync(miDato).ConfigureAwait(False)
            Await client.Child("global").Child("gb-aaaa").Child("VZ").PutAsync(miDato).ConfigureAwait(False)

            'Await client.Child("global").Child("gb-aaaa").PutAsync(elDato).ConfigureAwait(False)

        Catch ex As Exception
            err = ex.Message
        End Try

        'Invalid data; couldn't parse JSON object, array, or value.

    End Sub

    Public Async Function HazPutEnFireBase(ByVal elPath As String, ByVal elSet As DataSet) As Task(Of String)
        Dim client = New FirebaseClient(elPath)

        Dim i As Integer
        Dim j As Integer
        Dim miDato As String = ""
        Dim laYave As String = ""
        Dim err As String = ""
        For i = 0 To elSet.Tables(0).Rows.Count - 1
            laYave = elSet.Tables(0).Rows(i).Item(0)
            miDato = """" & elSet.Tables(0).Rows(i).Item(1) & """"
            Try
                Await client.Child(laYave).PutAsync(miDato).ConfigureAwait(False)
            Catch ex As Exception
                err = ex.Message
            End Try
        Next

        Return "ok"

    End Function

    Public Sub CreaDBLetraNumero()

        matAlpha(1) = "A"
        matAlpha(2) = "B"
        matAlpha(3) = "C"
        matAlpha(4) = "D"
        matAlpha(5) = "E"
        matAlpha(6) = "F"
        matAlpha(7) = "G"
        matAlpha(8) = "H"
        matAlpha(9) = "I"
        matAlpha(10) = "J"
        matAlpha(11) = "K"
        matAlpha(12) = "L"
        matAlpha(13) = "M"
        matAlpha(14) = "N"
        matAlpha(15) = "O"
        matAlpha(16) = "P"
        matAlpha(17) = "Q"
        matAlpha(18) = "R"
        matAlpha(19) = "S"
        matAlpha(20) = "T"
        matAlpha(21) = "U"
        matAlpha(22) = "V"
        matAlpha(23) = "W"
        matAlpha(24) = "X"
        matAlpha(25) = "Y"
        matAlpha(26) = "Z"

        Dim iAveprim(1) As DataColumn
        Dim kEys As New DataColumn()
        kEys.ColumnName = "Position"
        kEys.DataType = GetType(Integer)
        iAveprim(0) = kEys

        LetraNumero.Tables.Clear()
        LetraNumero.Tables.Add("Letra")
        LetraNumero.Tables(0).Columns.Add(kEys)
        LetraNumero.Tables(0).Columns.Add("Letter", GetType(String))
        LetraNumero.Tables(0).PrimaryKey = iAveprim

        Dim cadStr As String = ""
        Dim k As Integer = 0
        Dim n As Integer = 1

        For i = 1 To 26 * 5
            k = k + 1
            If n = 1 Then
                'primera vuelta
                LetraNumero.Tables(0).Rows.Add({i, matAlpha(k)})
            Else
                LetraNumero.Tables(0).Rows.Add({i, matAlpha(n - 1) & matAlpha(k)})
            End If

            If k = 26 Then
                n = n + 1
                k = 0
            End If

        Next


    End Sub

    Public Function getSHA1Hash(ByVal strToHash As String) As String

        Dim sha1Obj As New SHA1CryptoServiceProvider
        Dim bytesToHash() As Byte = System.Text.Encoding.ASCII.GetBytes(strToHash)

        bytesToHash = sha1Obj.ComputeHash(bytesToHash)

        Dim strResult As String = ""

        For Each b As Byte In bytesToHash
            strResult += b.ToString("x2")
        Next

        Return strResult

    End Function

    Public Sub AsignaYavePrimariaATabla(ByRef xTab As DataTable, ByVal nombreColumna As String, ByVal esCaseSens As Boolean)
        Dim iAveprim(1) As DataColumn

        Dim kEys As New DataColumn()

        kEys.ColumnName = nombreColumna

        iAveprim(0) = kEys

        xTab.Columns.Add(kEys)

        xTab.CaseSensitive = esCaseSens

        xTab.PrimaryKey = iAveprim

    End Sub
    Public Function TengoCataMatch(ByVal elDt As DataTable, ByVal laCadena As String, ByVal lasCondis As String, ByVal elFieldMach As String, ByRef cadeMach As String) As Boolean

        'el data table para la comparación
        Dim querCom As String = ""
        Dim cadPruf As String = ""
        Dim siEncontre As Boolean = False
        cadeMach = ""
        For i = 1 To laCadena.Length
            cadPruf = laCadena.Substring(0, i)
            querCom = lasCondis & elFieldMach & " = '" & cadPruf & "'"
            Dim misRows() As DataRow
            misRows = elDt.Select(querCom)
            If misRows.Length >= 1 Then
                'ya dimos!
                cadeMach = cadPruf
                siEncontre = True
                Exit For
            End If

        Next


        Return siEncontre

    End Function

    Public Function HayCadenainsideCadena(ByVal cadGral As String, ByVal cadBusca As String, ByRef cadResult As String) As Boolean

        'cadgral es cadleft
        Dim i As Integer
        Dim cadQuer As String = ""
        Dim siHayMach As Boolean = False
        For i = 1 To cadGral.Length
            cadQuer = cadGral.Substring(0, i)

            If cadQuer = cadBusca Then
                cadResult = cadQuer
                siHayMach = True
                Exit For
            End If

        Next

        Return siHayMach

    End Function


    Public Async Function PullDtFb(ByVal DtPath As String, ByVal elConcepto As String) As Task(Of DataSet)

        Dim xDt As New DataSet
        Dim client = New FirebaseClient(DtPath)

        Try

            Dim dinos = Await client.Child("").OnceAsync(Of Object)

            Select Case elConcepto

                Case Is = "inuse"

                    xDt.Tables.Add()
                    xDt.Tables(0).Columns.Add("LastUsed", GetType(String)) '0
                    xDt.Tables(0).Columns.Add("TableName", GetType(String)) '1
                    xDt.Tables(0).Columns.Add("User", GetType(String)) '2
                    xDt.Tables(0).Columns.Add("UserName", GetType(String)) '3
                    xDt.Tables(0).Columns.Add("inUse", GetType(String)) '4
                    xDt.Tables(0).ExtendedProperties.Add("inEdit", False)
                    xDt.Tables(0).ExtendedProperties.Add("Key", "")

                    xDt.Tables(0).Rows.Add({"", "", "", "", "X"}) 'iniciamos con la idea de que NO se puede editar!

                    If dinos.Count = 0 Then
                        'NO existía!, le ponemos false!
                        xDt.Tables(0).Rows(0).Item(4) = ""
                        Return xDt
                    End If

                    For Each dino In dinos

                        Dim ser As JObject = JObject.Parse(dino.Object.ToString)
                        Dim datos As List(Of JToken) = ser.Children().ToList

                        xDt.Tables(0).ExtendedProperties.Item("Key") = dino.Key

                        For Each item As JProperty In datos
                            Select Case item.Name

                                Case Is = "LastUsed"
                                    xDt.Tables(0).Rows(xDt.Tables(0).Rows.Count - 1).Item(0) = item.Value.ToString()

                                Case Is = "TableName"
                                    xDt.Tables(0).Rows(xDt.Tables(0).Rows.Count - 1).Item(1) = item.Value.ToString()

                                Case Is = "User"
                                    xDt.Tables(0).Rows(xDt.Tables(0).Rows.Count - 1).Item(2) = item.Value.ToString()

                                Case Is = "UserName"
                                    xDt.Tables(0).Rows(xDt.Tables(0).Rows.Count - 1).Item(3) = item.Value.ToString()

                                Case Is = "inUse"
                                    xDt.Tables(0).Rows(xDt.Tables(0).Rows.Count - 1).Item(4) = item.Value.ToString()

                            End Select

                        Next

                    Next

                    'caso extraordinario!, soy yo mismo!, puedo seguir editando
                    'If xDt.Tables(0).Rows(0).Item(2) = UsuarioCorreo Then
                    '    xDt.Tables(0).Rows(0).Item(4) = "" 'puede re-editar
                    'End If


                Case Is = "otro"



            End Select

        Catch ex As Exception



        End Try

        'si no tiene tablas o renglones entonces NO existia!!, no hay bronca!!

        Return xDt

    End Function

    Public Async Function HazPutConPathRowsyCols(ByVal elPath As String, ByVal elSet As DataTable, ByVal colYave As Integer) As Task(Of String)
        If elSet.Rows.Count = 0 Then Return "No records to write!"

        Dim client = New FirebaseClient(elPath)
        Dim laYave As String = ""
        Dim miDato As String = ""
        Dim elHijo As String = ""

        Dim elRegreso As String = ""
        Dim cuentaErrores As Long = 0
        Dim cuentaOk As Long = 0
        Dim errLog As String = ""
        For i = 0 To elSet.Rows.Count - 1

            If colYave >= 0 Then
                elHijo = elSet.Rows(i).Item(colYave)
            Else
                elHijo = ""
            End If

            For j = 0 To elSet.Columns.Count - 1

                If j = colYave Then Continue For

                laYave = elSet.Columns(j).ColumnName '0,1
                miDato = """" & elSet.Rows(i).Item(j) & """"

                Try

                    Await client.Child(elHijo).Child(laYave).PutAsync(miDato).ConfigureAwait(False)
                    cuentaOk = cuentaOk + 1
                Catch ex As Exception

                    cuentaErrores = cuentaErrores + 1
                    errLog = errLog & "Error writting on node: " & elPath & " > " & laYave & " > " & miDato & " : " & ex.Message & vbCrLf

                End Try

            Next

        Next


        If cuentaOk > 0 Then
            elRegreso = elRegreso & cuentaOk & " records added!" & vbCrLf & vbCrLf
        End If

        If cuentaErrores > 0 Then
            elRegreso = elRegreso & cuentaErrores & " Errors on writting..." & vbCrLf
            elRegreso = elRegreso & errLog
        End If

        Return elRegreso

    End Function

    Public Async Function HazPost1Set(ByVal elPath As String, ByVal elSet As DataTable, ByVal colYave As Integer) As Task(Of String)
        If elSet.Rows.Count = 0 Then Return "No records to write!"

        Dim client = New FirebaseClient(elPath)
        Dim laYave As String = ""
        Dim miDato As String = ""
        Dim elHijo As String = ""

        Dim elRegreso As String = "fail"
        Dim cuentaErrores As Long = 0
        Dim cuentaOk As Long = 0
        Dim errLog As String = ""

        Dim elJuntos As String = ""

        For i = 0 To elSet.Rows.Count - 1

            elJuntos = "{" & vbCrLf
            For j = 0 To elSet.Columns.Count - 1

                If j <> 0 Then elJuntos = elJuntos & "," & vbCrLf

                laYave = elSet.Columns(j).ColumnName
                miDato = elSet.Rows(i).Item(j)

                elJuntos = elJuntos & """" & laYave & """"
                elJuntos = elJuntos & ":" & """" & miDato & """"

            Next
            elJuntos = elJuntos & vbCrLf & "}"

            Try
                Dim fbObj = Await client.Child("").PostAsync(elJuntos, False)
                elRegreso = fbObj.Key
                cuentaOk = cuentaOk + 1
            Catch ex As Exception
                cuentaErrores = cuentaErrores + 1
                errLog = errLog & "Error writting on node: " & elPath & " > " & laYave & " > " & miDato & " : " & ex.Message & vbCrLf
            End Try

        Next

        Return elRegreso

    End Function

    Public Async Function PullDtFireBase(ByVal DtPath As String, ByVal elConcepto As String, Optional ByVal elHijo As String = "") As Task(Of DataTable)

        Dim xDt As New DataTable
        Dim client = New FirebaseClient(DtPath)

        Try

            Dim dinos = Await client.Child("").OnceAsync(Of Object)

            Select Case elConcepto

                Case Is = "tempfields"


                    AsignaYavePrimariaATabla(xDt, "FieldCode", False)
                    xDt.Columns.Add("FieldName", GetType(String))

                    For Each dino In dinos

                        If dino.Key = "TableName" Then Continue For

                        xDt.Rows.Add({dino.Key, ""})

                    Next


                Case Is = "tempunit"

                    Dim moduX As String = ""
                    Dim tabX As String = ""
                    Dim miTab As String = ""
                    Dim p1 As Integer = -1

                    For Each dino In dinos

                        Select Case dino.Key
                            Case Is = "Module"
                                moduX = dino.Object.ToString
                                Continue For

                            Case Is = "ObjectName"
                                Dim iAveprim(1) As DataColumn
                                Dim kEys As New DataColumn()

                                kEys.ColumnName = "KeyField"
                                iAveprim(0) = kEys

                                tabX = dino.Object.ToString
                                xDt.TableName = elHijo & "#" & tabX & "#" & moduX
                                'xSet.Tables.Add(dino.Key.ToString() & "#" & tabX & "#" & moduX) 'md01#companymaster
                                xDt.Columns.Add(kEys) '0
                                xDt.Columns.Add("TableCode", GetType(String)) '1
                                xDt.Columns.Add("TableName", GetType(String)) '2
                                xDt.Columns.Add("FieldCode", GetType(String)) '3
                                xDt.Columns.Add("FieldName", GetType(String)) '4
                                xDt.Columns.Add("isKey", GetType(String)) '5
                                xDt.Columns.Add("Position", GetType(Integer)) '6
                                xDt.Columns.Add("Letter", GetType(String)) '7

                                xDt.Columns.Add("MOC", GetType(String)) '8 Mandatory/Optional/Conditional
                                'xSet.Tables(xSet.Tables.Count - 1).Columns.Add("ConditionKind", GetType(String)) '9 From same table, from other table, from other object
                                xDt.Columns.Add("FillingRule", GetType(String)) '9 A,B,D,E,F
                                xDt.Columns.Add("DataType", GetType(String)) '10 Tipo de dato
                                xDt.Columns.Add("MaxChar", GetType(String)) '11 Máximo de caracteres
                                xDt.Columns.Add("ULCase", GetType(String)) '12 Upper/Lower Case
                                xDt.Columns.Add("Blanks", GetType(String)) '13 Spaces, left, right, both
                                xDt.Columns.Add("CatalogCode", GetType(String)) '14 Codigo de catalogo
                                xDt.Columns.Add("CatalogName", GetType(String)) '15 Nombre de catalogo
                                xDt.Columns.Add("ValueColumn", GetType(String)) '16 Columna Valor (para valores fijos)
                                xDt.Columns.Add("NonRep", GetType(String)) '17 No-repetibilidad?, checkbox, ó N/A en caso de ser campo llave!
                                xDt.Columns.Add("NonAllowedChars", GetType(String)) '18 Checkbox ó texto de caracteres válidos
                                xDt.Columns.Add("ConditionalPath", GetType(String)) '19 Ruta de condicionante md36>md36-0001>FIELD
                                xDt.Columns.Add("ConditionalObject", GetType(String)) '20 Objeto condicional md36
                                xDt.Columns.Add("ConditionalTable", GetType(String)) '21 Tabla condicionante
                                xDt.Columns.Add("ConditionalField", GetType(String)) '22 Campo condicionante
                                xDt.Columns.Add("ConditionalType", GetType(String)) '23 Internal/External
                                xDt.Columns.Add("ConditionalRule", GetType(String)) '24 Regla condicionante: OR, AND, NULL, STARTWITH,ENDWITH,EXCEPT, FIXED VALUE
                                xDt.Columns.Add("ConditionalValue", GetType(String)) '25 Valor condicionante: Si el objeto condicionante es de otra hoja, que el valor exista en la otra hoja, o que se aplique las reglas de arriba: OR, STARTWITH, ENDWITH, CONTAINS,EXCEPT,FIXED
                                xDt.Columns.Add("ConstructionRule", GetType(String)) '26 Valor
                                xDt.Columns.Add("MatchingFields", GetType(String)) '27 Valor
                                xDt.Columns.Add("ConditionalScope", GetType(String)) '28 Scope

                                xDt.Columns.Add("CatalogModule", GetType(String)) '29 Catalog Module
                                xDt.Columns.Add("CatMatchField", GetType(String)) '30 Catalog Match Field
                                xDt.Columns.Add("CatMatchConditions", GetType(String)) '31 Catalog Match conditions

                                xDt.PrimaryKey = iAveprim
                                Continue For

                            Case Else

                                p1 = -1

                                Dim ser As JObject = JObject.Parse(dino.Object.ToString)
                                Dim datos As List(Of JToken) = ser.Children().ToList
                                Dim doDatos As List(Of JToken)

                                For Each item As JProperty In datos

                                    Select Case item.Name
                                        Case Is = "TableName"
                                            miTab = item.Value.ToString
                                            'Siempre noo!
                                            'xSet.Tables(xSet.Tables.Count - 1).Rows.Add({item.Name & "#" & catitem.Name, item.Name, miTab, catitem.Name, "", "", 0, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "None", "To Condition", "", "", ""})

                                        Case Else

                                            'xSet.Tables(xSet.Tables.Count - 1).Rows.Add({item.Name & "#" & catitem.Name, item.Name, "", catitem.Name, "", "", 0, ""})
                                            xDt.Rows.Add({dino.Key & "#" & item.Name, dino.Key, "", item.Name, "", "", 0, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "None", "To Condition", "", "", ""})

                                            doDatos = item.Value.Children.ToList()

                                            For Each prap As JProperty In doDatos

                                                Select Case prap.Name
                                                    Case Is = "Letter"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(7) = prap.Value.ToString()

                                                    Case Is = "Name"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(4) = prap.Value.ToString()

                                                    Case Is = "Position"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(6) = CInt(prap.Value.ToString())

                                                    Case Is = "isKey"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(5) = prap.Value.ToString()

                                                    Case Is = "MOC"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(8) = prap.Value.ToString()

                                                    Case Is = "FillingRule"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(9) = prap.Value.ToString()

                                                    Case Is = "DataType"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(10) = prap.Value.ToString()

                                                    Case Is = "MaxChar"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(11) = prap.Value.ToString()

                                                    Case Is = "ULCase"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(12) = prap.Value.ToString()

                                                    Case Is = "Blanks"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(13) = prap.Value.ToString()

                                                    Case Is = "CatalogCode"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(14) = prap.Value.ToString()

                                                    Case Is = "CatalogName"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(15) = prap.Value.ToString()

                                                    Case Is = "ValueColumn"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(16) = prap.Value.ToString()

                                                    Case Is = "NonRep"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(17) = prap.Value.ToString()

                                                    Case Is = "NonAllowedChars"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(18) = prap.Value.ToString()

                                                    Case Is = "ConditionalPath"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(19) = prap.Value.ToString()

                                                    Case Is = "ConditionalObject"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(20) = prap.Value.ToString()

                                                    Case Is = "ConditionalTable"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(21) = prap.Value.ToString()

                                                    Case Is = "ConditionalField"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(22) = prap.Value.ToString()

                                                    Case Is = "ConditionalType"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(23) = prap.Value.ToString()

                                                    Case Is = "ConditionalRule"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(24) = prap.Value.ToString()

                                                    Case Is = "ConditionalValue"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(25) = prap.Value.ToString()

                                                    Case Is = "Construction"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(26) = prap.Value.ToString()

                                                    Case Is = "MatchingFields"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(27) = prap.Value.ToString()

                                                    Case Is = "ConditionalScope"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(28) = prap.Value.ToString()

                                                    Case Is = "CatalogModule"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(29) = prap.Value.ToString()

                                                    Case Is = "CatMatchField"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(30) = prap.Value.ToString()

                                                    Case Is = "CatMatchConditions"
                                                        xDt.Rows(xDt.Rows.Count - 1).Item(31) = prap.Value.ToString()

                                                End Select

                                            Next

                                            If p1 < 0 Then p1 = xDt.Rows.Count - 1

                                    End Select

                                Next

                                If p1 < 0 Then Continue For

                                For i = p1 To xDt.Rows.Count - 1
                                    xDt.Rows(i).Item(2) = miTab
                                Next

                        End Select


                    Next


                Case Is = "fieldunit"

                    xDt.Columns.Add("Blanks", GetType(String)) '0
                    xDt.Columns.Add("CatMatchConditions", GetType(String)) '1
                    xDt.Columns.Add("CatMatchField", GetType(String)) '2
                    xDt.Columns.Add("CatalogCode", GetType(String)) '3
                    xDt.Columns.Add("CatalogModule", GetType(String)) '4
                    xDt.Columns.Add("CatalogName", GetType(String)) '5
                    xDt.Columns.Add("DataType", GetType(String)) '6
                    xDt.Columns.Add("FillingRule", GetType(String)) '7
                    xDt.Columns.Add("Letter", GetType(String)) '8
                    xDt.Columns.Add("MOC", GetType(String)) '9
                    xDt.Columns.Add("MaxChar", GetType(String)) '10
                    xDt.Columns.Add("Name", GetType(String)) '11
                    xDt.Columns.Add("NonAllowedChars", GetType(String)) '12
                    xDt.Columns.Add("NonRep", GetType(String)) '13
                    xDt.Columns.Add("Position", GetType(String)) '14
                    xDt.Columns.Add("ULCase", GetType(String)) '15
                    xDt.Columns.Add("ValueColumn", GetType(String)) '16
                    xDt.Columns.Add("isKey", GetType(String)) '17

                    xDt.Rows.Add({"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""})

                    For Each dino In dinos

                        Select Case dino.Key
                            Case Is = "Blanks"
                                xDt.Rows(xDt.Rows.Count - 1).Item(0) = dino.Object.ToString()

                            Case Is = "CatMatchConditions"
                                xDt.Rows(xDt.Rows.Count - 1).Item(1) = dino.Object.ToString()

                            Case Is = "CatMatchField"
                                xDt.Rows(xDt.Rows.Count - 1).Item(2) = dino.Object.ToString()

                            Case Is = "CatalogCode"
                                xDt.Rows(xDt.Rows.Count - 1).Item(3) = dino.Object.ToString()

                            Case Is = "CatalogModule"
                                xDt.Rows(xDt.Rows.Count - 1).Item(4) = dino.Object.ToString()

                            Case Is = "CatalogName"
                                xDt.Rows(xDt.Rows.Count - 1).Item(5) = dino.Object.ToString()

                            Case Is = "DataType"
                                xDt.Rows(xDt.Rows.Count - 1).Item(6) = dino.Object.ToString()

                            Case Is = "FillingRule"
                                xDt.Rows(xDt.Rows.Count - 1).Item(7) = dino.Object.ToString()

                            Case Is = "Letter"
                                xDt.Rows(xDt.Rows.Count - 1).Item(8) = dino.Object.ToString()

                            Case Is = "MOC"
                                xDt.Rows(xDt.Rows.Count - 1).Item(9) = dino.Object.ToString()

                            Case Is = "MaxChar"
                                xDt.Rows(xDt.Rows.Count - 1).Item(10) = dino.Object.ToString()

                            Case Is = "Name"
                                xDt.Rows(xDt.Rows.Count - 1).Item(11) = dino.Object.ToString()

                            Case Is = "NonAllowedChars"
                                xDt.Rows(xDt.Rows.Count - 1).Item(12) = dino.Object.ToString()

                            Case Is = "NonRep"
                                xDt.Rows(xDt.Rows.Count - 1).Item(13) = dino.Object.ToString()

                            Case Is = "Position"
                                xDt.Rows(xDt.Rows.Count - 1).Item(14) = dino.Object.ToString()

                            Case Is = "ULCase"
                                xDt.Rows(xDt.Rows.Count - 1).Item(15) = dino.Object.ToString()

                            Case Is = "ValueColumn"
                                xDt.Rows(xDt.Rows.Count - 1).Item(16) = dino.Object.ToString()

                            Case Is = "isKey"
                                xDt.Rows(xDt.Rows.Count - 1).Item(17) = dino.Object.ToString()

                        End Select

                    Next


                Case Is = "otro"


            End Select

        Catch ex As Exception



        End Try

        'si no tiene tablas o renglones entonces NO existia!!, no hay bronca!!

        Return xDt

    End Function


End Module
