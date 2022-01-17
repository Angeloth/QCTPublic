Imports System.Net.Mail
Module Module4

    'modulo para procesar las validaciones

    Public Function isLengthOk(ByVal elTexto As String, ByVal elLargo As Integer, ByRef elError As String) As Boolean

        Dim estaOk As Boolean = False

        If elTexto.Length <= elLargo Then
            estaOk = True
        Else
            elError = "The maximum character limit for this field is " & elLargo & ", currently is exceeded by " & elTexto.Length - elLargo & " character(s)!, please review."
        End If

        Return estaOk

    End Function

    Public Function isIntegerNumberOk(ByVal elValor As String, ByRef elError As String) As Boolean

        elError = ""
        Dim estaOk As Boolean = False
        Dim unEntero1 As Long ' = False

        If IsNumeric(elValor) = False Then
            elError = "Please enter a valid integer number!"
            Return False
        End If

        If elValor.Contains(".") = True Then
            elError = "Please enter a valid integer number!, no decimals or fractional parts allowed!!"
            Return False
        End If

        If Long.TryParse(elValor, unEntero1) = True Then
            '
            estaOk = True
        Else
            elError = "Please enter a valid integer number!"
        End If

        Return estaOk

    End Function

    Public Function isDecimalOk(ByVal elValor As String, ByVal deciFormat As Decimal, ByRef elError As String) As Boolean

        If IsNumeric(elValor) = False Then
            elError = "This must be a decimal field!, please review"
            Return False
        End If

        If elValor.Contains(".") = False Then
            elError = "This must be a decimal number it should have a dot '.' inside"
            Return False
        End If

        Dim yObj As Object = Split(elValor, ".")

        Dim estaOk As Boolean = False
        Dim xDeci As Decimal = CDec(elValor)
        Dim entePart As Decimal = 0
        Dim fracPart As Decimal = 0
        Dim miDeci As String = CStr(deciFormat)
        Dim xObj As Object = Split(miDeci, ".")

        Call SplitDecimal(xDeci, entePart, fracPart)

        If CStr(entePart).Length <= CInt(xObj(0)) And CStr(yObj(1)).Length <= CInt(xObj(1)) Then
            estaOk = True
        Else
            elError = "This is a decimal field, it must have " & CStr(xObj(0)) & " maximum digits for integer part and " & CStr(xObj(1)) & " for fractional part, please review!"
        End If

        'si el numero 
        Return estaOk

    End Function

    Public Function isEmailValid(ByVal elMail As String, ByRef elError As String) As Boolean

        Try
            Dim a As New System.Net.Mail.MailAddress(elMail)
        Catch
            elError = "This field must be an email and this is not a valid email address, please review!"
            Return False
        End Try

        Return True

    End Function

    Public Function isDate8FormatOk(ByVal elValor As String, ByRef elError As String) As Boolean

        Dim estaOk As Boolean = False
        Dim partAnio As String
        Dim partMes As String
        Dim partDay As String
        If IsNumeric(elValor) = True Then
            partAnio = Left(elValor, 4)
            partMes = Mid(elValor, 5, 2)
            partDay = Right(elValor, 2)

            If IsDate(partAnio & "/" & partMes & "/" & partDay) = True Then
                estaOk = True
            Else
                elError = "This is an invalid date, please review it!!"
            End If
        Else
            elError = "This is a Date field, and is in 8 character format, please fill it in the format: YYYYMMDD, only numbers allowed!"
        End If

        Return estaOk

    End Function

    Public Function isFieldinDateFormat(ByVal elFecha As String, ByRef FechaCorregida As String, ByRef elError As String) As Boolean

        Dim estaOk As Boolean = False
        Dim xObj As Object
        Dim yObj As Object
        Dim hayError As Boolean
        Dim i As Integer

        yObj = Nothing
        xObj = Nothing
        xObj = Split(elFecha, "/")

        hayError = True
        If UBound(xObj) = 2 Then
            yObj = xObj
            hayError = False
        Else
            xObj = Nothing
            xObj = Split(elFecha, "-")
            If UBound(xObj) = 2 Then
                yObj = xObj
                hayError = False
            Else
                xObj = Nothing
                xObj = Split(elFecha, ".")
                If UBound(xObj) = 2 Then
                    yObj = xObj
                    hayError = False
                Else
                    'error!!
                    elError = "Please enter a valid Date format: DD/MM/YYY, or DD-MM-YYYY, or DD.MM.YYYY"
                End If
            End If
        End If

        If hayError = True Then
            estaOk = False
            'CampoTieneFormatoDeFecha = False
        Else

            Dim dia As Integer
            Dim mes As Integer
            Dim diSt As String = ""
            Dim meSt As String = ""

            If Len(xObj(0)) = 4 Then
                'izquierda con el cero

                If IsDate(CStr(xObj(2)) & "/" & CStr(xObj(1)) & "/" & CStr(xObj(0))) = True Then
                    dia = CInt(xObj(2))
                    mes = CInt(xObj(1))

                    '2021-12-31
                    '2021-1-1
                    If mes > 12 Then
                        'entonces el mes es el dia
                        dia = CInt(xObj(1))
                        mes = CInt(xObj(2))
                    End If

                    'el mes es el mes
                    If dia < 10 Then
                        diSt = "0"
                    End If
                    diSt = diSt & CStr(dia)

                    If mes < 10 Then
                        meSt = "0"
                    End If
                    meSt = meSt & CStr(mes)

                    'todo ok!
                    'CampoTieneFormatoDeFecha = True
                    estaOk = True
                    FechaCorregida = CStr(diSt) & "." & CStr(meSt) & "." & CStr(xObj(0))

                Else
                    estaOk = False
                    'CampoTieneFormatoDeFecha = False
                    elError = "This is not a date, please enter a valid Date format: DD/MM/YYY, or DD-MM-YYYY, or DD.MM.YYYY"
                End If

            Else

                If Len(xObj(2)) = 4 Then

                    If IsDate(CStr(xObj(0)) & "/" & CStr(xObj(1)) & "/" & CStr(xObj(2))) = True Then

                        dia = CInt(xObj(0))
                        mes = CInt(xObj(1))

                        '2021-12-31
                        '2021-1-1
                        If mes > 12 Then
                            'entonces el mes es el dia
                            dia = CInt(xObj(1))
                            mes = CInt(xObj(0))

                        End If

                        'el mes es el mes
                        If dia < 10 Then
                            diSt = "0"
                        End If
                        diSt = diSt & CStr(dia)

                        If mes < 10 Then
                            meSt = "0"
                        End If
                        meSt = meSt & CStr(mes)

                        estaOk = True
                        'CampoTieneFormatoDeFecha = True
                        FechaCorregida = CStr(diSt) & "." & CStr(meSt) & "." & CStr(xObj(2))

                    Else
                        estaOk = False
                        'CampoTieneFormatoDeFecha = False
                        elError = "This is not a date, please enter a valid Date format: DD/MM/YYY, or DD-MM-YYYY, or DD.MM.YYYY"
                    End If

                Else
                    estaOk = False
                    'CampoTieneFormatoDeFecha = False
                    elError = "This is not a date, please enter a valid Date format: DD/MM/YYY, or DD-MM-YYYY, or DD.MM.YYYY"
                End If
            End If

        End If

        Return estaOk

    End Function

    Public Function isIndicatorFieldOk(ByVal elValor As String, ByRef elError As String) As Boolean

        If elValor.Length <> 1 Then
            elError = "This is an indicator field, please type either X or a white space"
            Return False
        End If

        If elValor = "X" Or elValor = "" Then '
            Return True
        Else
            elError = "This is an indicator field, please type either X or a white space"
            Return False
        End If

    End Function

    Public Function isFieldinTimeFormat(ByVal elValor As String, ByVal elAnchoChar As Integer, ByRef elError As String, ByRef TimeCorregido As String) As Boolean

        Dim i As Long
        Dim j As Long
        Dim partHr As String = ""
        Dim partMin As String = ""
        Dim partSec As String = ""
        Dim estaOk As Boolean = False
        'Dim miCad As String
        elValor = Format(elValor, "HH:mm:ss")

        If Len(elValor) = 6 Or Len(elValor) = 8 Then
            Select Case elAnchoChar
                Case Is = 6
                    partHr = Left(elValor, 2)
                    partMin = Mid(elValor, 3, 2)
                    partSec = Mid(elValor, 5, 2)

                Case Is = 8
                    partHr = Left(elValor, 2)
                    partMin = Mid(elValor, 4, 2)
                    partSec = Right(elValor, 2)
            End Select

            If IsNumeric(partHr) = True Then
                If IsNumeric(partMin) = True Then
                    If IsNumeric(partSec) = True Then
                        If CInt(partHr) >= 0 And CInt(partHr) <= 23 Then
                            If CInt(partMin) >= 0 And CInt(partMin) <= 59 Then
                                If CInt(partSec) >= 0 And CInt(partSec) <= 59 Then
                                    'CampoTimeValido = True
                                    estaOk = True
                                Else
                                    elError = "Make sure the seconds part is between 00 and 59 seconds"
                                End If
                            Else
                                elError = "Make sure the minute part is between 00 and 59  minutes"
                            End If
                        Else
                            elError = "Make sure the hour part is between 00 and 23 hours"
                        End If
                    Else
                        elError = "This is a time field, please fill it in the format: HH:MM:SS for 8 characters field, and HHMMSS for 6 character long!"
                    End If
                Else
                    elError = "This is a time field, please fill it in the format: HH:MM:SS for 8 characters field, and HHMMSS for 6 character long!"
                End If
            Else
                elError = "This is a time field, please fill it in the format: HH:MM:SS for 8 characters field, and HHMMSS for 6 character long!"
            End If

            If elAnchoChar = 6 Then
                TimeCorregido = partHr & partMin & partSec
            Else
                TimeCorregido = partHr & ":" & partMin & ":" & partSec
            End If

        Else
            elError = "The length of this field must be 6 or 8, please review!"
        End If

        Return estaOk

    End Function

    Public Function ContainsSpecialCharacters(str As String) As Boolean
        Dim ch As String = ""
        Dim Contiene As Boolean = False
        For i = 1 To Len(str)
            ch = Mid(str, i, 1)
            Select Case ch
                Case "0" To "9", "A" To "Z", "a" To "z", " "
                    Contiene = False
                Case Else
                    Contiene = True
                    Exit For
            End Select
        Next

        Return Contiene

    End Function

    Public Function ContainsInvalidChars(ByVal elTexto As String, ByVal losInvalids As String, ByRef elError As String) As Boolean

        Dim conTiene As Boolean = False
        Dim unCh As String = ""
        elError = ""
        For i = 0 To losInvalids.Length - 1
            unCh = Mid(losInvalids, i, 1)

            If elTexto.Contains(unCh) = True Then
                elError = "This field is not allowed to contain any of the following characters: " & losInvalids & " , please review!"
                conTiene = True
                Exit For
            End If

        Next

        Return conTiene

    End Function
    Public Function isFieldContainedOnTable(ByVal laDt As DataTable, ByVal elCampoName As String, ByVal elValor As String, ByRef elError As String) As Boolean
        'debe haber otra funcion que jale el dataset y los catálogos!
        'con el buscador Select podemos encontrar de la tabla fulana el campo fulano que contenga un row con esa
        'condicion!
        Dim siHay As Boolean = False

        Dim filterDT As DataTable = laDt.Clone()

        Dim result() As DataRow = laDt.Select(elCampoName & " = '" & elValor & "'")

        For Each row As DataRow In result
            filterDT.ImportRow(row)
        Next

        If filterDT.Rows.Count > 0 Then
            siHay = True
        Else
            elError = "The value " & elValor & " is not contained on the specified table!"
        End If

        Return siHay

    End Function

    Sub SplitDecimal(ByVal number As Decimal, ByRef wholePart As Decimal, ByRef fractionalPart As Decimal)
        wholePart = Math.Truncate(number)
        fractionalPart = number - wholePart
    End Sub


End Module
