Imports System.ComponentModel

Public Class Form6
    Public LlaveUsuario As String
    Public huboExito As Boolean
    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        huboExito = False
        Me.CenterToScreen()
        TextBox1.Text = ""
        TextBox2.Text = ""

    End Sub

    Private Async Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Please fill the two boxes with the password!", vbCritical, TitBox)
            Exit Sub
        End If

        If TextBox1.Text.Length < 8 Then
            MsgBox("Please make sure that the new password is at least 8 characters long!", vbCritical, TitBox)
            Exit Sub
        End If

        If TextBox1.Text <> TextBox2.Text Then
            MsgBox("The passwords doesn't match!!", vbCritical, TitBox)
            Exit Sub
        End If

        Dim cumpleUpper As Boolean = False
        Dim cumpleLower As Boolean = False
        Dim cumpleNum As Boolean = False
        Dim cumpleEspec As Boolean = False

        Dim cadUni As String = ""
        For i = 0 To TextBox1.Text.Length - 1
            cadUni = TextBox1.Text.Substring(i, 1)
            'If cadUni.
            If IsNumeric(cadUni) = True Then cumpleNum = True

            If IsNumeric(cadUni) = False Then
                If IsUpperCase(cadUni) = True Then cumpleUpper = True Else cumpleLower = True
            End If
            If ContainsSpecialChars(cadUni) = True Then cumpleEspec = True
        Next

        If cumpleNum = False Then
            MsgBox("Please make sure that your password contains the following: " & vbCrLf & "At least 8 characters long" & vbCrLf & "One Upper Case letter" & vbCrLf & "One Lower Case letter" & vbCrLf & "One number" & vbCrLf & "One special character!" & vbCrLf & "Please correct and try again!!", vbCritical, TitBox)
            Exit Sub
        End If

        If cumpleUpper = False Then
            MsgBox("Please make sure that your password contains the following: " & vbCrLf & "At least 8 characters long" & vbCrLf & "One Upper Case letter" & vbCrLf & "One Lower Case letter" & vbCrLf & "One number" & vbCrLf & "One special character!" & vbCrLf & "Please correct and try again!!", vbCritical, TitBox)
            Exit Sub
        End If

        If cumpleLower = False Then
            MsgBox("Please make sure that your password contains the following: " & vbCrLf & "At least 8 characters long" & vbCrLf & "One Upper Case letter" & vbCrLf & "One Lower Case letter" & vbCrLf & "One number" & vbCrLf & "One special character!" & vbCrLf & "Please correct and try again!!", vbCritical, TitBox)
            Exit Sub
        End If

        If cumpleEspec = False Then
            MsgBox("Please make sure that your password contains the following: " & vbCrLf & "At least 8 characters long" & vbCrLf & "One Upper Case letter" & vbCrLf & "One Lower Case letter" & vbCrLf & "One number" & vbCrLf & "One special character!" & vbCrLf & "Please correct and try again!!", vbCritical, TitBox)
            Exit Sub
        End If

        Dim elRet As String = ""
        Dim elCamino As String = ""
        Dim niuPass As String = ""
        niuPass = getSHA1Hash(TextBox1.Text)

        elCamino = RaizFire
        elCamino = elCamino & "/users/list/" & LlaveUsuario

        elRet = Await HazPutEnFbSimple(elCamino, "Pass", niuPass)
        'HazPutEnFireBaseConPath()

        If elRet <> "Ok" Then
            MsgBox(elRet, vbCritical, TitBox)
            Exit Sub
        End If

        elRet = Await HazPutEnFbSimple(elCamino, "FirstTime", " ")

        huboExito = True
        Me.Close()

    End Sub

    Private Sub Form6_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If huboExito = True Then LlaveUsuario = ""
        e.Cancel = Not huboExito
    End Sub

    Public Shared Function IsUpperCase(inputChar As Char) As Boolean
        Return inputChar >= "A"c AndAlso inputChar <= "Z"c
    End Function

    Public Function ContainsSpecialChars(s As String) As Boolean
        Return s.IndexOfAny("[~`!@#$%^&*()-+=|{}':;.,<>/?]".ToCharArray) <> -1
    End Function

End Class