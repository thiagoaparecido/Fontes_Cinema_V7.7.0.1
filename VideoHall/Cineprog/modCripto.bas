Attribute VB_Name = "modCripto"
Option Explicit

Private Const sAlfa = "0123456789Aa����������BbCc��DdEe��������FfGgHhIi��������JjKkLlMmNnOo����������PpQqRrSsTtUu��������VvWwXxYyZz!@#$%&*()-_+={[}]<,>.:;?/|\"
Private Const sChave = "1QAZ2WSX3EDC4RFV5TGB6YHN7UJM8IK"

Public Function gsCripto(texto As String) As String
    Dim sResult As String
    Dim iOffSet As Integer
    Dim iPos    As Integer
    Dim I       As Integer
    Dim k       As Integer
    Dim n       As Integer

    If texto = "" Then
        sResult = ""
    Else
        n = Len(texto)
        For I = 1 To n
            If I > 31 Then
                k = I Mod 31
                If k = 0 Then
                    k = 31
                End If
            Else
                k = I
            End If

            iOffSet = InStr(sAlfa, Mid(texto, I, 1)) - 1
            If iOffSet > -1 Then
                iPos = iOffSet + Asc(Mid(sChave, k, 1))
                If iPos > 135 Then
                    iPos = iPos Mod 135
                End If
                sResult = sResult + Mid(sAlfa, iPos + 1, 1)
            Else
                sResult = sResult + Mid(texto, I, 1)
                Exit For
            End If
        Next I
    End If
    
    gsCripto = sResult
End Function

Public Function gsDecripto(texto As String) As String
    Dim sResult As String
    Dim iOffSet As Integer
    Dim iPos    As Integer
    Dim I       As Integer
    Dim k       As Integer
    Dim n       As Integer
    Dim p       As Integer

    If texto = "" Then
        sResult = ""
    Else
        n = Len(texto)
        For I = 1 To n
            If I > 31 Then
                k = I Mod 31
                If k = 0 Then
                    k = 31
                End If
            Else
                k = I
            End If

            iOffSet = InStr(sAlfa, Mid(texto, I, 1)) - 1
            If iOffSet > -1 Then
                iPos = iOffSet - Asc(Mid(sChave, k, 1))
                p = 0
                While iPos < 0
                    iPos = 135 + iPos
                Wend
                sResult = sResult + Mid(sAlfa, iPos + 1, 1)
            Else
                sResult = sResult + Mid(texto, I, 1)
                Exit For
            End If
        Next I
    End If

    gsDecripto = sResult
End Function


