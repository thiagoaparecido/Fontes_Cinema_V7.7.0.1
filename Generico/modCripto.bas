Attribute VB_Name = "modCripto"
Option Explicit

Private Const sAlfa = "0123456789AaáÁàÀãÃâÂäÄBbCcçÇDdEeéÉèÈêÊëËFfGgHhIiíÍìÌîÎïÏJjKkLlMmNnOoóÓòÒõÕôÔöÖPpQqRrSsTtUuúÚùÙûÛüÜVvWwXxYyZz!@#$%&*()-_+={[}]<,>.:;?/|\"
Private Const sChave = "1QAZ2WSX3EDC4RFV5TGB6YHN7UJM8IK"

Public Function gsCripto(texto As String) As String
    Dim sResult As String
    Dim iOffSet As Integer
    Dim iPos    As Integer
    Dim i       As Integer
    Dim k       As Integer
    Dim n       As Integer

    If texto = "" Then
        sResult = ""
    Else
        n = Len(texto)
        For i = 1 To n
            If i > 31 Then
                k = i Mod 31
                If k = 0 Then
                    k = 31
                End If
            Else
                k = i
            End If

            iOffSet = InStr(sAlfa, Mid(texto, i, 1)) - 1
            If iOffSet > -1 Then
                iPos = iOffSet + Asc(Mid(sChave, k, 1))
                If iPos > 135 Then
                    iPos = iPos Mod 135
                End If
                sResult = sResult + Mid(sAlfa, iPos + 1, 1)
            Else
                sResult = sResult + Mid(texto, i, 1)
                Exit For
            End If
        Next i
    End If
    
    gsCripto = sResult
End Function

Public Function gsDecripto(texto As String) As String
    Dim sResult As String
    Dim iOffSet As Integer
    Dim iPos    As Integer
    Dim i       As Integer
    Dim k       As Integer
    Dim n       As Integer
    Dim p       As Integer

    If texto = "" Then
        sResult = ""
    Else
        n = Len(texto)
        For i = 1 To n
            If i > 31 Then
                k = i Mod 31
                If k = 0 Then
                    k = 31
                End If
            Else
                k = i
            End If

            iOffSet = InStr(sAlfa, Mid(texto, i, 1)) - 1
            If iOffSet > -1 Then
                iPos = iOffSet - Asc(Mid(sChave, k, 1))
                p = 0
                While iPos < 0
                    iPos = 135 + iPos
                Wend
                sResult = sResult + Mid(sAlfa, iPos + 1, 1)
            Else
                sResult = sResult + Mid(texto, i, 1)
                Exit For
            End If
        Next i
    End If

    gsDecripto = sResult
End Function
