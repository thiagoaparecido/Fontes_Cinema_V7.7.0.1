Attribute VB_Name = "modProtec"
Option Explicit

Const senhaProteq = "31Z16XQ03"
Const STATUS_OK = 0 '

Private Const sAlfa = "0123456789AaáÁàÀãÃâÂäÄBbCcçÇDdEeéÉèÈêÊëËFfGgHhIiíÍìÌîÎïÏJjKkLlMmNnOoóÓòÒõÕôÔöÖPpQqRrSsTtUuúÚùÙûÛüÜVvWwXxYyZz!@#$%&*()-_+={[}]<,>.:;?/|\"
Private Const sChave = "1QAZ2WSX3EDC4RFV5TGB6YHN7UJM8IK"

'DLL para Proteq
Declare Function C500 Lib "c50032.DLL" (ByVal Entra As String) As Integer

'Private clsCrypto As New cnCrypto
Private clsDb As New CineBaseDados

Public Function verificaProteq(ByRef dbConnect As ADODB.Connection, _
                               ByVal iCheck As Integer, _
                               ByRef msgErro As String) As Boolean
   Dim codRetorno  As Integer
   Dim i           As Integer
   Dim retorno     As String
   Dim parmEntrada As String
   Dim memoria     As String
   
   Dim aux1        As String
   Dim aux2        As String
   Dim dtLimite    As Date
   Dim dtUltAcess  As Date
   Dim dtAgora     As Date
   
   verificaProteq = False
   
   dtAgora = Now
   
   'Obtem inicialização
   parmEntrada = String$(10, Chr$(0))
   parmEntrada = Chr$(3) & senhaProteq
   codRetorno = C500(parmEntrada)
   
   'se ocorreu algum problema, sai da aplicação
   If codRetorno <> STATUS_OK Then
      retorno = Trim$(Format$(codRetorno, "########0"))
      msgErro = "Impossível continuar; Problemas com a Chave (" & retorno & ") !"
      
      Exit Function
   End If
   
   'Le dados na memória
   parmEntrada = String$(10, Chr$(0))
   memoria = String$(96, Chr$(0))
   
   For i = 0 To 47
      parmEntrada = Chr$(1) & "xx" & Chr$(i)
      codRetorno = C500(parmEntrada)
      
      If codRetorno <> STATUS_OK Then
         retorno = Trim$(Format$(codRetorno, "########0"))
         msgErro = "Impossível continuar; Problemas com a Chave (" & retorno & ") !"
         
         'Finaliza Proteq
         parmEntrada = String$(10, Chr$(0))
         parmEntrada = Chr$(5)
         codRetorno = C500(parmEntrada)
         Exit Function
      End If
      
      Mid$(memoria, 2 * i + 1, 2) = Mid$(parmEntrada, 2, 2)
   Next i
   
   'Le data limite
   aux1 = ""
   For i = 60 To 63
      parmEntrada = Chr$(1) & "xx" & Chr$(i)
      codRetorno = C500(parmEntrada)
   
      If codRetorno <> STATUS_OK Then
         retorno = Trim$(Format$(codRetorno, "########0"))
         msgErro = "Impossível continuar; Problemas com a Chave (" & retorno & ") !"
         
         'Finaliza Proteq
         parmEntrada = String$(10, Chr$(0))
         parmEntrada = Chr$(5)
         codRetorno = C500(parmEntrada)
         Exit Function
      End If
      
      aux1 = aux1 & Mid$(parmEntrada, 2, 2)
   Next i
   
   'Le data do ultimo acesso
   aux2 = ""
   For i = 64 To 70
      parmEntrada = Chr$(1) & "xx" & Chr$(i)
      codRetorno = C500(parmEntrada)
      
      If codRetorno <> STATUS_OK Then
         retorno = Trim$(Format$(codRetorno, "########0"))
         msgErro = "Impossível continuar; Problemas com a Chave (" & retorno & ") !"
          
         'Finaliza Proteq
         parmEntrada = String$(10, Chr$(0))
         parmEntrada = Chr$(5)
         codRetorno = C500(parmEntrada)
         Exit Function
      End If
      
      aux2 = aux2 & Mid$(parmEntrada, 2, 2)
   Next i
   
   If Len(Trim(aux1)) = 8 Then
      dtLimite = DateSerial(CInt(Mid(aux1, 5, 4)), CInt(Mid(aux1, 3, 2)), CInt(Mid(aux1, 1, 2)))
   ElseIf Len(Trim(aux1)) <> 0 Then
      msgErro = "Impossível continuar; Problemas com a Chave! Código de erro 100"
      
      If iCheck <> 77 Then
         dtUltAcess = DateSerial(1900, 1, 1)
         If Not CorrigeErros(dtAgora, dtUltAcess, msgErro) Then
            'Finaliza Proteq
            parmEntrada = String$(10, Chr$(0))
            parmEntrada = Chr$(5)
            codRetorno = C500(parmEntrada)
            Exit Function
          End If
      Else
         'Finaliza Proteq
         parmEntrada = String$(10, Chr$(0))
         parmEntrada = Chr$(5)
         codRetorno = C500(parmEntrada)
         Exit Function
      End If
   End If
   
   If Len(Trim(aux2)) = 14 Then
      dtUltAcess = DateSerial(CInt(Mid(aux2, 5, 4)), CInt(Mid(aux2, 3, 2)), CInt(Mid(aux2, 1, 2))) + _
                   TimeSerial(CInt(Mid(aux2, 9, 2)), CInt(Mid(aux2, 11, 2)), CInt(Mid(aux2, 13, 2)))
   End If
   
   If Len(Trim(aux1)) = 8 Then
      If Len(Trim(aux2)) = 14 Then
         If dtAgora < dtUltAcess Then
            msgErro = "Impossível continuar; Problemas com a Chave! Código de erro 300"
            
            If iCheck <> 77 Then
               If Not CorrigeErros(dtAgora, dtUltAcess, msgErro) Then
                  'Finaliza Proteq
                  parmEntrada = String$(10, Chr$(0))
                  parmEntrada = Chr$(5)
                  codRetorno = C500(parmEntrada)
                  Exit Function
               Else
                  aux2 = Format(dtUltAcess, "ddmmyyyyHhNnSs")
               End If
            Else
               'Finaliza Proteq
               parmEntrada = String$(10, Chr$(0))
               parmEntrada = Chr$(5)
               codRetorno = C500(parmEntrada)
               Exit Function
            End If
         End If
      End If
   
      If dtAgora > dtLimite Then
         If Not liberaNovoPer() Then
            msgErro = "Impossível continuar; Problemas com a Chave! Código de erro 200"
            
            'Finaliza Proteq
            parmEntrada = String$(10, Chr$(0))
            parmEntrada = Chr$(5)
            codRetorno = C500(parmEntrada)
            Exit Function
        End If
      End If
   End If
   
   If iCheck <> 77 Then
      If Not clsDb.VerificaUltAcessoBanco(dbConnect, aux2, msgErro) Then
         If msgErro <> "" Then
            msgErro = "Impossível continuar; Problemas com a Chave! Código de erro 400"
         
            If Not CorrigeErros(dtAgora, dtUltAcess, msgErro) Then
               'Finaliza Proteq
               parmEntrada = String$(10, Chr$(0))
               parmEntrada = Chr$(5)
               codRetorno = C500(parmEntrada)
               Exit Function
            Else
               aux2 = Format(dtUltAcess, "ddmmyyyyHhNnSs")
            End If
         Else
            'Finaliza Proteq
            parmEntrada = String$(10, Chr$(0))
            parmEntrada = Chr$(5)
            codRetorno = C500(parmEntrada)
            Exit Function
         End If
      End If
       
      'Grava data/hora do ultimo acesso
      aux2 = Format(dtAgora, "ddmmyyyyHhNnSs")
      If Not clsDb.atuUltAcessoBanco(dbConnect, aux2, msgErro) Then
         If msgErro <> "" Then
            msgErro = "Impossível continuar; Problemas com a Chave! Código de erro 400"
         
            If Not CorrigeErros(dtAgora, dtUltAcess, msgErro) Then
               'Finaliza Proteq
               parmEntrada = String$(10, Chr$(0))
               parmEntrada = Chr$(5)
               codRetorno = C500(parmEntrada)
               Exit Function
            Else
               aux2 = Format(dtUltAcess, "ddmmyyyyHhNnSs")
            End If
         Else
            'Finaliza Proteq
            parmEntrada = String$(10, Chr$(0))
            parmEntrada = Chr$(5)
            codRetorno = C500(parmEntrada)
            Exit Function
         End If
      End If
   
      For i = 64 To 70
         parmEntrada = Chr$(2) & Mid(aux2, 1 + (i - 64) * 2, 2) & Chr$(i)
         codRetorno = C500(parmEntrada)
         
         If codRetorno <> STATUS_OK Then
            retorno = Trim$(Format$(codRetorno, "########0"))
            msgErro = "Impossível continuar; Problemas com a Chave (" & retorno & ") !"
            
            'Finaliza Proteq
            parmEntrada = String$(10, Chr$(0))
            parmEntrada = Chr$(5)
            codRetorno = C500(parmEntrada)
            Exit Function
         End If
      Next i
   End If
   
   'Finaliza Proteq
   parmEntrada = String$(10, Chr$(0))
   parmEntrada = Chr$(5)
   codRetorno = C500(parmEntrada)
   
   'se ocorreu algum problema, sai da aplicação
   If codRetorno <> STATUS_OK Then
      retorno = Trim$(Format$(codRetorno, "########0"))
      msgErro = "Impossível continuar; Problemas com a Chave (" & retorno & ") !"
      Exit Function
   End If
   
   verificaProteq = True
End Function

Private Function liberaNovoPer() As Boolean
    Dim chave       As String
    Dim dia         As String
    Dim mes         As String
    Dim ano         As String
    Dim dias        As String
    Dim dtRef       As Date
    Dim codRetorno  As Integer
    Dim i           As Integer
    Dim retorno     As String
    Dim parmEntrada As String
    Dim aux2        As String
    
    On Error GoTo trataErro
    
    liberaNovoPer = False
    
    chave = InputBox("Entre com a chave de ativação", App.ProductName)
    
    If Len(chave) <> 0 Then
       chave = gsDecripto(chave)
       
       dia = Mid(chave, 7, 2)
       mes = Mid(chave, 10, 2)
       ano = Mid(chave, 1, 2) & Mid(chave, 4, 2)
       dias = Mid(chave, 3, 1) & Mid(chave, 6, 1) & Mid(chave, 9, 1)
       
       If (Not IsNumeric(dia)) Or _
          (Not IsNumeric(mes)) Or _
          (Not IsNumeric(ano)) Or _
          (Not IsNumeric(dias)) Then
          
          Exit Function
       End If
       
       dtRef = DateSerial(CInt(ano), CInt(mes), CInt(dia))
       
       If dtRef <> Date Then
          Exit Function
       End If
       
       dtRef = DateAdd("d", CInt(dias), dtRef)
       
       aux2 = Format(dtRef, "DDMMYYYY")
       
       For i = 60 To 63
          parmEntrada = Chr$(2) & Mid(aux2, 1 + (i - 60) * 2, 2) & Chr$(i)
          codRetorno = C500(parmEntrada)
         
          If codRetorno <> STATUS_OK Then
             'Finaliza Proteq
             parmEntrada = String$(10, Chr$(0))
             parmEntrada = Chr$(5)
             codRetorno = C500(parmEntrada)
             
             Exit Function
          End If
       Next i
    Else
        Exit Function
    End If
    
    liberaNovoPer = True
    
    Exit Function
trataErro:

End Function

Private Function CorrigeErros(ByVal dtAgora As Date, ByRef dtHardLock As Date, msgErro As String) As Boolean
    Dim chave       As String
    Dim dia         As String
    Dim mes         As String
    Dim ano         As String
    Dim dias        As String
    Dim dtRef       As Date
    Dim dtBanco     As Date
    Dim msgEr       As String
    Dim strAux      As String
    Dim aux2        As String
    Dim codRetorno  As Integer
    Dim i           As Integer
    Dim retorno     As String
    Dim parmEntrada As String
    
    CorrigeErros = False
    
    dtBanco = clsDb.UltAcessoBanco(dbConnect, msgEr)
    
    strAux = msgErro & Chr(10) & Chr(10) & Chr(10)
    strAux = strAux & "Data da Maquina: " & Format(dtAgora, "dd/mm/yyyy Hh:Nn:Ss") & Chr(10)
    strAux = strAux & "Data da Chave: " & Format(dtHardLock, "dd/mm/yyyy Hh:Nn:Ss") & Chr(10)
    strAux = strAux & "Data do Banco: " & Format(dtBanco, "dd/mm/yyyy Hh:Nn:Ss") & Chr(10) & Chr(10)
    strAux = strAux & "Entre com a chave de ativação"
    
    chave = InputBox(strAux, App.ProductName)
    
    If Len(chave) <> 0 Then
       chave = gsDecripto(chave)
       
       dia = Mid(chave, 7, 2)
       mes = Mid(chave, 10, 2)
       ano = Mid(chave, 1, 2) & Mid(chave, 4, 2)
       dias = Mid(chave, 3, 1) & Mid(chave, 6, 1) & Mid(chave, 9, 1)
       
       If (Not IsNumeric(dia)) Or _
          (Not IsNumeric(mes)) Or _
          (Not IsNumeric(ano)) Or _
          (Not IsNumeric(dias)) Then
          
          MsgBox "Problemas na chave de ativação!", vbCritical, App.ProductName
          
          Exit Function
       End If
       
       dtRef = DateSerial(CInt(ano), CInt(mes), CInt(dia))
       
       If dtRef <> Date Then
          MsgBox "Problemas na chave de ativação!", vbCritical, App.ProductName
          
          Exit Function
       End If
       
       aux2 = Format(dtAgora, "ddmmyyyyHhNnSs")
       If Not clsDb.atuUltAcessoBanco(dbConnect, aux2, msgErro) Then
           MsgBox "Problemas na atualização do banco!", vbCritical, App.ProductName
          
           Exit Function
       End If
       
       For i = 64 To 70
          parmEntrada = Chr$(2) & Mid(aux2, 1 + (i - 64) * 2, 2) & Chr$(i)
          codRetorno = C500(parmEntrada)
         
          If codRetorno <> STATUS_OK Then
             retorno = Trim$(Format$(codRetorno, "########0"))
             msgErro = "Impossível continuar; Problemas com Proteq (" & retorno & ") !"
            
             'Finaliza Proteq
             parmEntrada = String$(10, Chr$(0))
             parmEntrada = Chr$(5)
             codRetorno = C500(parmEntrada)
             Exit Function
          End If
       Next i
       
       dtHardLock = dtAgora
       
       CorrigeErros = True
    End If
    
    Exit Function
trataErro:

End Function


Private Function gsCripto(texto As String) As String
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

Private Function gsDecripto(texto As String) As String
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



