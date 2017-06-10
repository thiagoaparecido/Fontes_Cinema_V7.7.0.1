Attribute VB_Name = "modConfCineProg"
Option Explicit
Global Const senhaProteq = "31Z16XQ03"
Global Const senhaBanco = "31Z16XQ03"

'DLL para Proteq
Declare Function C500 Lib "c50032.DLL" (ByVal Entra As String) As Integer
Global Const STATUS_OK = 0 '* API call was succesfull        *

Public gsBandoDados  As String
Public gsArqsVideo   As String
Public gbCheck       As Integer

Public dtReferencia1 As Date
Public dtReferencia2 As Date
Public dtStrRef1     As String
Public dtStrRef2     As String

Public clsPC        As New clsPainelControle

Public Sub leRgistro()
    'Caminho do banco de dados
    gsBandoDados = GetSetting("CineProg", "Diretorios", "bancoDados", "")
    If gsBandoDados = "" Then
        gsBandoDados = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
        SaveSetting "CineProg", "Diretorios", "bancoDados", gsBandoDados
    End If

    'Caminho dos arquivos de videos
    gsArqsVideo = GetSetting("CineProg", "Diretorios", "arqsVideo", "")
    If gsArqsVideo = "" Then
        gsArqsVideo = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
        SaveSetting "CineProg", "Diretorios", "arqsVideo", gsArqsVideo
    End If

   If Not IsNumeric(GetSetting("CineProg", "Diretorios", "CHECK", "")) Then
      gbCheck = 1
      SaveSetting "CineProg", "Diretorios", "CHECK", "1"
   Else
      gbCheck = CInt(GetSetting("CineProg", "Diretorios", "CHECK", ""))
   End If
End Sub

'Sub Main()
'    Call leRgistro
'
'    If Not gbAbreBase() Then
'        End
'    End If
'
'    frmConfig.Show
'End Sub

Public Sub gInicializaComboFilme(ByRef cboFilme As ComboBox, ByRef todos As Boolean)
   Dim sMsg    As String
   Dim strSql  As String
   Dim strSql1 As String
   Dim i       As Integer
   Dim rsFilme As New ADODB.Recordset
   
   On Error GoTo TrataErro
   
   dtReferencia1 = CDate("01/01/1900")
   dtReferencia2 = DateAdd("d", 1, dtReferencia1)
   
   dtStrRef1 = Format(dtReferencia1, "Short Date")
   dtStrRef2 = Format(dtReferencia2, "Short Date")
   
   
   If todos Then
     strSql = "SELECT codFilme, "
     strSql = strSql & "descricao "
     strSql = strSql & "FROM tb_filmes "
     strSql = strSql & "ORDER BY descricao"
   
     strSql1 = "SELECT codFilme "
     strSql1 = strSql1 & "FROM tb_filmes "
     strSql1 = strSql1 & "WHERE codFilme = 0 "
   
     rsFilme.Open strSql1, gConnect, adOpenDynamic, adLockReadOnly, adCmdText
      
     If rsFilme.EOF And rsFilme.BOF Then
        strSql1 = "INSERT INTO tb_filmes(codFilme,  descricao, duracao) "
        strSql1 = strSql1 & "VALUES (0, 'Filmes sem preço especifico', 0) "
        gConnect.Execute strSql1
     End If
     
     rsFilme.Close
   Else
     strSql = "SELECT codFilme, "
     strSql = strSql & "descricao "
     strSql = strSql & "FROM tb_filmes "
     strSql = strSql & "WHERE codFilme <> 0 "
     strSql = strSql & "ORDER BY descricao"
   End If
   
   rsFilme.Open strSql, gConnect, adOpenDynamic, adLockReadOnly, adCmdText
   
   cboFilme.Clear
       
   If Not (rsFilme.EOF And rsFilme.BOF) Then
       Do While Not rsFilme.EOF
           cboFilme.AddItem rsFilme.Fields("descricao").Value
           cboFilme.ItemData(cboFilme.NewIndex) = rsFilme.Fields("codFilme").Value
           
           rsFilme.MoveNext
       Loop
   End If
   rsFilme.Close
   
   Exit Sub
   
TrataErro:
   If rsFilme.State = adStateOpen Then
       rsFilme.Close
   End If
   
   sMsg = "Ocorreu um erro em gInicializaComboFilme." & vbCrLf
   sMsg = sMsg & Err.Number & " - " & Err.Description
   MsgBox sMsg, vbCritical, App.ProductName
End Sub

Public Sub gInicializaComboSala(ByRef cboSala As ComboBox)
    Dim sMsg    As String
    Dim strSql  As String
    Dim i       As Integer
    Dim rsSala As New ADODB.Recordset
    
    On Error GoTo TrataErro
    
    strSql = "SELECT codSala, "
    strSql = strSql & "descricao "
    strSql = strSql & "FROM tb_salas "
    strSql = strSql & "ORDER BY descricao"
    
    rsSala.Open strSql, gConnect, adOpenDynamic, adLockReadOnly, adCmdText
    
    cboSala.Clear
        
    If Not (rsSala.EOF And rsSala.BOF) Then
        Do While Not rsSala.EOF
            cboSala.AddItem rsSala.Fields("descricao").Value
            cboSala.ItemData(cboSala.NewIndex) = rsSala.Fields("codSala").Value
            
            rsSala.MoveNext
        Loop
    End If
    rsSala.Close
    
    Exit Sub
    
TrataErro:
    If rsSala.State = adStateOpen Then
        rsSala.Close
    End If
    
    sMsg = "Ocorreu um erro em gInicializaComboSala." & vbCrLf
    sMsg = sMsg & Err.Number & " - " & Err.Description
    MsgBox sMsg, vbCritical, App.ProductName
End Sub

Public Sub gSetaComboFilme(ByRef cboFilme As ComboBox, ByVal codFilme As Integer)
    Dim i As Integer
        
    cboFilme.ListIndex = -1
    
    For i = 0 To cboFilme.ListCount - 1
        If cboFilme.ItemData(i) = codFilme Then
            cboFilme.ListIndex = i
            Exit For
        End If
    Next i
End Sub

Public Sub gSetaComboSala(ByRef cboSala As ComboBox, ByVal codSala As Integer)
    Dim i As Integer
        
    cboSala.ListIndex = -1
    
    For i = 0 To cboSala.ListCount - 1
        If cboSala.ItemData(i) = codSala Then
            cboSala.ListIndex = i
            Exit For
        End If
    Next i
End Sub

Private Function chkUsuProteq() As Boolean
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
    
    chkUsuProteq = False

    'Obtem inicialização
    parmEntrada = String$(10, Chr$(0))
    parmEntrada = Chr$(3) & senhaProteq
    codRetorno = C500(parmEntrada)

    'se ocorreu algum problema, sai da aplicação
    If codRetorno <> STATUS_OK Then
       retorno = Trim$(Format$(codRetorno, "########0"))
       MsgBox "Impossível continuar; Problemas com Proteq (" & retorno & ") !", vbOKOnly + vbCritical + vbSystemModal, App.ProductName
       
       'Finaliza Proteq
       parmEntrada = String$(10, Chr$(0))
       parmEntrada = Chr$(5)
       codRetorno = C500(parmEntrada)
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
           MsgBox "Impossível continuar; Problemas com Proteq (" & retorno & ") !", vbOKOnly + vbCritical + vbSystemModal, App.ProductName
           
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
           MsgBox "Impossível continuar; Problemas com Proteq (" & retorno & ") !", vbOKOnly + vbCritical + vbSystemModal, App.ProductName
           
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
           MsgBox "Impossível continuar; Problemas com Proteq (" & retorno & ") !", vbOKOnly + vbCritical + vbSystemModal, App.ProductName
            
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
        MsgBox "Impossível continuar; Problemas com hardLock! Código de erro 100", vbOKOnly + vbCritical + vbSystemModal, App.ProductName
        
        'Finaliza Proteq
        parmEntrada = String$(10, Chr$(0))
        parmEntrada = Chr$(5)
        codRetorno = C500(parmEntrada)
        Exit Function
    End If
    
    If Len(Trim(aux2)) = 14 Then
        dtUltAcess = DateSerial(CInt(Mid(aux2, 5, 4)), CInt(Mid(aux2, 3, 2)), CInt(Mid(aux2, 1, 2))) + _
                     TimeSerial(CInt(Mid(aux2, 9, 2)), CInt(Mid(aux2, 11, 2)), CInt(Mid(aux2, 13, 2)))
    End If
    
    dtAgora = Now
    
    If Len(Trim(aux1)) = 8 Then
        If dtAgora > dtLimite Then
            MsgBox "Impossível continuar; Problemas com hardLock! Código de erro 200", vbOKOnly + vbCritical + vbSystemModal, App.ProductName
            
            'Finaliza Proteq
            parmEntrada = String$(10, Chr$(0))
            parmEntrada = Chr$(5)
            codRetorno = C500(parmEntrada)
            Exit Function
        End If
    
        If Len(Trim(aux2)) = 14 Then
            If dtAgora < dtUltAcess Then
                MsgBox "Impossível continuar; Problemas com hardLock! Código de erro 300", vbOKOnly + vbCritical + vbSystemModal, App.ProductName
                'Finaliza Proteq
                parmEntrada = String$(10, Chr$(0))
                parmEntrada = Chr$(5)
                codRetorno = C500(parmEntrada)
                Exit Function
            End If
        End If
    End If
    
    If gbCheck <> 1 Then
        If Not verificaDtUltAcessBanco(aux2) Then
            MsgBox "Impossível continuar; Problemas com hardLock! Código de erro 400", vbOKOnly + vbCritical + vbSystemModal, App.ProductName
            
            'Finaliza Proteq
            parmEntrada = String$(10, Chr$(0))
            parmEntrada = Chr$(5)
            codRetorno = C500(parmEntrada)
            Exit Function
        End If
        
        'Grava data/hora do ultimo acesso
        aux2 = Format(dtAgora, "ddmmyyyyHhNnSs")
        If Not atuDtUltAcessBanco(aux2) Then
            MsgBox "Impossível continuar; Problemas com hardLock! Código de erro 500", vbOKOnly + vbCritical + vbSystemModal, App.ProductName
            
            'Finaliza Proteq
            parmEntrada = String$(10, Chr$(0))
            parmEntrada = Chr$(5)
            codRetorno = C500(parmEntrada)
            Exit Function
        End If
    
        For i = 64 To 70
            parmEntrada = Chr$(2) & Mid(aux2, 1 + (i - 64) * 2, 2) & Chr$(i)
            codRetorno = C500(parmEntrada)
        
            If codRetorno <> STATUS_OK Then
               retorno = Trim$(Format$(codRetorno, "########0"))
               MsgBox "Impossível continuar; Problemas com Proteq (" & retorno & ") !", vbOKOnly + vbCritical + vbSystemModal, App.ProductName
               
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
       MsgBox "Impossível continuar; Problemas com Proteq (" & retorno & ") !", vbOKOnly + vbCritical + vbSystemModal, App.ProductName
       Exit Function
    End If

    'Retorno da função
    'chkUsuProteq = Trim$(Mid$(memoria, 55, 26)) & "  NS# " & Trim$(Mid$(memoria, 81, 6)) & " (" & Trim$(Mid$(memoria, 87, 10)) & ")"
   
    chkUsuProteq = True
End Function


