Attribute VB_Name = "modBancoDados"
Option Explicit

Public gsProvider  As String               'Provaider utilizado para conexão com o banco de dados
Public gConnect    As New ADODB.Connection 'Conexão com o banco de dados
Public bBaseAberta As Boolean
Public cmdConLotacao  As New ADODB.Command
Public dbConnect   As New ADODB.Connection


Public Function gbAbreBase() As Boolean
    Dim lErro     As Long
    Dim sMsg      As String
    
    On Error GoTo TrataErro
    
    gbAbreBase = False
    bBaseAberta = False
        
    gsProvider = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                "Data Source=" & gsBandoDados & "cineprog.mdb;" & _
                "Jet OLEDB:Database Password=" & senhaBanco
   
    'gsProvider = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    '             "Data Source=" & gsBandoDados & "cineprog.mdb;" & _
    '             "User Id=admin;" & _
    '             "Password="
    
    gConnect.Open gsProvider

    gbAbreBase = True
    bBaseAberta = True
    
    Exit Function
    
TrataErro:
    sMsg = "Ocorreu um erro em gbAbreBase." & vbCrLf
    sMsg = sMsg & Err.Number & " - " & Err.Description
    MsgBox sMsg, vbCritical, App.ProductName
End Function

Public Sub gFechaBase()

    'If gRSDados1.State = adStateOpen Then
    '    gRSDados1.Close
    'End If
    'gConnect.Close
    'bBaseAberta = False
End Sub

Public Function verificaDtUltAcessBanco(dtUltAcess As String) As Boolean
    Dim tbUltAcess As New ADODB.Recordset
    Dim dtAux      As String
    
    verificaDtUltAcessBanco = False
    
    tbUltAcess.Open "tbUltimoAcess", gConnect, adOpenDynamic, adLockOptimistic, adCmdTableDirect
    
    If tbUltAcess.EOF And tbUltAcess.BOF Then
        tbUltAcess.Close
        
        Exit Function
    End If

    dtAux = gsDecripto(tbUltAcess.Fields("ultimoAcess").Value)
    
    If dtAux <> "31121950000000" Then
        If dtAux <> dtUltAcess Then
            tbUltAcess.Close
            
            Exit Function
        End If
    End If
    
    verificaDtUltAcessBanco = True
End Function

Public Function atuDtUltAcessBanco(dtUltAcess As String) As Boolean
    Dim tbUltAcess As New ADODB.Recordset
    Dim dtAux      As String
    
    atuDtUltAcessBanco = False
    
    tbUltAcess.Open "tbUltimoAcess", gConnect, adOpenDynamic, adLockOptimistic, adCmdTableDirect
    
    If tbUltAcess.EOF And tbUltAcess.BOF Then
        tbUltAcess.Close
        
        Exit Function
    End If

    dtAux = gsCripto(dtUltAcess)
    
    tbUltAcess.Fields("ultimoAcess").Value = dtAux
    tbUltAcess.Update
    tbUltAcess.Close
    
    atuDtUltAcessBanco = True
End Function

Public Function consLotacaoSel(sala As Integer, _
                               filme As Long, _
                               datasessao As Date, _
                               horasessao As Date, _
                               ByRef Lugares As Integer, _
                               ByRef Meias As Integer, _
                               ByRef Cortesias As Integer) As Integer

  
   cmdConLotacao.Parameters.Item("@sal_cd").Value = sala
   cmdConLotacao.Parameters.Item("@fil_cd").Value = filme
   cmdConLotacao.Parameters.Item("@sre_data").Value = datasessao
   cmdConLotacao.Parameters.Item("@sre_horario").Value = horasessao
   
   cmdConLotacao.Execute
   
   If cmdConLotacao.Parameters.Item("@Erro").Value <> 0 Then
      'cmdConLotacao.Parameters.Item("@MsgErr").Value
      consLotacaoSel = -1
      Lugares = 0
      Meias = 0
      Cortesias = 0
   Else
      consLotacaoSel = cmdConLotacao.Parameters.Item("@sea_lugares_ven").Value - cmdConLotacao.Parameters.Item("@sea_lugares_sel").Value
      Lugares = cmdConLotacao.Parameters.Item("@sre_lugares").Value
      Meias = cmdConLotacao.Parameters.Item("@sre_meias").Value + cmdConLotacao.Parameters.Item("@sea_meias").Value
      Cortesias = cmdConLotacao.Parameters.Item("@sre_cortesias").Value + cmdConLotacao.Parameters.Item("@sea_cortesias").Value
   End If
End Function

