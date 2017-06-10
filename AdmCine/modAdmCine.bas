Attribute VB_Name = "modAdmCine"
Option Explicit

Public pbsupervisor     As Boolean
Public piSupervisor     As Integer
Public psSupervisor     As String

Public linha()   As String
Public tplinha() As String

Public iEmp As Integer
Public iCin As Integer
Public bConfirmaDtSistema As Boolean

Sub Main()
    Dim expurgo As New clsExpurgo
    Dim msgErro As String
    Dim intAux  As Integer

    If App.PrevInstance Then
        MsgBox "Já existe uma instância do Administração aberta!", vbCritical, App.ProductName
        End
    End If
    
    bConfirmaDtSistema = False
    
    frmDataSistema.Show vbModal
    
    If Not bConfirmaDtSistema Then
        End
    End If
    
    Call LeRegistro
    

On Error Resume Next
    If Dir(pDirExport, vbDirectory) = "" Then
        MkDir pDirExport
    End If
    
On Error GoTo 0

    dbConnect.Open "FILE NAME=" & App.Path & "\Cinema.udl"
    
    Call varsBaseDados(dbConnect)


    pCheckProtec = 99
    If CInt(pCheckProtec) <> 99 Then
       If Not verificaProteq(dbConnect, pCheckDB, msgErro) Then
          MsgBox msgErro, vbCritical, App.ProductName
          End
       End If
    End If
    
    If Not CarregaParametros() Then End
    
    If Not AjustaDataHora() Then End

    intAux = pTempVend
    pTempVend = intTempoEntreSessoes
    intTempoEntreSessoes = intAux

    'frmLogin.TipoUsuario = GERENTE
    frmLogin.mod_cd = 1
    frmLogin.fun_cd = 1
    frmLogin.Show vbModal

    If strLogin <> "" Then
            
         frmExpurgo.Show vbModal
        'Set expurgo.ConexaoADO = dbConnect
        'expurgo.dias = CInt(pDiasExpurgo)
        'Call expurgo.expurgo
        
        'If expurgo.CodigoErro <> 0 Then
        '    MsgBox "Ocorreu um erro no processo de expurgo: " & expurgo.MensagemErro, vbCritical, App.ProductName
        'End If
        
        MDIFrmAdmCine.Show
    End If
    
End Sub

