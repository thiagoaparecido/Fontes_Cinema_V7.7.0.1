Attribute VB_Name = "modAdmCentral"
Option Explicit

Public pDirImport As String

Sub Main()

    If App.PrevInstance Then
        MsgBox "Já existe uma instância do Administração Central aberta!", vbCritical, App.ProductName
        End
    End If
    
'    Dim msgErro As String
    Call LeRegistro

    pDirImport = App.Path & "\MovtoImport"
    
    If Dir(pDirImport, vbDirectory) = "" Then
        MkDir pDirImport
    End If

'    dbConnect.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=cine2005;Data Source=(local)"  '    App.Path & "\cine2005.udl"
    
    dbConnect.Open "FILE NAME=" & App.Path & "\AdmCentral.udl"
    
    Call varsBaseDados(dbConnect)

'    If chechProteg <> 99 Then
'       If Not cineProt.verificaProteq(dbConnect, checkDataAcess, msgErro) Then
'          MsgBox msgErro, vbCritical, App.ProductName
'          End
'       End If
'    End If
    
'    If Not CarregaParametros() Then End
    If Not AjustaDataHora() Then End

'    frmLogin.TipoUsuario = GERENTE
'    frmLogin.Show vbModal
    
'    If strLogin <> "" Then
        MDIFrmAdmCentral.Show
'    End If
    
End Sub

