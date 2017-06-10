VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImporta 
   Caption         =   "Importa Movimento"
   ClientHeight    =   1305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1305
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdArq 
      Caption         =   "..."
      Height          =   315
      Left            =   7410
      TabIndex        =   3
      Top             =   90
      Width           =   315
   End
   Begin VB.TextBox txtArq 
      Height          =   315
      Left            =   915
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   90
      Width           =   6450
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2175
      TabIndex        =   1
      Top             =   630
      Width           =   1590
   End
   Begin VB.CommandButton cmdCancela 
      Cancel          =   -1  'True
      Caption         =   "Cancela"
      Height          =   420
      Left            =   4095
      TabIndex        =   0
      Top             =   630
      Width           =   1590
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6270
      Top             =   585
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      MaxFileSize     =   32767
   End
   Begin VB.Label lblArq 
      AutoSize        =   -1  'True
      Caption         =   "Arquivo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   75
      TabIndex        =   4
      Top             =   135
      Width           =   720
   End
End
Attribute VB_Name = "frmImporta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdArq_Click()
    Dim sMsg As String
   
   On Error GoTo TrataErro
   
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = App.ProductName
   CommonDialog1.InitDir = pDirImport
   CommonDialog1.Filter = "Arquivo Compactado (*.zip)|*.zip"
   CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNShareAware Or _
                         cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNAllowMultiselect
                         
   CommonDialog1.ShowOpen
   
   txtArq.Text = CommonDialog1.FileName
   
   Exit Sub
TrataErro:
   If Err.Number <> cdlCancel Then
      sMsg = "Ocorreu um erro em cmdArq_Click." & vbCrLf
      sMsg = sMsg & Err.Number & " - " & Err.Description
      MsgBox sMsg, vbCritical, App.ProductName
   End If
End Sub

Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim impotMov As New clsImportMovto
    Dim dtMovto  As Date
    Dim empCd    As Integer
    Dim cinCd    As Integer
    Dim msgErr   As String
    Dim arqs     As Variant
    Dim strAqr   As String
    Dim i        As Integer
    Dim carrega  As Boolean
    Dim n        As Integer
    
    If txtArq.Text = Empty Then Exit Sub
    
    cmdOK.Enabled = False
    cmdCancela.Enabled = False
    
    Set impotMov.ConexaoADO = dbConnect
    impotMov.DirTrab = pDirImport
    
    arqs = Split(txtArq.Text, " ")
    
    If LBound(arqs) < UBound(arqs) Then
        n = 1
    Else
        n = 0
    End If
    
    
    For i = LBound(arqs) + n To UBound(arqs)
        If LBound(arqs) < UBound(arqs) Then
            strAqr = arqs(0) & arqs(i)
        Else
            strAqr = arqs(0)
        End If
        
        carrega = True
        
        If Dir(strAqr) <> "" Then
            If verificaNomeArq(strAqr, dtMovto, empCd, cinCd, msgErr) Then
                If impotMov.existeMovto(dtMovto, empCd, cinCd) Then
                    If MsgBox("Este movimento (" & arqs(i) & ") já foi carregado. Deseja sobrepor?(S/N)", vbYesNo, App.ProductName) = vbYes Then
                        If Not impotMov.excluiMovto(dtMovto, empCd, cinCd) Then
                            MsgBox "Problemas na exclusão do movimento: " + impotMov.MensagemErro, vbInformation, App.ProductName
                            
                            cmdOK.Enabled = True
                            cmdCancela.Enabled = True
                            
                            Exit Sub
                        End If
                    Else
                        carrega = False
                    End If
                ElseIf impotMov.CodigoErro <> 1 Then
                    MsgBox "Problemas na importação: " + impotMov.MensagemErro, vbInformation, App.ProductName
                    
                    cmdOK.Enabled = True
                    cmdCancela.Enabled = True
                    
                    Exit Sub
                End If
                
                If carrega Then
                    If Not impotMov.ImportMovto(strAqr) Then
                        MsgBox "Problemas na importação: " + impotMov.MensagemErro, vbInformation, App.ProductName
                        
                        Exit Sub
                    End If
                End If
            Else
                MsgBox msgErr, vbInformation, App.ProductName
                
                Exit Sub
            End If
        Else
            MsgBox "Arquivo " & arqs(i) & " não encontrado!", vbInformation, App.ProductName
            
            Exit Sub
        End If
    Next i
    
    MsgBox "Importação realizada com sucesso", vbInformation, App.ProductName
    
    cmdOK.Enabled = True
    cmdCancela.Enabled = True
    
    Unload Me
End Sub

Private Function verificaNomeArq(ByVal nmArq As String, ByRef dtMovto As Date, ByRef empCd As Integer, ByRef cinCd As Integer, ByRef msgErr As String) As Boolean
    Dim strAux  As String
    Dim chrAux  As String
    Dim strData As String
    Dim strEmp  As String
    Dim strCin  As String
    Dim iPos    As Integer
    Dim iPos2   As Integer

    On Error GoTo TrataErro

    chrAux = ""
    verificaNomeArq = False
    
    If InStr(nmArq, "\") > 0 Then
        chrAux = "\"
    ElseIf InStr(nmArq, ":") > 0 Then
        chrAux = ":"
    End If
    
    If chrAux <> "" Then
        iPos = 0
        Do While InStr(iPos + 1, nmArq, chrAux) > 0
            iPos = InStr(iPos + 1, nmArq, chrAux)
        Loop
    End If
    
    strAux = Mid(nmArq, iPos + 1)
    
    If InStr(strAux, "_") > 0 Then
        iPos = InStr(strAux, "_")
        strData = Mid(strAux, 1, iPos - 1)
    Else
        msgErr = "Formato do nome do arquivo invalido!"
        Exit Function
    End If
    
    If InStr(iPos + 1, strAux, "_") > 0 Then
        iPos2 = InStr(iPos + 1, strAux, "_")
        strEmp = Mid(strAux, iPos + 1, iPos2 - (iPos + 1))
    Else
        msgErr = "Formato do nome do arquivo invalido!"
        Exit Function
    End If
    
    iPos = iPos2
    
    If InStr(iPos + 1, strAux, ".zip") > 0 Then
        iPos2 = InStr(iPos + 1, strAux, ".zip")
        strCin = Mid(strAux, iPos + 1, iPos2 - (iPos + 1))
    Else
        msgErr = "Formato do nome do arquivo invalido!"
        Exit Function
    End If
    
    If Len(strData) <> 8 Or (Not IsNumeric(strData)) Then
        msgErr = "Formato do nome do arquivo invalido!"
        Exit Function
    Else
        dtMovto = DateSerial(CInt(Mid(strData, 1, 4)), CInt(Mid(strData, 5, 2)), CInt(Mid(strData, 7, 2)))
    End If
    
    If IsNumeric(strEmp) Then
        empCd = CInt(strEmp)
    Else
        msgErr = "Formato do nome do arquivo invalido!"
        Exit Function
    End If
    
    If IsNumeric(strCin) Then
        cinCd = CInt(strCin)
    Else
        msgErr = "Formato do nome do arquivo invalido!"
        Exit Function
    End If
    
    verificaNomeArq = True
    
    Exit Function
TrataErro:
        msgErr = "Formato do nome do arquivo invalido!"

End Function

