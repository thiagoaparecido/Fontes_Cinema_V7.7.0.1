VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envia Movimento"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   2415
      TabIndex        =   2
      Top             =   165
      Width           =   1438
   End
   Begin VB.CommandButton cmdCancela 
      Caption         =   "&Cancela"
      Height          =   390
      Left            =   2415
      TabIndex        =   1
      Top             =   660
      Width           =   1438
   End
   Begin VSFlex7DAOCtl.VSFlexGrid vsfMovtos 
      Height          =   2715
      Left            =   180
      TabIndex        =   0
      Top             =   165
      Width           =   2085
      _cx             =   3678
      _cy             =   4789
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmExport.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   1
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i          As Integer
    Dim movtos     As clsMovtos
    Dim expMovto   As New clsExportBol
    Dim arqDest    As String
    Dim arqDestAux As String
    Dim arqOrig    As String
    Dim arqErr     As String
    Dim erro       As Boolean
    
    On Error GoTo trataErro_cmdOK
    
    If Dir(pDirExport, vbDirectory) = "" Then
        MsgBox "Diretório para gerar arquivo de movimento, não existe!", vbInformation, App.ProductName
        Exit Sub
    End If
    
    erro = False
    
    cmdOK.Enabled = False
    cmdCancela.Enabled = False
    
    Set expMovto.ConexaoADO = dbConnect
    expMovto.DirExoprt = pDirExport
    
    For i = 1 To vsfMovtos.Rows - 1
        If vsfMovtos.Cell(flexcpChecked, i, 0) = vbChecked Then
            Set movtos = New clsMovtos
            movtos.Add CVDate(vsfMovtos.TextMatrix(i, 1))
            
            If expMovto.TransfMovtos(movtos) Then
                arqDestAux = pDirExport & "\" & Format(CVDate(vsfMovtos.TextMatrix(i, 1)), "YYYYMMDD") & "_" & Format("000", iEmpresa) & "_" & Format("000", iCinema) & ".zip"
                arqDest = gGetShortPathName(pDirExport) & "\" & Format(CVDate(vsfMovtos.TextMatrix(i, 1)), "YYYYMMDD") & "_" & Format("000", iEmpresa) & "_" & Format("000", iCinema) & ".zip"
                arqOrig = gGetShortPathName(pDirExport) & "\*.bak"
                arqErr = gGetShortPathName(pDirExport) & "\*.err"
                
                erro = True
                If compactaArquivos(arqOrig, arqDest) Then
                    If expMovto.ConfTransfMovtos(movtos) Then
                        Kill arqOrig
                        'Kill arqErr
                        
                        erro = False
                    End If
                End If
                
                If erro Then
                    Kill arqOrig
                    Kill arqDest
                    Kill arqErr
                    
                    MsgBox "Problemas na geração do arquivo", vbInformation, App.ProductName
                    
                    cmdOK.Enabled = True
                    cmdCancela.Enabled = True
                    
                    Exit Sub
                Else
                    MsgBox "Arquivo: " & arqDestAux & " gerado com sucesso", vbInformation, App.ProductName
                End If
            Else
                MsgBox "Problemas na geração do arquivo: " + expMovto.MensagemErro, vbInformation, App.ProductName
                
                cmdOK.Enabled = True
                cmdCancela.Enabled = True
                
                Exit Sub
            End If
        End If
    Next i
    
    cmdOK.Enabled = True
    cmdCancela.Enabled = True
    
    Unload Me
    
    Exit Sub
trataErro_cmdOK:
    If Err.Number = 53 Then
        Resume Next
    Else
        MsgBox "Problemas na geração do arquivo: " + Err.Description, vbInformation, App.ProductName
        cmdOK.Enabled = True
        cmdCancela.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Call carregaMovtos
End Sub

Private Sub carregaMovtos()
    Dim tBol    As New clsExportBol
    Dim movtos1 As New clsMovtos
    Dim movtos2 As New clsMovtos
    Dim i As Integer
    
    Set tBol.ConexaoADO = dbConnect
    
    Set movtos1 = tBol.movtosTransf
    Set movtos2 = tBol.movtosPTransf
    
    vsfMovtos.Rows = 1
    
    If movtos1.Count > 0 Then
        For i = 0 To movtos1.Count - 1
            vsfMovtos.Rows = vsfMovtos.Rows + 1
            vsfMovtos.Row = vsfMovtos.Rows - 1
            vsfMovtos.TextMatrix(vsfMovtos.Rows - 1, 1) = Format(movtos1.Item(i + 1).dtMovto, "dd/mm/yyyy")
            vsfMovtos.Cell(flexcpForeColor, vsfMovtos.Rows - 1, 0, vsfMovtos.Rows - 1, vsfMovtos.Cols - 1) = &HFF&
        Next i
    End If
    
    If movtos2.Count > 0 Then
        For i = 0 To movtos2.Count - 1
            vsfMovtos.Rows = vsfMovtos.Rows + 1
            vsfMovtos.TextMatrix(vsfMovtos.Rows - 1, 1) = Format(movtos2.Item(i + 1).dtMovto, "dd/mm/yyyy")
        Next i
    End If
End Sub

Private Sub vsfMovtos_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If vsfMovtos.Cell(flexcpForeColor, Row, Col, Row, Col) = &HFF& Then
        vsfMovtos.Cell(flexcpChecked, Row, Col) = False
    End If
End Sub
