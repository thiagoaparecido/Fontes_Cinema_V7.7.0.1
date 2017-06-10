Attribute VB_Name = "modGenerico"
Option Explicit

Public Enum eAlinhamento
    Direita = 0
    esquerda = 1
    Centro = 2
End Enum

Public Enum verPeriodo
    anterior = 0
    meio = 1
    posterior = 2
End Enum
    
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Declare Function GetSystemDirectory& Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal p$, ByVal S%)

Declare Function GetShortPathName Lib "kernel32" _
                        Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
                        ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long

Declare Function OpenProcess Lib "kernel32" _
                             (ByVal dwDesiredAccess As Long, _
                             ByVal bInheritHandleas As Long, _
                             ByVal dwProcessId As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" _
                                     (ByVal hHandle As Long, _
                                      ByVal dwMilliseconds As Long) As Long

Public Const giMAX_PATH = 255
'Constantes API 32
Const SYNCHRONIZE = &H100000
Const INFINITE = &HFFFF

'Public cineProt       As New cineProteg

Public dbConnect      As New ADODB.Connection

Public servidor       As String
Public baseDados      As String
Public usuarioDB      As String
Public senhaDB        As String

Public intTempoEntreSessoes As Integer
Public dtHoraMaxSessao      As Date
Public dtHoraLimite         As Date
Public dtHoraLimite12       As Date
Public dtHoraLimite23       As Date
Public dtHoraLimite34       As Date
Public dtHoraLimite45       As Date
Public dtHoraLimite56       As Date
Public bImprimeCodBarra     As Boolean
Public bImprimeLocacao      As Boolean
Public bImprimeEndereco     As Boolean
Public bImprimeCNPJ         As Boolean
Public bImprimeIE           As Boolean
Public bImprimeTck          As Boolean
Public bImprimeMFMI         As Boolean
Public dCustoIngresso       As Double
Public dImpostoMun          As Double
Public dDireitosAut         As Double
Public dOutros              As Double
Public dPercentualMeias     As Double
Public dPercentualCortesias As Double
Public sMsg1                As String
Public sMsg2                As String
Public sMsg3                As String
Public pdtSistema           As Date
Public iEmpresa             As Integer
Public nEmpresa             As String
Public iCinema              As Integer
Public nCinema              As String

Public intUsuario As Integer
Public strLogin   As String
Public strUsuario As String

Public strDataRef1 As String
Public strDataRef2 As String

Public Enum eTipoUsuario
    ADMINISTRADOR = 9
    GERENTE = 8
    Caixa = 1
End Enum
    
Public aDiaSemana(8) As String
Public Const CPS_SEPARADOR = "€"

Public Const COLUNAS_IMP = 38

Public LinhaNormal
Public LinhaTitulo
Public LinhaDupla

Public pCommPort    As String
Public pBaudRate    As String
Public pParity      As String
Public pDataBit     As String
Public pStopBit     As String
Public pTempImp     As String
Public pTempImp2    As String
Public pCheckProtec As String
Public pCheckDB     As String
Public pCaixaTalao  As Boolean
Public pTempComb    As String
Public pbErroImp    As Boolean
Public lotadoEXE    As String
Public pTempVend    As String
Public pDirExport   As String
Public pZipProg     As String
Public pImpSessao   As String
Public pDiasExpurgo As String
Public pDiasBackup   As String
Public pUltimoBackup As String
Public TelaPoltronas As String

Public pPortaBematech As String
Public pImpressEpsom  As String

Public pImpressora    As String

      
Public Sub CarregaGridAutomatico(vsfDados As VSFlexGrid, oRs As ADODB.Recordset, Optional ByVal lRow As Long)

    Dim i As Integer
    
    
    vsfDados.Rows = 1
    vsfDados.Cols = oRs.Fields.Count
    
    For i = 0 To oRs.Fields.Count - 1

        Select Case oRs.Fields(i).Type
            Case adNumeric
                vsfDados.ColAlignment(i) = flexAlignRightCenter
                vsfDados.ColDataType(i) = flexDTLong
                vsfDados.ColFormat(i) = "0" & IIf(oRs.Fields(i).NumericScale = 0, "", "." & String(oRs.Fields(i).NumericScale, "0"))
            Case adVarChar
                vsfDados.ColAlignment(i) = flexAlignLeftCenter
                vsfDados.ColDataType(i) = flexDTString
            Case adVarNumeric
                vsfDados.ColAlignment(i) = flexAlignRightCenter
                vsfDados.ColDataType(i) = flexDTLong
            Case adDBTimeStamp
                vsfDados.ColAlignment(i) = flexAlignCenterCenter
                vsfDados.ColDataType(i) = flexDTDate
            Case adBoolean
                vsfDados.ColAlignment(i) = flexAlignCenterCenter
                vsfDados.ColDataType(i) = flexDTBoolean
            Case adCurrency
                vsfDados.ColAlignment(i) = flexAlignRightCenter
                vsfDados.ColDataType(i) = flexDTCurrency
                vsfDados.ColFormat(i) = "#,###.##"
        End Select
        
        vsfDados.ColKey(i) = oRs.Fields(i).Name
        vsfDados.TextMatrix(0, i) = oRs.Fields(i).Name
        
    Next
    
    If Not oRs.EOF Then
        vsfDados.LoadArray oRs.GetRows
        oRs.Close
        If lRow <> 0 Then
            vsfDados.row = lRow
        Else
            vsfDados.row = 1
        End If
    End If
    
    vsfDados.AutoSizeMode = flexAutoSizeColWidth
    Call vsfDados.AutoSize(0, vsfDados.Cols - 1)

End Sub

Public Function CarregaParametros() As Boolean

    On Error GoTo CarregaParametros_Erro
        

    Call CarregaDiaSemana
    
    strDataRef1 = Format(CDate("01/01/1900"), "Short Date")
    strDataRef2 = Format(DateAdd("d", 1, CDate("01/01/1900")), "Short Date")
    

    Dim oRs As New ADODB.Recordset
    Dim clsTB_PARAMETRO As New Cine2005.clsTB_PARAMETRO
    

    
    Set clsTB_PARAMETRO.ConexaoADO = dbConnect
    
    If Not clsTB_PARAMETRO.Selecionar(oRs) Then
        MsgBox clsTB_PARAMETRO.MensagemErro, vbCritical, App.ProductName
        GoTo CarregaParametros_Fim
    End If
    

    If oRs.EOF() Then
        intTempoEntreSessoes = 15
        dtHoraMaxSessao = "03:00"
        dtHoraLimite = "15:00"
        dtHoraLimite12 = "15:00"
        dtHoraLimite23 = "15:00"
        dtHoraLimite34 = "15:00"
        dtHoraLimite45 = "15:00"
        dtHoraLimite56 = "15:00"
        bImprimeCodBarra = False
        bImprimeLocacao = False
        bImprimeEndereco = False
        bImprimeCNPJ = False
        bImprimeIE = False
        bImprimeTck = False
        bImprimeMFMI = False
        dCustoIngresso = 0
        dImpostoMun = 0
        dDireitosAut = 0
        dOutros = 0
        dPercentualMeias = 0
        dPercentualCortesias = 0
        sMsg1 = ""
        sMsg2 = ""
        sMsg3 = ""
    Else
        intTempoEntreSessoes = oRs.Fields("par_tmp_ses")
        dtHoraMaxSessao = oRs.Fields("par_hora_max_ses")
        
        dtHoraLimite = oRs.Fields("par_hora_limite")
        dtHoraLimite12 = oRs.Fields("par_hora_limite12")
        dtHoraLimite23 = oRs.Fields("par_hora_limite23")
        dtHoraLimite34 = oRs.Fields("par_hora_limite34")
        dtHoraLimite45 = oRs.Fields("par_hora_limite45")
        dtHoraLimite56 = oRs.Fields("par_hora_limite56")
        
        bImprimeCodBarra = IIf(IsNull(oRs.Fields("par_imp_cod_barra")), False, oRs.Fields("par_imp_cod_barra"))
        bImprimeLocacao = IIf(IsNull(oRs.Fields("par_imp_lotacao")), False, oRs.Fields("par_imp_lotacao"))
        bImprimeEndereco = IIf(IsNull(oRs.Fields("par_imp_endereco")), False, oRs.Fields("par_imp_endereco"))
        bImprimeCNPJ = IIf(IsNull(oRs.Fields("par_imp_CNPJ")), False, oRs.Fields("par_imp_CNPJ"))
        bImprimeIE = IIf(IsNull(oRs.Fields("par_imp_IE")), False, oRs.Fields("par_imp_IE"))
        bImprimeTck = IIf(IsNull(oRs.Fields("par_imp_tck")), False, oRs.Fields("par_imp_tck"))
        
        bImprimeMFMI = IIf(IsNull(oRs.Fields("par_imp_MFIM")), False, oRs.Fields("par_imp_MFIM"))
        
        dCustoIngresso = IIf(IsNull(oRs.Fields("par_custo_ingresso")), 0, oRs.Fields("par_custo_ingresso"))
        dImpostoMun = IIf(IsNull(oRs.Fields("par_imposto_mun")), 0, oRs.Fields("par_imposto_mun"))
        dDireitosAut = IIf(IsNull(oRs.Fields("par_direitos_aut")), 0, oRs.Fields("par_direitos_aut"))
        dOutros = IIf(IsNull(oRs.Fields("par_outros")), 0, oRs.Fields("par_outros"))
        dPercentualMeias = IIf(IsNull(oRs.Fields("par_perc_meias")), 0, oRs.Fields("par_perc_meias"))
        dPercentualCortesias = IIf(IsNull(oRs.Fields("par_perc_cortesias")), 0, oRs.Fields("par_perc_cortesias"))
        sMsg1 = IIf(IsNull(oRs.Fields("par_msg1")), "", oRs.Fields("par_msg1"))
        sMsg2 = IIf(IsNull(oRs.Fields("par_msg2")), "", oRs.Fields("par_msg2"))
        sMsg3 = IIf(IsNull(oRs.Fields("par_msg3")), "", oRs.Fields("par_msg3"))
    End If

    oRs.Close
    
    Call carregaEmpresa
    Call carregaCinema
    
    CarregaParametros = True
    
    GoTo CarregaParametros_Fim
    
CarregaParametros_Erro:
    MsgBox "Erro de execução! 'CarregaParametros'", vbCritical, App.ProductName
    
CarregaParametros_Fim:
    If oRs.State = 1 Then oRs.Close
    Set oRs = Nothing
    Set clsTB_PARAMETRO = Nothing
   
End Function

Public Function AjustaDataHora() As Boolean

    On Error GoTo AjustaDataHora_Erro
        
    Dim clsGerais As New Cine2005.clsGerais
    
    Set clsGerais.ConexaoADO = dbConnect
    
    Date = clsGerais.DataServidor
    Time = clsGerais.HoraServidor
    
    AjustaDataHora = True
    
    GoTo AjustaDataHora_Fim
    
AjustaDataHora_Erro:
    MsgBox "Erro de execução! 'AjustaDataHora'", vbCritical, App.ProductName
    
AjustaDataHora_Fim:
    Set clsGerais = Nothing
   
End Function

Public Sub CarregaDiaSemana()
    aDiaSemana(2) = "Segunda"
    aDiaSemana(3) = "Terça"
    aDiaSemana(4) = "Quarta"
    aDiaSemana(5) = "Quinta"
    aDiaSemana(6) = "Sexta"
    aDiaSemana(7) = "Sábado"
    aDiaSemana(1) = "Domingo"
    aDiaSemana(8) = "Feriado"
End Sub

Public Function PegaColuna(ByVal sString As String, ByVal iCol As Integer, Optional ByVal sSep As String) As String

    Dim sColString As String
    Dim iCount As Integer
    Dim iPos As Integer
    Dim iColAtual As Integer
    
    sSep = IIf(sSep = "", CPS_SEPARADOR, sSep)
        
    sString = Trim$(sString)
    sString = IIf(Right$(sString, 1) <> sSep, sString & sSep, sString)
    
    iPos = 1
    
    For iCount = 1 To iCol
        iPos = InStr(1, sString, sSep)
        If iPos = 0 Then
            sColString = ""
            Exit For
        End If
        sColString = Left$(sString, iPos - 1)
        sString = Right$(sString, Len(sString) - iPos)
    Next
        
    PegaColuna = sColString
    
End Function

Public Function EliminaAcentos(sString As String) As String
    Dim i   As Integer
    Dim str As String

    str = ""
    
    For i = 1 To Len(sString)
        Select Case Mid(sString, i, 1)
            Case "á", "à", "ã", "â", "ä"
                str = str & "a"
            Case "Á", "À", "Ã", "Â", "Ä"
                str = str & "A"
            Case "é", "è", "ê", "ë"
                str = str & "e"
            Case "É", "È", "Ê", "Ë"
                str = str & "E"
            Case "í", "ì", "î", "ï"
                str = str & "i"
            Case "Í", "Ì", "Î", "Ï"
                str = str & "I"
            Case "ó", "ò", "õ", "ô", "ö"
                str = str & "o"
            Case "Ó", "Ò", "Õ", "Ô", "Ö"
                str = str & "O"
            Case "ú", "ù", "û", "ü"
                str = str & "u"
            Case "Ú", "Ù", "Û", "Ü"
                str = str & "U"
            Case "ñ"
                str = str & "n"
            Case "Ñ"
                str = str & "N"
            Case "ç"
                str = str & "c"
            Case "Ç"
                str = str & "C"
            Case "ÿ", "ý"
                str = str & "y"
            Case "Ÿ", "Ý"
                str = str & "Y"
            Case Else
                If Asc(Mid(sString, i, 1)) >= 32 And _
                   Asc(Mid(sString, i, 1)) <= 126 Then
                    str = str & Mid(sString, i, 1)
                Else
                    str = str & " "
                End If
        End Select
    Next i

    EliminaAcentos = str
End Function

Public Function Alinha(ByVal texto As String, ByVal Tamanho As Integer, ByVal Alinhamento As eAlinhamento, Optional ByVal Caracter As String) As String
    
    If Caracter = "" Then Caracter = " "
    
    If (Tamanho - Len(texto)) <= 0 Then
        Alinha = Mid(texto, 1, Tamanho)
    Else
        Select Case Alinhamento
            Case esquerda
                Alinha = texto & String(Tamanho - Len(texto), Caracter)
            Case Direita
                Alinha = String(Tamanho - Len(texto), Caracter) & texto
            Case Centro
                Alinha = String((Tamanho - Len(texto) / 2), Caracter) & texto
        End Select
    End If

End Function

Public Sub SoNumero(KeyAscii As Integer)
    If Not IsNumeric(Chr$(KeyAscii)) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Public Sub AjustaCOM(ByRef MSC As MSComm)

    If Not MSC.PortOpen Then
        MSC.CommPort = pCommPort
        MSC.Settings = pBaudRate & "," & pParity & "," & pDataBit & "," & pStopBit
        MSC.PortOpen = True
    End If
End Sub

Public Function ImpOk(ByRef MSC As MSComm, ByRef ErroImp As Boolean) As Boolean

    Dim erro As Integer
    Dim lInicio As Long
    Dim impLimpa

    impLimpa = MSC.Input
    ErroImp = True
    
    'Checa para ver se a impressora esta ligada e com papel
    MSC.Output = Chr$(27) & Chr$(118)
    
    lInicio = timeGetTime
    
    Do While ErroImp
        DoEvents
        If lInicio + CInt(pTempImp) < timeGetTime Then
           Exit Do
        End If
    Loop
    
    ImpOk = Not ErroImp
   
End Function

Public Sub LeRegistro()

    Dim Registry As New Cine2005.ManipulaRegistry 'Variável para permitir a leitura do Registry
    
    Dim x As Integer
    
    Registry.LeString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "CommPort", pCommPort
    
    If pCommPort = "" Then
        pCommPort = 1
        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "CommPort", pCommPort
    End If
    
    Call Registry.LeString(RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "BaudRate", pBaudRate)
    If pBaudRate = "" Then
        pBaudRate = 19200
        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "BaudRate", pBaudRate
    End If
    
    Call Registry.LeString(RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "Parity", pParity)
    
    If pParity = "" Then
        pParity = "n"
        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "Parity", pParity
    End If
    
    Call Registry.LeString(RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "DataBit", pDataBit)
    
    If pDataBit = "" Then
        pDataBit = 8
        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "DataBit", pDataBit
    End If
    Call Registry.LeString(RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "StopBit", pStopBit)
    
    If pStopBit = "" Then
        pStopBit = 1
        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "StopBit", pStopBit
    End If

    Call Registry.LeString(RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "TempImp", pTempImp)
    
    If pTempImp = "" Then
        pTempImp = 2500
        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "TempImp", pTempImp
    End If
    
    Call Registry.LeString(RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "TempImp2", pTempImp2)
    
    If pTempImp2 = "" Then
        pTempImp2 = 300
        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "TempImp2", pTempImp2
    End If
    
    Call Registry.LeString(RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "TempComb", pTempComb)
'
'    If pTempComb = "" Then
'        pTempComb = 10
'        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "TempComb", pTempComb
'    End If

    Call Registry.LeString(RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "CheckProtec", pCheckProtec)
    
    If pCheckProtec = "" Then
        pCheckProtec = 1
        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "CheckProtec", pCheckProtec
    End If
    
    Call Registry.LeString(RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "CheckDB", pCheckDB)
    If pCheckDB = "" Then
        If App.ProductName = "Administração" Then
           pCheckDB = 1
        Else
           pCheckDB = 77
        End If
        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "CheckDB", pCheckDB
    End If
    
    Call Registry.LeString(RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "TempVend", pTempVend)
    
    If pTempVend = "" Then
        pTempVend = 20
        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "TempVend", pTempVend
    End If
    
   
    Call Registry.LeString(RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "lotado", lotadoEXE)
   
    If lotadoEXE = "" Then
        lotadoEXE = ""
        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "lotado", lotadoEXE
    End If
    
    Call Registry.LeString(RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "DirExport", pDirExport)
    
    If pDirExport = "" Then
        pDirExport = "C:\MovtoExport"
        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "DirExport", pDirExport
    End If
    
    
    Call Registry.LeString(RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "ZipProg", pZipProg)
    
    If pZipProg = "" Then
        pZipProg = "C:\Arquivos de programas\7-ZIP\7Z.EXE"
        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "ZipProg", pZipProg
    End If
    
    Call Registry.LeString(RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "ImpSessao", pImpSessao)
    
    If pImpSessao = "" Then
        pImpSessao = "S"
        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "ImpSessao", pImpSessao
    End If
    
    Call Registry.LeString(RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "DiasExpurgo", pDiasExpurgo)
    
    If pDiasExpurgo = "" Then
        pDiasExpurgo = "200"
        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "DiasExpurgo", pDiasExpurgo
    End If
    
    Call Registry.LeString(RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "Impressora", pImpressora)
    
    If pImpressora = "" Then
        pImpressora = "Epson"
        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "Impressora", pImpressora
    End If
    
    Call Registry.LeString(RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "ImpressEpsom", pImpressEpsom)
    
    If pImpressEpsom = "" Then
        pImpressEpsom = "TM-T88IIIS1"
        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "ImpressEpsom", pImpressEpsom
    End If
    
    Call Registry.LeString(RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "PortaBematech", pPortaBematech)
    
    If pPortaBematech = "" Then
        pPortaBematech = "COM1"
        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "PortaBematech", pPortaBematech
    End If
    
    Call Registry.LeString(RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "DiasBackup", pDiasBackup)
    
    If pDiasBackup = "" Then
        pDiasBackup = "30"
        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "DiasBackup", pDiasBackup
    End If

    Call Registry.LeString(RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "UltimoBackup", pUltimoBackup)
    
    If pUltimoBackup = "" Then
        pUltimoBackup = "01/01/1900"
        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "UltimoBackup", pUltimoBackup
    End If
    
    Call Registry.LeString(RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "TelaPoltronas", TelaPoltronas)
    
    If TelaPoltronas = "" Then
        TelaPoltronas = "N"
        Registry.GravaString RegKey_LOCAL_MACHINE, "Software\" & App.ProductName, "TelaPoltronas", TelaPoltronas
    End If
End Sub

Public Sub varsBaseDados(dbConnect As ADODB.Connection)
    Dim connectStr As String
    Dim parmStr()  As String
    Dim i          As Integer
    
    connectStr = dbConnect.ConnectionString
    parmStr = Split(connectStr, ";")
    
    For i = LBound(parmStr) To UBound(parmStr)
        If InStr(parmStr(i), "Password") > 0 Then
            senhaDB = Mid(parmStr(i), InStr(parmStr(i), "=") + 1)
        ElseIf InStr(parmStr(i), "User ID") > 0 Then
            usuarioDB = Mid(parmStr(i), InStr(parmStr(i), "=") + 1)
        ElseIf InStr(parmStr(i), "Initial Catalog") > 0 Then
            baseDados = Mid(parmStr(i), InStr(parmStr(i), "=") + 1)
        ElseIf InStr(parmStr(i), "Data Source") > 0 Then
            servidor = Mid(parmStr(i), InStr(parmStr(i), "=") + 1)
        End If
    Next i
End Sub

Public Function diretorioWindows() As String
    Dim strSysDirPath As String
    
    strSysDirPath = String$(145, Chr$(0))
    strSysDirPath = Left$(strSysDirPath, GetSystemDirectory(strSysDirPath, Len(strSysDirPath)))

    diretorioWindows = strSysDirPath
End Function

Public Function gGetShortPathName(ByVal vsLongPath As String) As String
   Dim sShortPath As String
   Dim lPathLen   As Long
   Dim iLen       As Long
   
   sShortPath = Space$(giMAX_PATH)
   iLen = Len(sShortPath)
   lPathLen = GetShortPathName(vsLongPath, sShortPath, iLen)
   gGetShortPathName = Left$(sShortPath, lPathLen)
End Function



Public Function verificaPeriodo(ByVal dtIni As Date, ByVal dtFim As Date, ByVal dtRef) As verPeriodo
    If dtFim < dtRef Then
        verificaPeriodo = verPeriodo.anterior
    ElseIf dtIni <= dtRef And dtFim >= dtRef Then
        verificaPeriodo = verPeriodo.meio
    Else
        verificaPeriodo = verPeriodo.posterior
    End If
End Function

Public Function compactaArquivos(strSourceFiles As String, strDestFile As String) As Boolean
    Dim ret          As Long
    Dim CheckName    As String
    Dim HwndOfWinzip As Long
    Dim strAux       As String
    Dim pos          As Integer
    Dim lProcessHandle As Long
    Dim lDummy         As Long
    
    
    On Error GoTo TrataErro_compactaArquivos
    
    compactaArquivos = False

    'ret = Shell(pZipProg & " -a " & strDestFile & " " & strSourceFiles, vbHide)
    ret = Shell(pZipProg & " a " & strDestFile & " " & strSourceFiles, vbHide)
    
    If ret <> 0 Then
        lProcessHandle = OpenProcess(SYNCHRONIZE, True, ret)
        lDummy = WaitForSingleObject(lProcessHandle, INFINITE)
        
        compactaArquivos = True
    End If
    
    Exit Function
    
TrataErro_compactaArquivos:
    compactaArquivos = False
    Exit Function
    Resume 0
End Function

Public Function expandeArquivos(strSourceFiles As String, strDestFile As String) As Boolean
    Dim ret          As Long
    Dim CheckName    As String
    Dim HwndOfWinzip As Long
    Dim strAux       As String
    Dim pos          As Integer
    Dim lProcessHandle As Long
    Dim lDummy         As Long
    
    
    On Error GoTo TrataErro_expandeArquivos
    
    expandeArquivos = False

    ret = Shell(pZipProg & " e -aoa " & strSourceFiles & " -o" & strDestFile, vbHide)
    'ret = Shell(pZipProg & " -e " & strSourceFiles & " " & strDestFile, vbHide)
    
    If ret <> 0 Then
        lProcessHandle = OpenProcess(SYNCHRONIZE, True, ret)
        lDummy = WaitForSingleObject(lProcessHandle, INFINITE)
        
        expandeArquivos = True
    End If
    
    Exit Function
    
TrataErro_expandeArquivos:
    expandeArquivos = False
    Exit Function
    Resume 0
End Function

Public Sub carregaEmpresa()
    Dim oRs As New ADODB.Recordset
    Dim clsTB_EMPRESA As New Cine2005.clsTB_EMPRESA
    
    On Error GoTo carregaEmpresa_Erro
    
    Set clsTB_EMPRESA.ConexaoADO = dbConnect
    
    If Not clsTB_EMPRESA.Selecionar(oRs) Then
        Exit Sub
    End If

    iEmpresa = oRs.Fields("emp_cd")
    nEmpresa = oRs.Fields("emp_nm")

carregaEmpresa_Erro:
    If oRs.State = 1 Then oRs.Close
    
    Set oRs = Nothing
    Set clsTB_EMPRESA = Nothing

End Sub

Public Sub carregaCinema()
    Dim oRs As New ADODB.Recordset
    Dim clsTB_CINEMA As New Cine2005.clsTB_CINEMA
    
    On Error GoTo carregaCinema_Erro
    
    Set clsTB_CINEMA.ConexaoADO = dbConnect
    
    If Not clsTB_CINEMA.Selecionar(oRs) Then
        Exit Sub
    End If

    iCinema = oRs.Fields("cin_cd")
    nCinema = oRs.Fields("cin_nm")

carregaCinema_Erro:
    If oRs.State = 1 Then oRs.Close
    
    Set oRs = Nothing
    Set clsTB_CINEMA = Nothing

End Sub

