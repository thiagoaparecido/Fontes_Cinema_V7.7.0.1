Attribute VB_Name = "bematch"
Option Explicit

' Declaração das Funções da DLL

Public Declare Function AcionaGuilhotina Lib "mp2032.dll" (ByVal Modo As Integer) As Integer
Public Declare Function AutenticaDoc Lib "mp2032.dll" (ByVal BufTras As String, ByVal Tempo As Integer) As Integer
Public Declare Function BematechTX Lib "mp2032.dll" (ByVal BufTrans As String) As Integer
Public Declare Function CaracterGrafico Lib "mp2032.dll" (ByVal Buffer As String, ByVal TamBuffer As Integer) As Integer
Public Declare Function ComandoTX Lib "mp2032.dll" (ByVal BufTrans As String, ByVal TamBufTrans As Integer) As Integer
Public Declare Function ConfiguraModeloImpressora Lib "mp2032.dll" (ByVal ModeloImpressora As Integer) As Integer
Public Declare Function ConfiguraTamanhoExtrato Lib "mp2032.dll" (ByVal NumeroLinhas As Integer) As Integer
Public Declare Function DocumentInserted Lib "mp2032.dll" () As Integer
Public Declare Function EsperaImpressao Lib "mp2032.dll" () As Integer
Public Declare Function FechaPorta Lib "mp2032.dll" () As Integer
Public Declare Function FormataTX Lib "mp2032.dll" (ByVal BufTras As String, ByVal TpoLtra As Integer, ByVal Italic As Integer, ByVal Sublin As Integer, ByVal expand As Integer, ByVal enfat As Integer) As Integer
Public Declare Function HabilitaEsperaImpressao Lib "mp2032.dll" (ByVal Flag As Integer) As Integer
Public Declare Function HabilitaExtratoLongo Lib "mp2032.dll" (ByVal Flag As Integer) As Integer
Public Declare Function HabilitaPresenterRetratil Lib "mp2032.dll" (ByVal Flag As Integer) As Integer
Public Declare Function IniciaPorta Lib "mp2032.dll" (ByVal iPorta As String) As Integer
Public Declare Function Le_Status Lib "mp2032.dll" () As Integer
Public Declare Function Le_Status_Gaveta Lib "mp2032.dll" () As Integer
Public Declare Function ProgramaPresenterRetratil Lib "mp2032.dll" (ByVal Tempo As Integer) As Integer
Public Declare Function Status_Porta Lib "mp2032.dll" () As Integer
Public Declare Function VerificaPapelPresenter Lib "mp2032.dll" () As Integer
'funçõo para configuração dos códigos de barras
Public Declare Function ConfiguraCodigoBarras Lib "mp2032.dll" (ByVal Altura As Integer, ByVal Largura As Integer, ByVal PosicaoCaracteres As Integer, ByVal Fonte As Integer, ByVal Margem As Integer) As Integer


'funções para impressão do bitmap
Public Declare Function ImprimeBmpEspecial Lib "mp2032.dll" (ByVal FileName As String, ByVal xScale As Integer, _
                                                            ByVal yScale As Integer, ByVal angle As Integer) As Integer
                                                            
                                                            
Public Declare Function ImprimeBitmap Lib "mp2032.dll" (ByVal FileName As String, ByVal mode As Integer) As Integer

Public Declare Function AjustaLarguraPapel Lib "mp2032.dll" (ByVal width As Integer) As Integer
Public Declare Function SelectDithering Lib "mp2032.dll" (ByVal algorithm As Integer) As Integer


'funções para impressão dos códigos de barras
Public Declare Function ImprimeCodigoBarrasUPCA Lib "mp2032.dll" (ByVal codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasUPCE Lib "mp2032.dll" (ByVal codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasEAN13 Lib "mp2032.dll" (ByVal codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasEAN8 Lib "mp2032.dll" (ByVal codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasCODE39 Lib "mp2032.dll" (ByVal codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasCODE93 Lib "mp2032.dll" (ByVal codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasCODE128 Lib "mp2032.dll" (ByVal codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasITF Lib "mp2032.dll" (ByVal codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasCODABAR Lib "mp2032.dll" (ByVal codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasISBN Lib "mp2032.dll" (ByVal codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasMSI Lib "mp2032.dll" (ByVal codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasPLESSEY Lib "mp2032.dll" (ByVal codigo As String) As Integer
Public Declare Function ImprimeCodigoBarrasPDF417 Lib "mp2032.dll" (ByVal NivelCorrecaoErros As Integer, ByVal Altura As Integer, ByVal Largura As Integer, ByVal Colunas As Integer, ByVal codigo As String) As Integer

Private Function strCentralizado() As String
    strCentralizado = Chr$(27) & Chr$(97) & Chr$(1)
End Function

Private Function strEsquerda() As String
    strEsquerda = Chr$(27) & Chr$(97) & Chr$(0)
End Function

Private Function strIniImpress() As String
    strIniImpress = Chr$(27) & Chr$(64)
End Function

Public Function iniImpressoraBematech() As Boolean
    iniImpressoraBematech = False
    ComandoTX Chr$(29) & Chr$(249) & Chr$(32) & Chr$(0), Len(Chr$(29) & Chr$(249) & Chr$(32) & Chr$(0))
    If ComandoTX(strIniImpress, Len(strIniImpress)) <= 0 Then
        Exit Function
    End If

    iniImpressoraBematech = True
End Function

Public Function centralizadoBematech() As Boolean
    centralizadoBematech = False

    If ComandoTX(strCentralizado, Len(strCentralizado)) <= 0 Then
        Exit Function
    End If

    centralizadoBematech = True
End Function

Public Function strCentralizadoBematech() As String
    strCentralizadoBematech = strCentralizado
End Function

Public Function esquerdaBematech() As Boolean
    esquerdaBematech = False

    If ComandoTX(strEsquerda, Len(strEsquerda)) <= 0 Then
        Exit Function
    End If

    esquerdaBematech = True
End Function

Public Function strEsquerdaBematech() As String
    strEsquerdaBematech = strEsquerda
End Function

Public Function abreImpBematech(porta As String) As Boolean
        
    abreImpBematech = False
    If ConfiguraModeloImpressora(5) <= 0 Then 'modelos MP-20 TH, MP-2000 CI, MP-2000 TH e Blocos Impressores
        Exit Function
    End If
    
    If IniciaPorta(porta) <= 0 Then
        Exit Function
    End If
    
    If AjustaLarguraPapel(80) <= 0 Then
        Exit Function
    End If
    
    If ComandoTX(Chr$(27) & "t" & Chr$(2), Len(Chr$(27) & "t" & Chr$(2))) <= 0 Then
        Exit Function
    End If
    
    Call HabilitaEsperaImpressao(1)
    
    abreImpBematech = True
End Function

Public Function imprimeTracoBematech() As Boolean
    imprimeTracoBematech = imprimeNormalBematech(String(48, "_") & Chr$(10))
End Function

Public Function strTracoBematech() As String
    strTracoBematech = String(48, "_") & Chr$(10)
End Function

Public Function fechaImpBematech() As Boolean
    fechaImpBematech = False
    
    If FechaPorta() <= 0 Then
        Exit Function
    End If
 
    fechaImpBematech = True
End Function

Public Function cortaPapelBematech() As Boolean
    cortaPapelBematech = False
    
    If AcionaGuilhotina(0) <= 0 Then
        Exit Function
    End If
 
    cortaPapelBematech = True
End Function

Public Function strCortaPapelBematech() As String
    strCortaPapelBematech = Chr$(27) & "m"
End Function

Public Function imprimeComprimidoBematech(texto As String) As Boolean
    imprimeComprimidoBematech = False

    If ComandoTX(Chr$(27) & Chr$(80), Len(Chr$(27) & Chr$(80))) <= 0 Then
        Exit Function
    End If
    
    If FormataTX(texto, 1, 0, 0, 0, 0) <= 0 Then
        Exit Function
    End If

    imprimeComprimidoBematech = True
End Function

Public Function strComprimidoBematech(texto As String) As String
    strComprimidoBematech = Chr$(27) & Chr$(15)                                   'Ativa comprimido
    strComprimidoBematech = strComprimidoBematech & Chr$(27) & Chr$(53)           'Desativa italico
    strComprimidoBematech = strComprimidoBematech & Chr$(27) & Chr$(45) & Chr$(0) 'Desativa sublinhado
    strComprimidoBematech = strComprimidoBematech & Chr$(27) & Chr$(87) & Chr$(0) 'Desativa expandido
    strComprimidoBematech = strComprimidoBematech & Chr$(27) & Chr$(70)           'Desativa negrito
    strComprimidoBematech = strComprimidoBematech & texto
End Function

Public Function imprimeNormalBematech(texto As String) As Boolean
    imprimeNormalBematech = False

    If ComandoTX(Chr$(27) & Chr$(80), Len(Chr$(27) & Chr$(80))) <= 0 Then
        Exit Function
    End If
    
    If FormataTX(texto, 2, 0, 0, 0, 1) <= 0 Then
        Exit Function
    End If

    imprimeNormalBematech = True
End Function

Public Function strNormalBematech(texto As String) As String
    strNormalBematech = Chr$(27) & Chr$(80)                               'Ativa normal
    strNormalBematech = strNormalBematech & Chr$(27) & Chr$(53)           'Desativa italico
    strNormalBematech = strNormalBematech & Chr$(27) & Chr$(45) & Chr$(0) 'Desativa sublinhado
    strNormalBematech = strNormalBematech & Chr$(27) & Chr$(87) & Chr$(0) 'Desativa expandido
    strNormalBematech = strNormalBematech & Chr$(27) & Chr$(69)           'Desativa negrito
    strNormalBematech = strNormalBematech & texto
End Function

Public Function imprimeBematech(texto As String) As Boolean
    imprimeBematech = False

    ComandoTX Chr$(29) & Chr$(249) & Chr$(32) & Chr$(0), Len(Chr$(29) & Chr$(249) & Chr$(32) & Chr$(0))
    If ComandoTX(texto, Len(texto)) <= 0 Then
        Exit Function
    End If

    imprimeBematech = True
End Function

Public Function imprimeNormalReversoBematech(texto As String) As Boolean
    imprimeNormalReversoBematech = False
    
    If ComandoTX(Chr$(27) & Chr$(80), Len(Chr$(27) & Chr$(80))) <= 0 Then
        Exit Function
    End If
    
    If FormataTX(texto, 2, 0, 0, 0, 1) <= 0 Then
        Exit Function
    End If

    imprimeNormalReversoBematech = True
End Function

Public Function imprimeDuplaBematech(texto As String) As Boolean
    imprimeDuplaBematech = False

    If ComandoTX(Chr$(27) & Chr$(80), Len(Chr$(27) & Chr$(80))) <= 0 Then
        Exit Function
    End If

    If FormataTX(texto, 2, 0, 0, 1, 1) <= 0 Then
        Exit Function
    End If

    imprimeDuplaBematech = True
End Function

Public Function strDuplaBematech(texto As String) As String
    strDuplaBematech = Chr$(27) & Chr$(80)                                   'Ativa normal
    strDuplaBematech = strDuplaBematech & Chr$(27) & Chr$(53)           'Desativa italico
    strDuplaBematech = strDuplaBematech & Chr$(27) & Chr$(45) & Chr$(0) 'Desativa sublinhado
    strDuplaBematech = strDuplaBematech & Chr$(27) & Chr$(87) & Chr$(1) 'Ativa expandido
    strDuplaBematech = strDuplaBematech & Chr$(27) & Chr$(69)           'Ativa negrito
    strDuplaBematech = strDuplaBematech & texto
End Function

Public Function imprimeDuplaReversoBematech(texto As String) As Boolean
    imprimeDuplaReversoBematech = False
    
    If ComandoTX(Chr$(27) & Chr$(86) & Chr$(0), Len(Chr$(27) & Chr$(87) & Chr$(0))) <= 0 Then
        Exit Function
    End If
    
    If FormataTX(texto, 1, 0, 0, 1, 1) <= 0 Then
        Exit Function
    End If

    imprimeDuplaReversoBematech = True
End Function

Public Function imprimeTituloBematech(texto As String) As Boolean
    imprimeTituloBematech = False

    If ComandoTX(Chr$(27) & Chr$(80), Len(Chr$(27) & Chr$(80))) <= 0 Then
        Exit Function
    End If

    If FormataTX(texto, 2, 0, 0, 1, 0) <= 0 Then
        Exit Function
    End If

    imprimeTituloBematech = True
End Function

Public Function strTituloBematech(texto As String) As String
    'strTituloBematech = Chr$(27) & Chr$(80)                                   'Ativa normal
    'strTituloBematech = strTituloBematech & Chr$(27) & Chr$(53)           'Desativa italico
    'strTituloBematech = strTituloBematech & Chr$(27) & Chr$(45) & Chr$(0) 'Desativa sublinhado
    'strTituloBematech = strTituloBematech & Chr$(27) & Chr$(87) & Chr$(1) 'Ativa expandido
    'strTituloBematech = strTituloBematech & Chr$(27) & Chr$(70)           'Desativa negrito
    'strTituloBematech = strTituloBematech & texto
    strTituloBematech = Chr(27) & "!" & Chr(56) & texto & Chr(27) & "!" & Chr(0)
End Function

Public Function imprimeTituloReversoBematech(texto As String) As Boolean
    
    imprimeTituloReversoBematech = False
    
    If ComandoTX(Chr$(27) & Chr$(80), Len(Chr$(27) & Chr$(80))) <= 0 Then
        Exit Function
    End If

    If FormataTX(texto, 1, 0, 0, 1, 1) <= 0 Then
        Exit Function
    End If

    imprimeTituloReversoBematech = True
End Function

Public Function strTituloReversoBematech(texto As String) As String
    'strTituloReversoBematech = Chr$(27) & Chr$(15)                                      'Ativa compromido
    'strTituloReversoBematech = strTituloReversoBematech & Chr$(27) & Chr$(53)           'Desativa italico
    'strTituloReversoBematech = strTituloReversoBematech & Chr$(27) & Chr$(45) & Chr$(0) 'Desativa sublinhado
    'strTituloReversoBematech = strTituloReversoBematech & Chr$(27) & Chr$(87) & Chr$(1) 'Ativa expandido
    'strTituloReversoBematech = strTituloReversoBematech & Chr$(27) & Chr$(69)           'Ativa negrito
    'strTituloReversoBematech = strTituloReversoBematech & texto
    
    strTituloReversoBematech = Chr$(29) & Chr(66) & Chr(1) & Chr(27) & "!" & Chr(56) & texto & Chr$(29) & Chr(66) & Chr(0) & Chr(27) & "!" & Chr(0)
End Function

Public Function imprimeCodigoBarrasBematech(codigo As String) As Boolean
    imprimeCodigoBarrasBematech = False

    If ConfiguraCodigoBarras(90, 0, 0, 0, 30) <= 0 Then
        Exit Function
    End If
    
    If ImprimeCodigoBarrasCODE39(codigo) <= 0 Then
        Exit Function
    End If
    
    If Not centralizadoBematech Then
        Exit Function
    End If
    
    If Not imprimeCodigoBematech(codigo) Then
        Exit Function
    End If
    
    If Not imprimeNormalBematech(Chr$(10)) Then
        Exit Function
    End If

    
    imprimeCodigoBarrasBematech = True
End Function

Public Function strCodigoBarrasBematech(codigo As String) As String
    strCodigoBarrasBematech = ""
    
    'Configura o codigo de barras
    strCodigoBarrasBematech = Chr$(29) & Chr$(104) & Chr$(90)                                     'Define a altura para 90
    strCodigoBarrasBematech = strCodigoBarrasBematech & Chr$(29) & Chr$(119) & Chr$(2)            'Define a largura das barras
    strCodigoBarrasBematech = strCodigoBarrasBematech & Chr$(29) & Chr$(72) & Chr$(0)             'Define a posição do texto
    strCodigoBarrasBematech = strCodigoBarrasBematech & Chr$(29) & Chr$(102) & Chr$(0)            'Define a fonte do texto
    strCodigoBarrasBematech = strCodigoBarrasBematech & Chr$(29) & Chr$(107) & Chr$(132) & Chr$(30) & Chr$(0) 'Define a margem para 30
    
    'Imprime o codigo de barras
    strCodigoBarrasBematech = strCodigoBarrasBematech & Chr$(29) & Chr$(107) & Chr$(4) & codigo & Chr$(0)
    
    'Imprime o texto
    strCodigoBarrasBematech = strCentralizadoBematech & strCodigoBarrasBematech & strCentralizadoBematech & _
                              strCodigoBematech(codigo) & Chr$(10)
End Function

Public Function imprimeCodigoBematech(codigo As String) As Boolean

    imprimeCodigoBematech = imprimeNormalBematech("* " & Format(codigo, "@ @ @ @ @ @ @ @ @ @ @ @") & " *")

End Function

Public Function strCodigoBematech(codigo As String) As String

    strCodigoBematech = strNormalBematech("* " & Format(codigo, "@ @ @ @ @ @ @ @ @ @ @ @") & " *")

End Function


Public Function verificaImpBematech() As Boolean
    verificaImpBematech = False

    If VerificaPapelPresenter() <= 0 Then
        Exit Function
    End If

    verificaImpBematech = True
End Function


Public Function esperaImpressBematech() As Boolean
    esperaImpressBematech = True
    
    Call EsperaImpressao

End Function

