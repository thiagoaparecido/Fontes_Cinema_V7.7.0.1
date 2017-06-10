Attribute VB_Name = "modImpressora"
Option Explicit

Public Function iniImpressora(impEpson As OPOSPOSPrinter) As Boolean
    iniImpressora = False

    If UCase(pImpressora) = "EPSON" Then
        iniImpressora = iniImpressoraEpson(impEpson)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        iniImpressora = iniImpressoraBematech()
    Else
        Exit Function
    End If
End Function

Public Function cortaPapel(impEpson As OPOSPOSPrinter) As Boolean
    cortaPapel = False

    If UCase(pImpressora) = "EPSON" Then
        cortaPapel = cortaPapelEpson(impEpson)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        cortaPapel = cortaPapelBematech()
    Else
        Exit Function
    End If
End Function

Public Function strCortaPapel() As String
    strCortaPapel = ""

    If UCase(pImpressora) = "EPSON" Then
        strCortaPapel = strCortaPapelEpson
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        strCortaPapel = strCortaPapelBematech()
    Else
        Exit Function
    End If
End Function

Public Function AlinhaCentralizado(impEpson As OPOSPOSPrinter) As Boolean
    AlinhaCentralizado = False

    If UCase(pImpressora) = "EPSON" Then
        AlinhaCentralizado = centralizadoEpson(impEpson)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        AlinhaCentralizado = centralizadoBematech()
    Else
        Exit Function
    End If
End Function

Public Function strAlinhaCentralizado() As String
    strAlinhaCentralizado = ""

    If UCase(pImpressora) = "EPSON" Then
        strAlinhaCentralizado = strCentralizadoEpson()
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        strAlinhaCentralizado = strCentralizadoBematech()
    Else
        Exit Function
    End If
End Function

Public Function AlinhaEsquerda(impEpson As OPOSPOSPrinter) As Boolean
    AlinhaEsquerda = False

    If UCase(pImpressora) = "EPSON" Then
        AlinhaEsquerda = esquerdaEpson(impEpson)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        AlinhaEsquerda = esquerdaBematech()
    Else
        Exit Function
    End If
End Function

Public Function strAlinhaEsquerda() As String
    strAlinhaEsquerda = ""

    If UCase(pImpressora) = "EPSON" Then
        strAlinhaEsquerda = strEsquerdaEpson
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        strAlinhaEsquerda = strEsquerdaBematech()
    Else
        Exit Function
    End If
End Function

Public Function imprimeTraco(impEpson As OPOSPOSPrinter) As Boolean
    imprimeTraco = False

    If UCase(pImpressora) = "EPSON" Then
        imprimeTraco = imprimeTracoEpson(impEpson)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        imprimeTraco = imprimeTracoBematech()
    Else
        Exit Function
    End If
End Function

Public Function strTraco() As String
    strTraco = ""

    If UCase(pImpressora) = "EPSON" Then
        strTraco = strTracoEpson()
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        strTraco = strTracoBematech
    Else
        Exit Function
    End If
End Function

Public Function imprimeComprimido(impEpson As OPOSPOSPrinter, texto As String) As Boolean
    imprimeComprimido = False

    If UCase(pImpressora) = "EPSON" Then
        imprimeComprimido = imprimeComprimidoEpson(impEpson, texto)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        imprimeComprimido = imprimeComprimidoBematech(texto)
    Else
        Exit Function
    End If
End Function

Public Function strComprimido(texto As String) As String
    strComprimido = texto

    If UCase(pImpressora) = "EPSON" Then
        strComprimido = strComprimidoEpson(texto)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        strComprimido = strComprimidoBematech(texto)
    Else
        Exit Function
    End If
End Function

Public Function imprimeNormal(impEpson As OPOSPOSPrinter, texto As String) As Boolean
    imprimeNormal = False

    If UCase(pImpressora) = "EPSON" Then
        imprimeNormal = imprimeNormalEpson(impEpson, texto)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        imprimeNormal = imprimeNormalBematech(texto)
    Else
        Exit Function
    End If
End Function

Public Function imprime(impEpson As OPOSPOSPrinter, texto As String) As Boolean
    imprime = False

    If UCase(pImpressora) = "EPSON" Then
        imprime = imprimeEpson(impEpson, texto)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        imprime = imprimeBematech(texto)
    Else
        Exit Function
    End If
End Function

Public Function strNormal(texto As String) As String
    strNormal = texto

    If UCase(pImpressora) = "EPSON" Then
        strNormal = strNormalEpson(texto)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        strNormal = strNormalBematech(texto)
    Else
        Exit Function
    End If
End Function

Public Function imprimeNormalReverso(impEpson As OPOSPOSPrinter, texto As String) As Boolean
    imprimeNormalReverso = False
    
    If UCase(pImpressora) = "EPSON" Then
        imprimeNormalReverso = imprimeNormalReversoEpson(impEpson, texto)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        imprimeNormalReverso = imprimeNormalReversoBematech(texto)
    Else
        Exit Function
    End If
End Function

Public Function imprimeTitulo(impEpson As OPOSPOSPrinter, texto As String) As Boolean
    imprimeTitulo = False

    If UCase(pImpressora) = "EPSON" Then
        imprimeTitulo = imprimeTituloEpson(impEpson, texto)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        imprimeTitulo = imprimeTituloBematech(texto)
    Else
        Exit Function
    End If
End Function

Public Function strTitulo(texto As String) As String
    strTitulo = texto

    If UCase(pImpressora) = "EPSON" Then
        strTitulo = strTituloEpson(texto)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        strTitulo = strTituloBematech(texto)
    Else
        Exit Function
    End If
End Function

Public Function imprimeTituloReverso(impEpson As OPOSPOSPrinter, texto As String) As Boolean
    imprimeTituloReverso = False
    
    If UCase(pImpressora) = "EPSON" Then
        imprimeTituloReverso = imprimeTituloReversoEpson(impEpson, texto)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        imprimeTituloReverso = imprimeTituloReversoBematech(texto)
    Else
        Exit Function
    End If
End Function

Public Function strTituloReverso(texto As String) As String
    strTituloReverso = texto
    
    If UCase(pImpressora) = "EPSON" Then
        strTituloReverso = strTituloReversoEpson(texto)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        strTituloReverso = strTituloReversoBematech(texto)
    Else
        Exit Function
    End If
End Function

Public Function imprimeDupla(impEpson As OPOSPOSPrinter, texto As String) As Boolean
    imprimeDupla = False

    If UCase(pImpressora) = "EPSON" Then
        imprimeDupla = imprimeDuplaEpson(impEpson, texto)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        imprimeDupla = imprimeDuplaBematech(texto)
    Else
        Exit Function
    End If
End Function

Public Function strDupla(texto As String) As String
    strDupla = texto

    If UCase(pImpressora) = "EPSON" Then
        strDupla = strDuplaEpson(texto)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        strDupla = strDuplaBematech(texto)
    Else
        Exit Function
    End If
End Function

Public Function imprimeDuplaReverso(impEpson As OPOSPOSPrinter, texto As String) As Boolean
    imprimeDuplaReverso = False
    
    If UCase(pImpressora) = "EPSON" Then
        imprimeDuplaReverso = imprimeDuplaReversoEpson(impEpson, texto)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        imprimeDuplaReverso = imprimeDuplaReversoBematech(texto)
    Else
        Exit Function
    End If
End Function

Public Function abreImp(ByRef impEpson As OPOSPOSPrinter) As Boolean
    abreImp = False
        
    If UCase(pImpressora) = "EPSON" Then
        abreImp = abreImpEpson(impEpson, pImpressEpsom)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        abreImp = abreImpBematech(pPortaBematech)
    Else
        Exit Function
    End If
End Function

Public Function fechaImp(ByRef impEpson As OPOSPOSPrinter) As Boolean
    fechaImp = False

    If UCase(pImpressora) = "EPSON" Then
        fechaImp = fechaImpEpson(impEpson)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        fechaImp = fechaImpBematech()
    Else
        Exit Function
    End If
End Function

Public Function imprimeCodigoBarras(impEpson As OPOSPOSPrinter, codigo As String) As Boolean
    imprimeCodigoBarras = False

    If UCase(pImpressora) = "EPSON" Then
        imprimeCodigoBarras = imprimeCodigoBarrasEpson(impEpson, codigo)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        imprimeCodigoBarras = imprimeCodigoBarrasBematech(codigo)
    Else
        Exit Function
    End If
End Function

Public Function strCodigoBarras(codigo As String) As String
    strCodigoBarras = codigo

    If UCase(pImpressora) = "EPSON" Then
        strCodigoBarras = strCodigoBarrasEpson(codigo)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        strCodigoBarras = strCodigoBarrasBematech(codigo)
    Else
        Exit Function
    End If
End Function

Public Function imprimeCodigo(impEpson As OPOSPOSPrinter, codigo As String) As Boolean
    imprimeCodigo = False

    If UCase(pImpressora) = "EPSON" Then
        imprimeCodigo = imprimeCodigoEpson(impEpson, codigo)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        imprimeCodigo = imprimeCodigoBematech(codigo)
    Else
        Exit Function
    End If
End Function

Public Function strCodigo(codigo As String) As String
    strCodigo = codigo

    If UCase(pImpressora) = "EPSON" Then
        strCodigo = strCodigoEpson(codigo)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        strCodigo = strCodigoBematech(codigo)
    Else
        Exit Function
    End If
End Function

Public Function verificaImp(impEpson As OPOSPOSPrinter) As Boolean
    verificaImp = False

    If UCase(pImpressora) = "EPSON" Then
        verificaImp = verificaImpEpson(impEpson)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        verificaImp = verificaImpBematech()
    Else
        Exit Function
    End If
End Function

Public Function esperaImpress(impEpson As OPOSPOSPrinter) As Boolean
    esperaImpress = False

    If UCase(pImpressora) = "EPSON" Then
        esperaImpress = esperaImpressEpson(impEpson)
    ElseIf UCase(pImpressora) = "BEMATECH" Then
        esperaImpress = esperaImpressBematech()
    Else
        Exit Function
    End If
End Function
