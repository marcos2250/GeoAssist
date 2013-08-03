Attribute VB_Name = "RelatoriosUtil"
Option Explicit


Public Sub RepoeDadosDocumento(ByRef ms As String)
Dim PJ As Boolean
Dim i As Integer
        
    If UCase(Left(GetParamData("CAMPO_TIPOPESSOA"), 1)) = "J" Then
        PJ = True
    Else
        PJ = False
    End If
    
    If PJ Then
        ms = Replace(ms, "CAMPO_VOC", "A")
        ms = Replace(ms, "CAMPO_CPFCNPJ", "CNPJ")
        ms = Replace(ms, "CAMPO_RGINSC", "INSC.EST.")
        ms = Replace(ms, "CAMPO_RGORGAO", "")
        ms = Replace(ms, "CAMPO_SITUADO", UTF8_Encode("situada à"))
    Else
        ms = Replace(ms, "CAMPO_VOC", "Eu")
        ms = Replace(ms, "CAMPO_CPFCNPJ", "CPF")
        ms = Replace(ms, "CAMPO_RGINSC", "RG")
        ms = Replace(ms, "CAMPO_SITUADO", UTF8_Encode("residente e domiciliado(a) à"))
    End If
    
    For i = 1 To NumParamsGeral
        If pImovel.Dados(i, 3) <> "" Then
            ms = Replace(ms, pImovel.Dados(i, 1), UTF8_Encode(pImovel.Dados(i, 3)))
        Else
            ms = Replace(ms, pImovel.Dados(i, 1), UTF8_Encode(String(20, "_")))
        End If
    Next i
    
End Sub


Public Sub GeraDocSimples(DocOrigem As String, DocDestino As String)
    If DocDestino = "" Then Exit Sub
    Dim ms As String
    ms = LoadFileToString(DocOrigem)
    RepoeDadosDocumento ms
    WriteStringToFile ms, DocDestino
End Sub




Public Sub GeraRelatorio(ctipo As Byte, ByVal savedir As String)
    Dim dx1 As Double, dy1 As Double
    Dim dx2 As Double, dy2 As Double
    Dim ms As String
    Dim i As Integer, k As Integer
    
    If savedir = "" Then Exit Sub
      
    ms = ms & "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fswiss\fprq2\fcharset0 Arial Narrow;}{\f1\fswiss\fcharset0 Arial;}}" & vbCrLf & "\viewkind4\uc1\pard\qc\f0\fs24" & vbCrLf & "\par"

    If ctipo = 1 Then
        'If ActiveWorkbook.ActiveSheet.Cells(3, 1).Value = "Proprietário:" Or ActiveWorkbook.ActiveSheet.Cells(3, 1).Value = "Proprietario:" Then
                ms = ms & "\fs28 \b M E M O R I A L    D E S C R I T I V O \b0" & vbCrLf & vbCrLf & _
                    "\par \par \pard \ql \fs24 IMÓVEL: " & GetParamData("CAMPO_NOMEIMOVEL") & vbCrLf & _
                    "\par \ql PROPRIETÁRIO: " & GetParamData("CAMPO_NOMEPROP") & vbCrLf & _
                    "\par \ql ÁREA REGISTRADA: " & GetParamData("CAMPO_AREAREG") & " ha " & vbCrLf & _
                    "\par \ql ÁREA MEDIDA: " & GetParamData("CAMPO_AREAMED") & " ha " & vbCrLf & _
                    "\par \ql PERÍMETRO: " & GetParamData("CAMPO_PERIMETRO") & " m " & vbCrLf & _
                    "\par \ql MATRÍCULA: " & GetParamData("CAMPO_MATRICULA") & " " & vbCrLf & _
                    "\par \ql MUNICÍPIO/UF DOCUMENTO: " & GetParamData("CAMPO_MUNICIPIODOC") & " " & vbCrLf & _
                    "\par \ql MUNICÍPIO/UF ATUAL: " & GetParamData("CAMPO_MUNICIPIOATUAL") & " " & vbCrLf & _
                    "\par \ql COMARCA: " & GetParamData("CAMPO_COMARCA") & " " & vbCrLf & _
                    "\par \ql CCIR: " & GetParamData("CAMPO_CCIR") & " " & vbCrLf & _
                    "\par \ql ART: " & GetParamData("CAMPO_ART") & " " & vbCrLf & _
                "\par \par \pard \fs28 \qc \b D E S C R I Ç Ã O   D O   P E R Í M E T R O   D O   I M Ó V E L \b0" & vbCrLf
        'Else
        '        ms = ms & "\fs28 \b D E S C R I Ç Ã O   D O   P E R Í M E T R O   D O   I M Ó V E L \b0" & vbCrLf & vbCrLf
        'End If
    End If
    
    If ctipo = 2 Then
        ms = ms & "\b Relatório de Confrontações \b0" & vbCrLf
    End If
       
    If ctipo = 3 Then
        ms = ms & "\b DECLARAÇÃO DE CONFRONTANTES \b0" & vbCrLf
    End If
    
    If ctipo = 4 Then
        ms = ms & "\b Relatório de Confrontações \b0" & vbCrLf
    End If

    'For k = TabelaPrincipal To (TabelaPrincipal + NumeroIlhas)
    
        If pImovel.NumCoords = 0 Then Exit Sub
        
        'If ctipo = 1 Then 'If ctipo = 1 And k > TabelaPrincipal Then
        '    ms = ms & vbCrLf & "\par \par \par \par \par \pard \fs28 \qc \b D E S C R I Ç Ã O   D O   P E R Í M E T R O   I N T E R N O  \par ( Á R E A   D E   E X C L U S Ã O ) \b0" & vbCrLf
        'End If
        
        ms = ms & "\par \par \pard \qj \fs24"
        If ctipo = 1 Then ms = ms & "Inicia-se a descrição deste perímetro no vértice "
        
        
        For i = 1 To pImovel.NumCoords
        
            If pImovel.dPoligono(i).Confrontante <> "" And (ctipo = 2 Or ctipo = 4) Then
                If ctipo = 2 Then
                    ms = ms & vbCrLf & "\par \par \par Confrontante: " & pImovel.dPoligono(i).Confrontante
                End If
                ms = ms & vbCrLf & "\par \par Do vértice "
            End If
            
            ms = ms & "\b " & pImovel.dPoligono(i).Marco & "\b0 , de coordenadas " & _
                      "\b N " & decToStr(pImovel.dPoligono(i).Y, 2) & _
                      "m \b0 e \b E " & decToStr(pImovel.dPoligono(i).X, 2) & "\b0 ; "
                      
            If pImovel.dPoligono(i).Confrontante <> "" Then
                If ctipo = 1 Then
                        ms = ms & " deste, segue confrontando com \b " & _
                                  pImovel.dPoligono(i).Confrontante & _
                                  " \b0 , com os seguintes azimutes e distâncias: "
                ElseIf ctipo = 2 Then
                        ms = ms & " segue confrontando com o supracitado confrontante, com os seguintes azimutes e distâncias: "
                
                ElseIf ctipo = 4 Then
                        ms = ms & " segue confrontando com o supracitado confrontante indicado no mapa e no memorial descritivo como " & pImovel.dPoligono(i).Confrontante & ", com os seguintes azimutes e distâncias: "
                End If
            End If
           
            dx1 = CDec(pImovel.dPoligono(i).X)
            dy1 = CDec(pImovel.dPoligono(i).Y)
            
            If i < pImovel.NumCoords Then
                dx2 = CDec(pImovel.dPoligono(i + 1).X)
                dy2 = CDec(pImovel.dPoligono(i + 1).Y)
            Else
                dx2 = CDec(pImovel.dPoligono(1).X)
                dy2 = CDec(pImovel.dPoligono(1).Y)
            End If
            
            'Calcula azimute e distancia
            ms = ms & dec2dms(AnguloCoord(dx1, dy1, dx2, dy2) / deg2rad, 0) & " e " _
                    & decToStr(Sqr((Abs(dy1 - dy2) ^ 2) + (Abs(dx1 - dx2) ^ 2)), 2) & _
                    "m até o vértice "
                    
                    
            If i < pImovel.NumCoords And (ctipo = 2 Or ctipo = 4) Then
                ms = ms & "\b " & pImovel.dPoligono(i + 1).Marco & "\b0 , de coordenadas " & _
                      "\b N " & decToStr(pImovel.dPoligono(i + 1).Y, 2) & _
                      "m \b0 e \b E " & decToStr(pImovel.dPoligono(i + 1).X, 2) & "\b0 ; "
            End If
            
        Next i
        
        If ctipo = 1 Then
            ms = ms & " \b " & pImovel.dPoligono(1).Marco & _
                " \b0 , ponto inicial da descrição deste perímetro. Todas as coordenadas aqui descritas estão georeferenciadas ao Sistema Geodésico Brasileiro, e encontram-se representadas no Sistema UTM, referenciadas ao Meridiano Central nº " & datum.MC & "°00', fuso -" & datum.Zona & ", tendo como datum o " & datum.Nome & ". Todos os azimutes e distâncias, área e perímetro foram calculados no plano de projeção UTM." & vbCrLf
        ElseIf (ctipo = 2 Or ctipo = 4) Then
            ms = ms & "\b " & pImovel.dPoligono(1).Marco & "\b0 , de coordenadas " & _
                "\b N " & decToStr(pImovel.dPoligono(1).Y, 2) & _
                "m \b0 e \b E " & decToStr(pImovel.dPoligono(1).X, 2) & "\b0 ; "
        End If
   'Next k
    
    If ctipo = 1 Then
            ms = ms & "\par \par \par \pard \qc \fs20 \b ___________________________________________________________" & vbCrLf & _
             "\par " & GetParamData("CAMPO_TECNICO") & vbCrLf & _
             "\par \par \par \pard \ql \fs20 Código de certificação: " & GetParamData("CAMPO_MARCOINCRA") & vbCrLf & _
             "\par CREA " & GetParamData("CAMPO_CREA") & vbCrLf & _
             "\par Visto " & GetParamData("CAMPO_VISTO") & vbCrLf & _
             "\par ART: \b0" & GetParamData("CAMPO_ART") & vbCrLf
    End If
        
    ms = ms & "\par\par" & vbCrLf & "}"
   
    WriteStringToFile ms, savedir
End Sub


Public Sub GeraMemoDescrIncra(savedir As String)
    Dim i As Integer, k As Integer
    Dim dx1 As Double, dx2 As Double, dy1 As Double, dy2 As Double
    Dim ms As String, mr As String
    Dim auxStr As String
        
    If savedir = "" Then Exit Sub
    If pImovel.NumCoords = 0 Then Exit Sub
              
        
    'For k = TabelaPrincipal To (TabelaPrincipal + NumeroIlhas)
        
    For i = 1 To pImovel.NumCoords
        dx1 = pImovel.dPoligono(i).X
        dy1 = pImovel.dPoligono(i).Y
            
        If i < pImovel.NumCoords Then
            dx2 = pImovel.dPoligono(i + 1).X
            dy2 = pImovel.dPoligono(i + 1).Y
        Else
            dx2 = pImovel.dPoligono(1).X
            dy2 = pImovel.dPoligono(1).Y
        End If
        
        
        If pImovel.dPoligono(i).Confrontante <> "" Or i = 1 Then
            auxStr = pImovel.dPoligono(i).Confrontante
        End If
            
        mr = mr & "<w:tr>" & vbCrLf
        mr = mr & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "0" & Chr(34) & " w:type=" & Chr(34) & "auto" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/></w:pPr><w:r><w:t>" & pImovel.dPoligono(i).Marco & "</w:t></w:r></w:p></w:tc>" & vbCrLf
        
        If i < pImovel.NumCoords Then
            mr = mr & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "0" & Chr(34) & " w:type=" & Chr(34) & "auto" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/></w:pPr><w:r><w:t>" & pImovel.dPoligono(i + 1).Marco & "</w:t></w:r></w:p></w:tc>" & vbCrLf
        Else
            mr = mr & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "0" & Chr(34) & " w:type=" & Chr(34) & "auto" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/></w:pPr><w:r><w:t>" & pImovel.dPoligono(1).Marco & "</w:t></w:r></w:p></w:tc>" & vbCrLf
        End If
        
        mr = mr & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "0" & Chr(34) & " w:type=" & Chr(34) & "auto" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/></w:pPr><w:r><w:t>" & decToStr(Sqr((Abs(dy1 - dy2) ^ 2) + (Abs(dx1 - dx2) ^ 2)), 2) & "m</w:t></w:r></w:p></w:tc>" & vbCrLf
        mr = mr & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "0" & Chr(34) & " w:type=" & Chr(34) & "auto" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/></w:pPr><w:r><w:t>" & dec2dms(AnguloCoord(dx1, dy1, dx2, dy2) / deg2rad, 3) & "</w:t></w:r></w:p></w:tc>" & vbCrLf
        mr = mr & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "0" & Chr(34) & " w:type=" & Chr(34) & "auto" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/></w:pPr><w:r><w:t>" & datum.MC & "</w:t></w:r></w:p></w:tc>" & vbCrLf
        mr = mr & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "0" & Chr(34) & " w:type=" & Chr(34) & "auto" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/></w:pPr><w:r><w:t>" & decToStr(pImovel.dPoligono(i).X, 2) & "</w:t></w:r></w:p></w:tc>" & vbCrLf
        mr = mr & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "0" & Chr(34) & " w:type=" & Chr(34) & "auto" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/></w:pPr><w:r><w:t>" & decToStr(pImovel.dPoligono(i).Y, 2) & "</w:t></w:r></w:p></w:tc>" & vbCrLf
        mr = mr & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "0" & Chr(34) & " w:type=" & Chr(34) & "auto" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/></w:pPr><w:r><w:t>" & auxStr & "</w:t></w:r></w:p></w:tc>" & vbCrLf
        mr = mr & "</w:tr>" & vbCrLf
        
    Next i
    
    'Next k
 
    ms = LoadFileToString(App.Path & "\@MemorialIncra.dat")
    ms = Replace(ms, "CAMPO_GRADEREPLACE", UTF8_Encode(mr))
    RepoeDadosDocumento ms
    WriteStringToFile ms, savedir
End Sub


Public Sub GeraTabelaDL(savedir As String)
    Dim i As Integer, k As Integer
    Dim dx1 As Double, dx2 As Double, dy1 As Double, dy2 As Double
    Dim ms As String
        
    If savedir = "" Then Exit Sub
    
    ms = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "yes" & Chr(34) & "?><?mso-application progid=" & Chr(34) & "Word.Document" & Chr(34) & "?>" & vbCrLf
    ms = ms & "<w:wordDocument xmlns:w=" & Chr(34) & "http://schemas.microsoft.com/office/word/2003/wordml" & Chr(34) & " xmlns:v=" & Chr(34) & "urn:schemas-microsoft-com:vml" & Chr(34) & " xmlns:w10=" & Chr(34) & "urn:schemas-microsoft-com:office:word" & Chr(34) & " xmlns:sl=" & Chr(34) & "http://schemas.microsoft.com/schemaLibrary/2003/core" & Chr(34) & " xmlns:aml=" & Chr(34) & "http://schemas.microsoft.com/aml/2001/core" & Chr(34) & " xmlns:wx=" & Chr(34) & "http://schemas.microsoft.com/office/word/2003/auxHint" & Chr(34) & " xmlns:o=" & Chr(34) & "urn:schemas-microsoft-com:office:office" & Chr(34) & " xmlns:dt=" & Chr(34) & "uuid:C2F41010-65B3-11d1-A29F-00AA00C14882" & Chr(34) & " w:macrosPresent=" & Chr(34) & "no" & Chr(34) & " w:embeddedObjPresent=" & Chr(34) & "no" & Chr(34) & " w:ocxPresent=" & Chr(34) & "no" & Chr(34) & " xml:space=" & Chr(34) & "preserve" & Chr(34) & ">" & vbCrLf
    ms = ms & "<w:docPr><w:view w:val=" & Chr(34) & "print" & Chr(34) & "/><w:zoom w:val=" & Chr(34) & "best-fit" & Chr(34) & " w:percent=" & Chr(34) & "148" & Chr(34) & "/></w:docPr><w:body><wx:sect>" & vbCrLf
    
    ms = ms & "<w:p><w:r><w:t>Relatório de Confrontantes (Tabela)</w:t></w:r></w:p><w:p/>" & vbCrLf
       
    'For k = TabelaPrincipal To (TabelaPrincipal + NumeroIlhas)

    If pImovel.NumCoords = 0 Then Exit Sub
        
    For i = 1 To pImovel.NumCoords
        dx1 = pImovel.dPoligono(i).X
        dy1 = pImovel.dPoligono(i).Y
            
        If i < pImovel.NumCoords Then
            dx2 = pImovel.dPoligono(i + 1).X
            dy2 = pImovel.dPoligono(i + 1).Y
        Else
            dx2 = pImovel.dPoligono(1).X
            dy2 = pImovel.dPoligono(1).Y
        End If
        
        
        If pImovel.dPoligono(i).Confrontante <> "" Or i = 1 Then
        
            If i <> 1 Then
                ms = ms & "</w:tbl>" & vbCrLf
            End If
            
            ms = ms & "<w:p/><w:p><w:r><w:t>Confrontante: " & pImovel.dPoligono(i).Confrontante & " </w:t></w:r></w:p><w:p/>" & vbCrLf
            
            ms = ms & "<w:p/><w:tbl><w:tblPr><w:tblW w:w=" & Chr(34) & "5000" & Chr(34) & " w:type=" & Chr(34) & "pct" & Chr(34) & "/>" & vbCrLf
            ms = ms & "<w:tblBorders><w:top w:val=" & Chr(34) & "single" & Chr(34) & " w:sz=" & Chr(34) & "6" & Chr(34) & " wx:bdrwidth=" & Chr(34) & "10" & Chr(34) & " w:space=" & Chr(34) & "0" & Chr(34) & " w:color=" & Chr(34) & "auto" & Chr(34) & "/>" & vbCrLf
            ms = ms & "<w:left w:val=" & Chr(34) & "single" & Chr(34) & " w:sz=" & Chr(34) & "4" & Chr(34) & " wx:bdrwidth=" & Chr(34) & "10" & Chr(34) & " w:space=" & Chr(34) & "0" & Chr(34) & " w:color=" & Chr(34) & "auto" & Chr(34) & "/>" & vbCrLf
            ms = ms & "<w:bottom w:val=" & Chr(34) & "single" & Chr(34) & " w:sz=" & Chr(34) & "4" & Chr(34) & " wx:bdrwidth=" & Chr(34) & "10" & Chr(34) & " w:space=" & Chr(34) & "0" & Chr(34) & " w:color=" & Chr(34) & "auto" & Chr(34) & "/>" & vbCrLf
            ms = ms & "<w:right w:val=" & Chr(34) & "single" & Chr(34) & " w:sz=" & Chr(34) & "4" & Chr(34) & " wx:bdrwidth=" & Chr(34) & "10" & Chr(34) & " w:space=" & Chr(34) & "0" & Chr(34) & " w:color=" & Chr(34) & "auto" & Chr(34) & "/>" & vbCrLf
            ms = ms & "<w:insideH w:val=" & Chr(34) & "single" & Chr(34) & " w:sz=" & Chr(34) & "4" & Chr(34) & " wx:bdrwidth=" & Chr(34) & "10" & Chr(34) & " w:space=" & Chr(34) & "0" & Chr(34) & " w:color=" & Chr(34) & "auto" & Chr(34) & "/>" & vbCrLf
            ms = ms & "<w:insideV w:val=" & Chr(34) & "single" & Chr(34) & " w:sz=" & Chr(34) & "4" & Chr(34) & " wx:bdrwidth=" & Chr(34) & "10" & Chr(34) & " w:space=" & Chr(34) & "0" & Chr(34) & " w:color=" & Chr(34) & "auto" & Chr(34) & "/>" & vbCrLf
            ms = ms & "</w:tblBorders><w:tblCellMar><w:left w:w=" & Chr(34) & "10" & Chr(34) & " w:type=" & Chr(34) & "dxa" & Chr(34) & "/><w:right w:w=" & Chr(34) & "10" & Chr(34) & " w:type=" & Chr(34) & "dxa" & Chr(34) & "/></w:tblCellMar><w:tblLook w:val=" & Chr(34) & "01E0" & Chr(34) & "/></w:tblPr>" & vbCrLf
    
            ms = ms & "<w:tr>" & vbCrLf
            ms = ms & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "1500" & Chr(34) & " w:type=" & Chr(34) & "dxa" & Chr(34) & "/><w:vAlign w:val=" & Chr(34) & "center" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>De</w:t></w:r></w:p></w:tc>" & vbCrLf
            ms = ms & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "0" & Chr(34) & " w:type=" & Chr(34) & "auto" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Para</w:t></w:r></w:p></w:tc>" & vbCrLf
            ms = ms & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "0" & Chr(34) & " w:type=" & Chr(34) & "auto" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Azimute</w:t></w:r></w:p></w:tc>" & vbCrLf
            ms = ms & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "0" & Chr(34) & " w:type=" & Chr(34) & "auto" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Distância</w:t></w:r></w:p></w:tc>" & vbCrLf
            ms = ms & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "0" & Chr(34) & " w:type=" & Chr(34) & "auto" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Este</w:t></w:r></w:p></w:tc>" & vbCrLf
            ms = ms & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "0" & Chr(34) & " w:type=" & Chr(34) & "auto" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Norte</w:t></w:r></w:p></w:tc>" & vbCrLf
            ms = ms & "</w:tr>" & vbCrLf
        End If
            
        ms = ms & "<w:tr>" & vbCrLf
        ms = ms & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "0" & Chr(34) & " w:type=" & Chr(34) & "auto" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/></w:pPr><w:r><w:t>" & pImovel.dPoligono(i).Marco & "</w:t></w:r></w:p></w:tc>" & vbCrLf
        
        If i < pImovel.NumCoords Then
            ms = ms & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "0" & Chr(34) & " w:type=" & Chr(34) & "auto" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/></w:pPr><w:r><w:t>" & pImovel.dPoligono(i + 1).Marco & "</w:t></w:r></w:p></w:tc>" & vbCrLf
        Else
            ms = ms & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "0" & Chr(34) & " w:type=" & Chr(34) & "auto" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/></w:pPr><w:r><w:t>" & pImovel.dPoligono(1).Marco & "</w:t></w:r></w:p></w:tc>" & vbCrLf
        End If
        
        ms = ms & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "0" & Chr(34) & " w:type=" & Chr(34) & "auto" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/></w:pPr><w:r><w:t>" & dec2dms(AnguloCoord(dx1, dy1, dx2, dy2) / deg2rad, 3) & "</w:t></w:r></w:p></w:tc>" & vbCrLf
        ms = ms & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "0" & Chr(34) & " w:type=" & Chr(34) & "auto" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/></w:pPr><w:r><w:t>" & decToStr(Sqr((Abs(dy1 - dy2) ^ 2) + (Abs(dx1 - dx2) ^ 2)), 2) & "m</w:t></w:r></w:p></w:tc>" & vbCrLf
        ms = ms & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "0" & Chr(34) & " w:type=" & Chr(34) & "auto" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/></w:pPr><w:r><w:t>" & decToStr(pImovel.dPoligono(i).X, 2) & "</w:t></w:r></w:p></w:tc>" & vbCrLf
        ms = ms & "<w:tc><w:tcPr><w:tcW w:w=" & Chr(34) & "0" & Chr(34) & " w:type=" & Chr(34) & "auto" & Chr(34) & "/></w:tcPr><w:p><w:pPr><w:jc w:val=" & Chr(34) & "center" & Chr(34) & "/></w:pPr><w:r><w:t>" & decToStr(pImovel.dPoligono(i).Y, 2) & "</w:t></w:r></w:p></w:tc>" & vbCrLf
        ms = ms & "</w:tr>" & vbCrLf
        
    Next i
    
        ms = ms & "</w:tbl>" & vbCrLf
    
    'Next k

    ms = ms & "<w:sectPr><w:pgSz w:w=" & Chr(34) & "11906" & Chr(34) & " w:h=" & Chr(34) & "16838" & Chr(34) & "/>"
    ms = ms & "<w:pgMar w:top=" & Chr(34) & "1417" & Chr(34) & " w:right=" & Chr(34) & "1701" & Chr(34) & " w:bottom=" & Chr(34) & "1417" & Chr(34) & " w:left=" & Chr(34) & "1701" & Chr(34) & " w:header=" & Chr(34) & "720" & Chr(34) & " w:footer=" & Chr(34) & "720" & Chr(34) & " w:gutter=" & Chr(34) & "0" & Chr(34) & "/>"
    ms = ms & "<w:cols w:space=" & Chr(34) & "720" & Chr(34) & "/><w:docGrid w:line-pitch=" & Chr(34) & "360" & Chr(34) & "/>"
    ms = ms & "</w:sectPr></wx:sect></w:body></w:wordDocument>"
 
    WriteStringToFile UTF8_Encode(ms), savedir
End Sub



Public Sub CalculoAnalitico(savedir As String, tipo As Byte)
Dim ms As String, mr As String
Dim i As Integer, k As Integer
Dim dx1 As Double, dx2 As Double, dy1 As Double, dy2 As Double
Dim nx As Double, ny As Double
Dim auxStr As String
        
    If savedir = "" Then Exit Sub
    
    ms = ""
    mr = ""
    
    For i = 1 To pImovel.NumCoords
        dx1 = pImovel.dPoligono(i).X
        dy1 = pImovel.dPoligono(i).Y
            
        mr = mr & "   <Row>" & vbCrLf
        
        If tipo = 1 Then
            If i < pImovel.NumCoords Then
                dx2 = pImovel.dPoligono(i + 1).X
                dy2 = pImovel.dPoligono(i + 1).Y
            Else
                dx2 = pImovel.dPoligono(1).X
                dy2 = pImovel.dPoligono(1).Y
            End If
            
            mr = mr & "    <Cell ss:StyleID=" & Chr$(34) & "s23" & Chr$(34) & "><Data ss:Type=" & Chr$(34) & "String" & Chr$(34) & ">" & pImovel.dPoligono(i).Marco & "</Data></Cell>" & vbCrLf
            
            If i < pImovel.NumCoords Then
                mr = mr & "    <Cell ss:StyleID=" & Chr$(34) & "s23" & Chr$(34) & "><Data ss:Type=" & Chr$(34) & "String" & Chr$(34) & ">" & pImovel.dPoligono(i + 1).Marco & "</Data></Cell>" & vbCrLf
            Else
                mr = mr & "    <Cell ss:StyleID=" & Chr$(34) & "s23" & Chr$(34) & "><Data ss:Type=" & Chr$(34) & "String" & Chr$(34) & ">" & pImovel.dPoligono(1).Marco & "</Data></Cell>" & vbCrLf
            End If
            
            UTM2LL nx, ny, (pImovel.dPoligono(i).X), (pImovel.dPoligono(i).Y)
            mr = mr & "    <Cell ss:StyleID=" & Chr$(34) & "s25" & Chr$(34) & "><Data ss:Type=" & Chr$(34) & "String" & Chr$(34) & ">" & dec2dms(AnguloCoord(dx1, dy1, dx2, dy2) / deg2rad, 3) & "</Data></Cell>" & vbCrLf
            mr = mr & "    <Cell ss:StyleID=" & Chr$(34) & "s25" & Chr$(34) & "><Data ss:Type=" & Chr$(34) & "String" & Chr$(34) & ">" & decToStr(Sqr((Abs(dy1 - dy2) ^ 2) + (Abs(dx1 - dx2) ^ 2)), 2) & "m</Data></Cell>" & vbCrLf
            mr = mr & "    <Cell ss:StyleID=" & Chr$(34) & "s25" & Chr$(34) & "><Data ss:Type=" & Chr$(34) & "String" & Chr$(34) & ">" & decToStr(pImovel.dPoligono(i).X, 2) & "</Data></Cell>" & vbCrLf
            mr = mr & "    <Cell ss:StyleID=" & Chr$(34) & "s25" & Chr$(34) & "><Data ss:Type=" & Chr$(34) & "String" & Chr$(34) & ">" & decToStr(pImovel.dPoligono(i).Y, 2) & "</Data></Cell>" & vbCrLf
            mr = mr & "    <Cell ss:StyleID=" & Chr$(34) & "s25" & Chr$(34) & "><Data ss:Type=" & Chr$(34) & "String" & Chr$(34) & ">" & dec2dms(nx, 1) & "</Data></Cell>" & vbCrLf
            mr = mr & "    <Cell ss:StyleID=" & Chr$(34) & "s25" & Chr$(34) & "><Data ss:Type=" & Chr$(34) & "String" & Chr$(34) & ">" & dec2dms(ny, 2) & "</Data></Cell>" & vbCrLf
            mr = mr & "    <Cell ss:StyleID=" & Chr$(34) & "s25" & Chr$(34) & "><Data ss:Type=" & Chr$(34) & "String" & Chr$(34) & ">" & decToStr(FatorK_UTM(pImovel.dPoligono(i).X, pImovel.dPoligono(i).Y), 6) & "</Data></Cell>" & vbCrLf
        Else
        
            If pImovel.dPoligono(i).Confrontante <> "" Then auxStr = pImovel.dPoligono(i).Confrontante
        
            mr = mr & "    <Cell ss:StyleID=" & Chr$(34) & "s36" & Chr$(34) & "><Data ss:Type=" & Chr$(34) & "String" & Chr$(34) & ">" & pImovel.dPoligono(i).Marco & "</Data></Cell>" & vbCrLf
            mr = mr & "    <Cell ss:StyleID=" & Chr$(34) & "s36" & Chr$(34) & "><Data ss:Type=" & Chr$(34) & "String" & Chr$(34) & ">" & CStr(i) & "</Data></Cell>" & vbCrLf
            mr = mr & "    <Cell ss:StyleID=" & Chr$(34) & "s36" & Chr$(34) & "><Data ss:Type=" & Chr$(34) & "String" & Chr$(34) & ">" & decToStr(pImovel.dPoligono(i).X, 2) & "</Data></Cell>" & vbCrLf
            mr = mr & "    <Cell ss:StyleID=" & Chr$(34) & "s36" & Chr$(34) & "><Data ss:Type=" & Chr$(34) & "String" & Chr$(34) & ">" & decToStr(pImovel.dPoligono(i).dE, 3) & "</Data></Cell>" & vbCrLf
            mr = mr & "    <Cell ss:StyleID=" & Chr$(34) & "s36" & Chr$(34) & "><Data ss:Type=" & Chr$(34) & "String" & Chr$(34) & ">" & decToStr(pImovel.dPoligono(i).Y, 2) & "</Data></Cell>" & vbCrLf
            mr = mr & "    <Cell ss:StyleID=" & Chr$(34) & "s36" & Chr$(34) & "><Data ss:Type=" & Chr$(34) & "String" & Chr$(34) & ">" & decToStr(pImovel.dPoligono(i).dN, 3) & "</Data></Cell>" & vbCrLf
            mr = mr & "    <Cell ss:StyleID=" & Chr$(34) & "s36" & Chr$(34) & "><Data ss:Type=" & Chr$(34) & "String" & Chr$(34) & ">" & decToStr(pImovel.dPoligono(i).Z, 2) & "</Data></Cell>" & vbCrLf
            mr = mr & "    <Cell ss:StyleID=" & Chr$(34) & "s36" & Chr$(34) & "><Data ss:Type=" & Chr$(34) & "String" & Chr$(34) & ">" & decToStr(pImovel.dPoligono(i).dH, 3) & "</Data></Cell>" & vbCrLf
            mr = mr & "    <Cell ss:StyleID=" & Chr$(34) & "s36" & Chr$(34) & "><Data ss:Type=" & Chr$(34) & "String" & Chr$(34) & ">" & AptDefs(getAptDefIndex(CStr(pImovel.dPoligono(i).MetodoLevantamento))).CodigoIncra & "</Data></Cell>" & vbCrLf
            mr = mr & "    <Cell ss:StyleID=" & Chr$(34) & "s36" & Chr$(34) & "><Data ss:Type=" & Chr$(34) & "String" & Chr$(34) & ">" & AptDefs(getAptDefIndex(CStr(pImovel.dPoligono(i).TipoLimite))).CodigoIncra & "</Data></Cell>" & vbCrLf
            mr = mr & "    <Cell ss:StyleID=" & Chr$(34) & "s36" & Chr$(34) & "><Data ss:Type=" & Chr$(34) & "String" & Chr$(34) & ">" & auxStr & "</Data></Cell>" & vbCrLf
        
        End If
        
        mr = mr & "   </Row>" & vbCrLf
        
    Next i
    
    If tipo = 1 Then
        ms = LoadFileToString(App.Path & "\@Grade.dat")
    Else
        ms = LoadFileToString(App.Path & "\@Resultados.dat")
    End If
    
    ms = Replace(ms, "CAMPO_GRADENUMLINHAS", CStr(pImovel.NumCoords + 4))
    ms = Replace(ms, "CAMPO_GRADEREPLACE", UTF8_Encode(mr))
    
    WriteStringToFile ms, savedir
End Sub


Public Sub ArquivoResultado(savedir As String)
Dim i As Integer, auxStr As String, ms As String
    
        ms = ""
        For i = 1 To pImovel.NumCoords
            If pImovel.dPoligono(i).Confrontante <> "" Then auxStr = pImovel.dPoligono(i).Confrontante
            ms = ms & pImovel.dPoligono(i).Marco & ";" & _
                    CStr(i) & ";" & _
                    decToStrIngles(pImovel.dPoligono(i).X, 2) & ";" & _
                    decToStrIngles(pImovel.dPoligono(i).dE, 3) & ";" & _
                    decToStrIngles(pImovel.dPoligono(i).Y, 2) & ";" & _
                    decToStrIngles(pImovel.dPoligono(i).dN, 3) & ";" & _
                    decToStrIngles(pImovel.dPoligono(i).Z, 2) & ";" & _
                    decToStrIngles(pImovel.dPoligono(i).dH, 3) & ";" & _
                    AptDefs(getAptDefIndex(CStr(pImovel.dPoligono(i).MetodoLevantamento))).CodigoIncra & ";" & _
                    AptDefs(getAptDefIndex(CStr(pImovel.dPoligono(i).TipoLimite))).CodigoIncra & ";" & _
                    auxStr & vbCrLf
        Next i
    Close #1
    
    WriteStringToFile ms, savedir
End Sub


Public Sub ExpGoogleEarthKml(savedir As String)
Dim ms As String, mr As String
Dim i As Integer
Dim nx As Double, ny As Double
        
    If savedir = "" Then Exit Sub
    ms = ""
    mr = ""
    
    For i = 1 To pImovel.NumCoords
        UTM2LL nx, ny, (pImovel.dPoligono(i).X), (pImovel.dPoligono(i).Y)
        mr = mr & Fix(ny) & "." & strDec(ny, 10) & "," & Fix(nx) & "." & strDec(nx, 10) & "," & Fix(pImovel.dPoligono(i).Z) & "." & strDec(pImovel.dPoligono(i).Z, 2) & " "
    Next i
        UTM2LL nx, ny, (pImovel.dPoligono(1).X), (pImovel.dPoligono(1).Y)
        mr = mr & Fix(ny) & "." & strDec(ny, 10) & "," & Fix(nx) & "." & strDec(nx, 10) & "," & Fix(pImovel.dPoligono(1).Z) & "." & strDec(pImovel.dPoligono(1).Z, 2) & " "
    
    ms = LoadFileToString(App.Path & "\@gearth.dat")
    ms = Replace(ms, "CAMPO_NOMEIMOVEL", GetParamData("CAMPO_NOMEIMOVEL"))
    ms = Replace(ms, "CAMPO_COORDS", UTF8_Encode(mr))
    WriteStringToFile ms, savedir
End Sub



