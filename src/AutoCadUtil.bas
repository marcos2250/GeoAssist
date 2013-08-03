Attribute VB_Name = "AutoCadUtil"
Option Explicit

Public Sub ExpAutoCadDXF(savedir As String)
Dim ms As String, mr As String
Dim dxfScale As Single
Dim auxX As Double, auxY As Double, auxA As Double, auxB As Double, auxC As Single, auxD As Single
Dim entCount As Long
Dim i As Integer
        
        
    On Error GoTo erro
    If savedir = "" Then Exit Sub
        
    dxfScale = (1 / form1.Escala) * 100
    ms = ""
    mr = ""
    entCount = 3
    
    'POLIGONO PRINCIPAL
    For i = 1 To pImovel.NumCoords
        mr = mr & "  0" & vbCrLf
        mr = mr & "VERTEX" & vbCrLf
        mr = mr & "  8" & vbCrLf
        mr = mr & "1" & vbCrLf
        mr = mr & "  5" & vbCrLf
        mr = mr & CStr(Hex(entCount)) & vbCrLf
        mr = mr & "  6" & vbCrLf
        mr = mr & "CONTINUOUS" & vbCrLf
        mr = mr & " 62" & vbCrLf
        mr = mr & "7" & vbCrLf
        mr = mr & " 10" & vbCrLf
        mr = mr & CStr(Fix(pImovel.dPoligono(i).X)) & "." & CStr(strDec(pImovel.dPoligono(i).X, 2)) & vbCrLf
        mr = mr & " 20" & vbCrLf
        mr = mr & CStr(Fix(pImovel.dPoligono(i).Y)) & "." & CStr(strDec(pImovel.dPoligono(i).Y, 2)) & vbCrLf
        mr = mr & " 30" & vbCrLf
        mr = mr & CStr(Fix(pImovel.dPoligono(i).Z)) & "." & CStr(strDec(pImovel.dPoligono(i).Z, 2)) & vbCrLf
        entCount = entCount + 1
    Next i
                    
    mr = mr & "  0" & vbCrLf
    mr = mr & "SEQEND" & vbCrLf
    
    mr = mr & "  8" & vbCrLf
    mr = mr & "1" & vbCrLf
    mr = mr & "  5" & vbCrLf
    mr = mr & CStr(Hex(entCount)) & vbCrLf
    entCount = entCount + 1
    
    For i = 1 To pImovel.NumCoords
        'NOMES DE MARCOS
        auxA = Len(pImovel.dPoligono(i).Marco) * (dxfScale / 2)
        mr = mr & "  0" & vbCrLf
        mr = mr & "TEXT" & vbCrLf
        mr = mr & "  8" & vbCrLf
        mr = mr & "1" & vbCrLf
        mr = mr & "  5" & vbCrLf
        mr = mr & CStr(Hex(entCount)) & vbCrLf
        mr = mr & "  6" & vbCrLf
        mr = mr & "CONTINUOUS" & vbCrLf
        mr = mr & " 62" & vbCrLf
        mr = mr & "7" & vbCrLf
        mr = mr & " 10" & vbCrLf
        mr = mr & CStr(Fix(pImovel.dPoligono(i).X + auxA)) & "." & CStr(strDec(pImovel.dPoligono(i).X, 2)) & vbCrLf
        mr = mr & " 20" & vbCrLf
        mr = mr & CStr(Fix(pImovel.dPoligono(i).Y)) & "." & CStr(strDec(pImovel.dPoligono(i).Y, 2)) & vbCrLf
        mr = mr & " 30" & vbCrLf
        mr = mr & CStr(Fix(pImovel.dPoligono(i).Z)) & "." & CStr(strDec(pImovel.dPoligono(i).Z, 2)) & vbCrLf
        mr = mr & "  1" & vbCrLf
        mr = mr & pImovel.dPoligono(i).Marco & vbCrLf
        mr = mr & " 40" & vbCrLf
        mr = mr & Fix(dxfScale) & vbCrLf 'escala do texto
        mr = mr & " 41" & vbCrLf
        mr = mr & "1.0" & vbCrLf
        mr = mr & "  7" & vbCrLf
        mr = mr & "STANDARD" & vbCrLf
        mr = mr & " 72" & vbCrLf
        mr = mr & "4" & vbCrLf
        mr = mr & " 11" & vbCrLf
        mr = mr & CStr(Fix(pImovel.dPoligono(i).X + auxA)) & "." & CStr(strDec(pImovel.dPoligono(i).X, 2)) & vbCrLf
        mr = mr & " 21" & vbCrLf
        mr = mr & CStr(Fix(pImovel.dPoligono(i).Y)) & "." & CStr(strDec(pImovel.dPoligono(i).Y, 2)) & vbCrLf
        mr = mr & " 31" & vbCrLf
        mr = mr & CStr(Fix(pImovel.dPoligono(i).Z)) & "." & CStr(strDec(pImovel.dPoligono(i).Z, 2)) & vbCrLf
        
        entCount = entCount + 1
        
        'SIMBOLOS DOS MARCOS (Bolinhas)
        mr = mr & "  0" & vbCrLf
        mr = mr & "CIRCLE" & vbCrLf
        mr = mr & "  8" & vbCrLf
        mr = mr & "1" & vbCrLf
        mr = mr & "  5" & vbCrLf
        mr = mr & CStr(Hex(entCount)) & vbCrLf
        mr = mr & "  6" & vbCrLf
        mr = mr & "CONTINUOUS" & vbCrLf
        mr = mr & " 62" & vbCrLf
        mr = mr & "7" & vbCrLf
        mr = mr & " 10" & vbCrLf
        mr = mr & CStr(Fix(pImovel.dPoligono(i).X)) & "." & CStr(strDec(pImovel.dPoligono(i).X, 2)) & vbCrLf
        mr = mr & " 20" & vbCrLf
        mr = mr & CStr(Fix(pImovel.dPoligono(i).Y)) & "." & CStr(strDec(pImovel.dPoligono(i).Y, 2)) & vbCrLf
        mr = mr & " 30" & vbCrLf
        mr = mr & CStr(Fix(pImovel.dPoligono(i).Z)) & "." & CStr(strDec(pImovel.dPoligono(i).Z, 2)) & vbCrLf
        mr = mr & " 40" & vbCrLf
        mr = mr & Fix(dxfScale / 4) & vbCrLf 'tamanho das bolinhas
        
        entCount = entCount + 1
        
        
        'COTAS
        If i < pImovel.NumCoords Then
            auxX = pImovel.dPoligono(i + 1).X
            auxY = pImovel.dPoligono(i + 1).Y
        Else
            auxX = pImovel.dPoligono(1).X
            auxY = pImovel.dPoligono(1).Y
        End If
        
        auxC = DistM(pImovel.dPoligono(i).X, pImovel.dPoligono(i).Y, auxX, auxY)
        If auxC > Fix(dxfScale * 7) Then
            auxD = AnguloCoord(pImovel.dPoligono(i).X, pImovel.dPoligono(i).Y, auxX, auxY)
            auxA = ((pImovel.dPoligono(i).X + auxX) / 2) + (Cos(auxD + Pi) * dxfScale * 1.5)
            auxB = ((pImovel.dPoligono(i).Y + auxY) / 2) - (Sin(auxD + Pi) * dxfScale * 1.5)
            
            'COTAS (prop)
            mr = mr & "  0" & vbCrLf
            mr = mr & "TEXT" & vbCrLf
            mr = mr & "  8" & vbCrLf
            mr = mr & "1" & vbCrLf
            mr = mr & "  5" & vbCrLf
            mr = mr & CStr(Hex(entCount)) & vbCrLf
            mr = mr & "  6" & vbCrLf
            mr = mr & "CONTINUOUS" & vbCrLf
            mr = mr & " 62" & vbCrLf
            mr = mr & "7" & vbCrLf
            mr = mr & " 10" & vbCrLf
            mr = mr & CStr(Fix(auxA)) & "." & CStr(strDec(auxA, 2)) & vbCrLf
            mr = mr & " 20" & vbCrLf
            mr = mr & CStr(Fix(auxB)) & "." & CStr(strDec(auxB, 2)) & vbCrLf
            mr = mr & " 30" & vbCrLf
            mr = mr & "0.0" & vbCrLf
            mr = mr & "  1" & vbCrLf
            mr = mr & CStr(Fix(auxC)) & "," & CStr(strDec(CDbl(auxC), 2)) & "m" & vbCrLf
            mr = mr & " 40" & vbCrLf
            mr = mr & Fix(dxfScale) & vbCrLf 'escala do texto
            mr = mr & " 50" & vbCrLf
            mr = mr & Fix(450 - ((auxD / deg2rad) Mod 180)) & vbCrLf 'rotacao
            mr = mr & " 41" & vbCrLf
            mr = mr & "1.0" & vbCrLf
            mr = mr & "  7" & vbCrLf
            mr = mr & "STANDARD" & vbCrLf
            mr = mr & " 72" & vbCrLf
            mr = mr & "4" & vbCrLf
            mr = mr & " 11" & vbCrLf
            mr = mr & CStr(Fix(auxA)) & "." & CStr(strDec(auxA, 2)) & vbCrLf
            mr = mr & " 21" & vbCrLf
            mr = mr & CStr(Fix(auxB)) & "." & CStr(strDec(auxB, 2)) & vbCrLf
            mr = mr & " 31" & vbCrLf
            mr = mr & CStr(Fix(pImovel.dPoligono(i).Z)) & "." & CStr(strDec(pImovel.dPoligono(i).Z, 2)) & vbCrLf
            
            
            entCount = entCount + 1
                            
        End If
    Next i
    
    'NOMES DE CONFRONTANTES E LINHAS DIVISORIAS
    For i = 1 To pImovel.NumCoords
    If pImovel.dPoligono(i).Confrontante <> "" Then
        auxD = AnguloCoord(pImovel.dPoligono(i).X, pImovel.dPoligono(i).Y, pImovel.CentroFolha.X, pImovel.CentroFolha.Y)
        auxX = pImovel.dPoligono(i).X - (Sin(auxD) * Fix((1 / form1.Escala) * 1000))
        auxY = pImovel.dPoligono(i).Y - (Cos(auxD) * Fix((1 / form1.Escala) * 1000))
        auxA = (Len(pImovel.dPoligono(i).Confrontante) * (dxfScale / 2)) / 2
        
        If auxD < 1.57 Or auxD > 4.71 Then
            auxA = -auxA
        End If
        
        

        mr = mr & "  0" & vbCrLf
        mr = mr & "TEXT" & vbCrLf
        mr = mr & "  8" & vbCrLf
        mr = mr & "1" & vbCrLf
        mr = mr & "  5" & vbCrLf
        mr = mr & CStr(Hex(entCount)) & vbCrLf
        mr = mr & "  6" & vbCrLf
        mr = mr & "CONTINUOUS" & vbCrLf
        mr = mr & " 62" & vbCrLf
        mr = mr & "7" & vbCrLf
        mr = mr & " 10" & vbCrLf
        mr = mr & CStr(Fix(auxX + auxA)) & "." & CStr(strDec(auxX, 2)) & vbCrLf
        mr = mr & " 20" & vbCrLf
        mr = mr & CStr(Fix(auxY)) & "." & CStr(strDec(auxY, 2)) & vbCrLf
        mr = mr & " 30" & vbCrLf
        mr = mr & "0.0" & vbCrLf
        mr = mr & "  1" & vbCrLf
        mr = mr & pImovel.dPoligono(i).Confrontante & vbCrLf
        mr = mr & " 40" & vbCrLf
        mr = mr & Fix(dxfScale) & vbCrLf 'escala do texto
        mr = mr & " 41" & vbCrLf
        mr = mr & "1.0" & vbCrLf
        mr = mr & "  7" & vbCrLf
        mr = mr & "STANDARD" & vbCrLf
        mr = mr & " 72" & vbCrLf
        mr = mr & "4" & vbCrLf
        mr = mr & " 11" & vbCrLf
        mr = mr & CStr(Fix(auxX + auxA)) & "." & CStr(strDec(auxX, 2)) & vbCrLf
        mr = mr & " 21" & vbCrLf
        mr = mr & CStr(Fix(auxY)) & "." & CStr(strDec(auxY, 2)) & vbCrLf
        mr = mr & " 31" & vbCrLf
        mr = mr & "0.0" & vbCrLf
        entCount = entCount + 1
        
        mr = mr & "  0" & vbCrLf
        mr = mr & "LINE" & vbCrLf
        mr = mr & "  8" & vbCrLf
        mr = mr & "1" & vbCrLf
        mr = mr & "  5" & vbCrLf
        mr = mr & "11" & vbCrLf
        mr = mr & "  6" & vbCrLf
        mr = mr & "DASHED" & vbCrLf
        mr = mr & " 62" & vbCrLf
        mr = mr & "7" & vbCrLf
        mr = mr & " 10" & vbCrLf
        mr = mr & CStr(Fix(pImovel.dPoligono(i).X)) & "." & CStr(strDec(pImovel.dPoligono(i).X, 2)) & vbCrLf
        mr = mr & " 20" & vbCrLf
        mr = mr & CStr(Fix(pImovel.dPoligono(i).Y)) & "." & CStr(strDec(pImovel.dPoligono(i).Y, 2)) & vbCrLf
        mr = mr & " 30" & vbCrLf
        mr = mr & "0.0" & vbCrLf
        mr = mr & " 11" & vbCrLf
        mr = mr & CStr(Fix(auxX)) & "." & CStr(strDec(auxX, 2)) & vbCrLf
        mr = mr & " 21" & vbCrLf
        mr = mr & CStr(Fix(auxY)) & "." & CStr(strDec(auxY, 2)) & vbCrLf
        mr = mr & " 31" & vbCrLf
        mr = mr & "0.0" & vbCrLf
        entCount = entCount + 1
    End If
    Next i
    
    
    mr = mr & "  0" & vbCrLf
    mr = mr & "ENDSEC"
        
    ms = LoadFileToString(App.Path & "\@acad.dat")

    ms = Replace(ms, "CAMPO_TELAX", CStr(Fix(pImovel.CentroFolha.X)) & "." & CStr(strDec(pImovel.CentroFolha.X, 2)))
    ms = Replace(ms, "CAMPO_TELAY", CStr(Fix(pImovel.CentroFolha.Y)) & "." & CStr(strDec(pImovel.CentroFolha.Y, 2)))
    ms = Replace(ms, "CAMPO_FZOOM", CStr((1 / form1.Escala) * 7000))
    ms = Replace(ms, "CAMPO_REPLACE", mr)
        
    WriteStringToFile ms, savedir

Exit Sub
erro:
    MsgBox "Erro ao salvar o arquivo do AutoCAD" & vbCrLf & savedir, vbExclamation, "Erro"
End Sub

