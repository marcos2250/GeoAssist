Attribute VB_Name = "Geo"
'Módulo de funcoes Geograficas
Option Explicit

Public Const Pi As Double = 3.14159265358979
Public Const deg2rad As Double = 1.74532925199433E-02


Public Type DatumData
    ER As Double
    ERmenor As Double
    k0 As Double
    eS As Double
    fA As Double
    Zona As Double
    Letra As String
    Nome As String
    Codigo As String
    MC As Single
End Type
Public datum As DatumData


Public Function SetDatum(DatumCodigo As String, Zona As Byte, Letra As String) As DatumData
    Dim v1 As String, v2 As String, v3 As String, v4 As String
    
    Open App.Path & "\datums.txt" For Input As #2
        Do
            Input #2, v1, v2, v3, v4
            If v1 = DatumCodigo Or v2 = DatumCodigo Then
                SetDatum.Codigo = v1
                SetDatum.Nome = v2
                SetDatum.ER = strToDec(v3)
                SetDatum.ERmenor = strToDec(v4)
                'Datum.fA = strToDec(v5)
            End If
        Loop Until EOF(2)
    Close #2

    If SetDatum.Nome = "" Then Exit Function
    SetDatum.k0 = 0.9996 'Fator de Escala K
    SetDatum.Zona = Zona
    SetDatum.Letra = Letra
    SetDatum.fA = Sqr((SetDatum.ER ^ 2) - (SetDatum.ERmenor ^ 2)) / SetDatum.ER 'Fator de achatamento
    SetDatum.eS = SetDatum.fA ^ 2
    
    'Datum.eS = (2 * (1 / Datum.fA)) - ((1 / Datum.fA) ^ 2) 'Quadrado da Ecentricidade
    'Datum.fA = strToDec(v5)
    
    SetDatum.MC = MeridianoCentral(Zona) 'Meridiano Central
End Function

Public Function MeridianoCentral(ByVal Zona As Integer) As Single
    MeridianoCentral = (Zona - 1) * 6 - 180 + 3
End Function

Public Function ZonaUTMNumero(Longitude As Single) As Integer
    If Longitude < 0 Then
        ZonaUTMNumero = Fix((180 + Longitude) / 6) + 1
    Else
        ZonaUTMNumero = Fix((Longitude / 6) + 31)
    End If
End Function

Public Function ZonaUTMLetra(Latitude As Single) As String
    ZonaUTMLetra = Mid("ABCDEFGHJKLMNPQRSTUVWXYZ", Round((Latitude / 8) - 0.5) + 13, 1)
End Function


Public Function FatorK(Latitude As Double, Longitude As Double) As Double
'Calcula o Fator de Escala K
    'FatorK = Ko / Sqr(1 - (Cos(Latitude) * Sin(Longitude - MC)) ^ 2)
    
    'SetDatum Datum.Nome, ZonaUTMNumero((Longitude)), ZonaUTMLetra((Latitude))
    FatorK = datum.k0 / Sqr(1 - (Cos(Latitude * deg2rad) * Sin((Longitude - datum.MC) * deg2rad)) ^ 2)
End Function


Public Function FatorK_UTM(Easting As Double, Northing As Double) As Double
'Calcula o Fator de Escala baseando em coordenadas UTM
    Dim alpha As Double, M25 As Double, N25 As Double, O25 As Double, nAux As Double, mAux As Double
    
    alpha = 1.00505262472535 * datum.ER * (1 - datum.eS) * deg2rad

    M25 = Northing - (10 ^ 7)
    N25 = Easting - 500000
    
    O25 = M25 / (0.9996 * alpha * (180 / (Atn(1) * 4)))

    nAux = datum.ER / ((1 - datum.eS * (Sin(O25) ^ 2)) ^ (0.5))
    mAux = (datum.ER * (1 - datum.eS)) / ((1 - datum.eS * (Sin(O25) ^ 2)) ^ (3 / 2))

    FatorK_UTM = 0.9996 * (1 + (N25 ^ 2 / (2 * nAux * mAux)) + (N25 ^ 4 / (24 * nAux ^ 2 * mAux ^ 2)))

'Fonte:
'www.topografia.com.br/br/downloads/UTMGEO_GEOUTM.xls
End Function


Public Function ConvMeridiana(Latitude As Double, Longitude As Double) As Double
'Calcula a Convergencia Meridiana (valores em graus decimais)
    'SetDatum Datum.Nome, ZonaUTMNumero((Longitude)), ZonaUTMLetra((Latitude))
    ConvMeridiana = (((Longitude - datum.MC) * deg2rad) * Sin(Latitude * deg2rad)) / deg2rad
'Fonte
'http://www.geodesia.ufrgs.br/trabalhosdidaticos/Topografia_Aplicada_A_Engenharia_Civil/Apostila/Apostila_TopoAplicadaEng_2007.pdf
End Function

Public Sub LL2UTM(ByRef Easting As Double, ByRef Northing As Double, ByVal Latitude As Double, ByVal Longitude As Double)
'Converte coordenadas geograficas Latitude / Longitude para UTM
    Dim LatRad As Double, LongRad As Double
    Dim LongOrigin As Double, LongOriginRad As Double
    Dim zonenumber As Integer
    Dim eccSquared As Double, eccPrimeSquared As Double
    Dim n As Double, T As Double, c As Double, a As Double, M As Double, ER As Double, k0 As Double
    Dim UTMEasting As Double, UTMNorthing As Double
       
    ER = datum.ER
    k0 = datum.k0
    eccSquared = datum.eS
        
    LatRad = Latitude * deg2rad 'converte para radianos
    LongRad = Longitude * deg2rad
    
    zonenumber = Fix((Longitude + 180) / 6) + 1
    
    LongOrigin = (zonenumber - 1) * 6 - 180 + 3 '+3 puts origin in middle of zone
    LongOriginRad = LongOrigin * deg2rad
    
    eccPrimeSquared = (eccSquared) / (1 - eccSquared)
    
    n = ER / Sqr(1 - eccSquared * Sin(LatRad) * Sin(LatRad))
    T = Tan(LatRad) * Tan(LatRad)
    c = eccPrimeSquared * Cos(LatRad) * Cos(LatRad)
    a = Cos(LatRad) * (LongRad - LongOriginRad)

    M = ER * ((1 - eccSquared / 4 - 3 * eccSquared * eccSquared / 64 - 5 * eccSquared * eccSquared * eccSquared / 256) * LatRad _
                - (3 * eccSquared / 8 + 3 * eccSquared * eccSquared / 32 + 45 * eccSquared * eccSquared * eccSquared / 1024) * Sin(2 * LatRad) _
                                    + (15 * eccSquared * eccSquared / 256 + 45 * eccSquared * eccSquared * eccSquared / 1024) * Sin(4 * LatRad) _
                                    - (35 * eccSquared * eccSquared * eccSquared / 3072) * Sin(6 * LatRad))
    
    UTMEasting = (k0 * n * (a + (1 - T + c) * a * a * a / 6 _
                    + (5 - 18 * T + T * T + 72 * c - 58 * eccPrimeSquared) * a * a * a * a * a / 120) _
                    + 500000)

    UTMNorthing = (k0 * (M + n * Tan(LatRad) * (a * a / 2 + (5 - T + 9 * c + 4 * c * c) * a * a * a * a / 24 _
                 + (61 - 58 * T + T * T + 600 * c - 330 * eccPrimeSquared) * a * a * a * a * a * a / 720)))
                 
                 
    If (Latitude < 0) Then UTMNorthing = UTMNorthing + 10000000# '10000000 meter offset for southern hemisphere
    
    Easting = Round(UTMEasting, 2)
    Northing = Round(UTMNorthing, 2)
    
    'Bibliografia
    ' http://www.gpsy.com/gpsinfo/geotoutm/
End Sub



Public Sub UTM2LL(ByRef Latitude As Double, ByRef Longitude As Double, ByVal Easting As Double, ByVal Northing As Double)
'Converte coordenadas UTM para coordenadas geograficas Latitude / Longitude
    Dim LatRad As Double, LongRad As Double
    Dim LongOrigin As Double
    Dim zonenumber As Integer
    Dim eccSquared As Double, eccPrimeSquared As Double
    Dim e1 As Double, N1 As Double, T1 As Double, C1 As Double, R1 As Double, D As Double, M As Double
    Dim ER As Double, k0 As Double
    Dim mu As Double, phi1 As Double, phi1Rad As Double
    Dim X As Double, Y As Double
    
    
    'SetDatum DatumNome, ZonaNumero, ZonaLetra
    ER = datum.ER
    k0 = datum.k0
    eccSquared = datum.eS
    zonenumber = datum.Zona
    
    Select Case UCase(datum.Letra)
    Case "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"
        Y = Northing - 10000000 '//remove 10,000,000 meter offset used for southern hemisphere
    Case Else
        Y = Northing
    End Select
    
    X = Easting - 500000 'remove 500,000 meter offset for longitude


    LongOrigin = (zonenumber - 1) * 6 - 180 + 3 '+3 puts origin in middle of zone

    M = Y / k0
    
    e1 = (1 - Sqr(1 - eccSquared)) / (1 + Sqr(1 - eccSquared))
    
    eccPrimeSquared = (eccSquared) / (1 - eccSquared)

    
    mu = M / (ER * (1 - ((eccSquared) / 4) - (3 * (eccSquared ^ 2) / 64) - (5 * (eccSquared ^ 3) / 256)))

    phi1Rad = mu + (3 * e1 / 2 - 27 * e1 * e1 * e1 / 32) * Sin(2 * mu) _
                + (21 * e1 * e1 / 16 - 55 * e1 * e1 * e1 * e1 / 32) * Sin(4 * mu) _
                + (151 * e1 * e1 * e1 / 96) * Sin(6 * mu)
                
    N1 = ER / Sqr(1 - (eccSquared * Sin(phi1Rad) * Sin(phi1Rad)))
    T1 = Tan(phi1Rad) * Tan(phi1Rad)
    C1 = eccPrimeSquared * Cos(phi1Rad) * Cos(phi1Rad)
    R1 = ER * (1 - eccSquared) / ((1 - eccSquared * Sin(phi1Rad) * Sin(phi1Rad)) ^ 1.5)
    D = X / (N1 * k0)

    LatRad = phi1Rad - (N1 * Tan(phi1Rad) / R1) * (D * D / 2 - (5 + 3 * T1 + 10 * C1 - 4 * C1 * C1 - 9 * eccPrimeSquared) * D * D * D * D / 24 _
                    + (61 + 90 * T1 + 298 * C1 + 45 * T1 * T1 - 252 * eccPrimeSquared - 3 * C1 * C1) * D * D * D * D * D * D / 720)

    LongRad = (D - (1 + 2 * T1 + C1) * D * D * D / 6 + (5 - 2 * C1 + 28 * T1 - 3 * C1 * C1 + 8 * eccPrimeSquared + 24 * T1 * T1) _
                    * D * D * D * D * D / 120) / Cos(phi1Rad)
    
    Latitude = Round(LatRad / deg2rad, 8)
    Longitude = Round(LongOrigin + LongRad / deg2rad, 8)
    
    'Bibliografia
    ' http://www.gpsy.com/gpsinfo/geotoutm/
End Sub

Public Sub ConverteDatums(ByRef latitude_destino As Double, ByRef longitude_destino As Double, ByRef altura_destino As Double, ByRef datumDestino As String, ByVal latitude_origem As Double, ByVal longitude_origem As Double, ByVal altura_origem As Double, ByVal datumorigem As String)
'Converte coordenadas Lat/Long entre SIRGAS e SAD69
Dim dta_eixomaior As Double, dta_eixomenor As Double, dtb_eixomaior As Double, dtb_eixomenor As Double
Dim dtOrigem As DatumData, dtDestino As DatumData
Dim nOrigem As Double, dta_eS As Double
Dim Longitude As Double, Latitude As Double, altitude As Double
Dim X As Double, Y As Double, Z As Double
Dim sgnLat As Integer, sgnLon As Integer
    
    If latitude_origem > 0 Then
        sgnLat = 1
    Else
        sgnLat = -1
    End If
    
    If longitude_origem > 0 Then
        sgnLon = 1
    Else
        sgnLon = -1
    End If
    
    dtOrigem = SetDatum(datumorigem, (datum.Zona), datum.Letra)
    dtDestino = SetDatum(datumDestino, (datum.Zona), datum.Letra)

    nOrigem = dtOrigem.ER / (Sqr(1 - dtOrigem.eS * (Sin(latitude_origem * deg2rad) ^ 2)))

    X = (nOrigem + altura_origem) * Cos(latitude_origem * deg2rad) * Cos(longitude_origem * deg2rad)
    Y = (nOrigem + altura_origem) * Cos(latitude_origem * deg2rad) * Sin(longitude_origem * deg2rad)
    Z = (nOrigem * (1 - dtOrigem.eS) + altura_origem) * Sin(latitude_origem * deg2rad)

    If datumorigem = "SAD69" And datumDestino = "SIRGAS" Then
        X = X - 67.35 'deltax
        Y = Y + 3.88 'deltay
        Z = Z - 38.22 ' deltaz
    End If
    
    If datumorigem = "SIRGAS" And datumDestino = "SAD69" Then
        X = X + 67.35 'deltax
        Y = Y - 3.88 'deltay
        Z = Z + 38.22 ' deltaz
    End If

    Longitude = Abs(Atn(Y / X))
    Latitude = Abs((Atn((Z + (dtDestino.ERmenor * ((dtDestino.ER ^ 2 - dtDestino.ERmenor ^ 2) / dtDestino.ERmenor ^ 2)) * (Sin((Atn(((Z) / (Sqr((X) ^ 2 + (Y) ^ 2))) * (dtDestino.ER / dtDestino.ERmenor))))) ^ 3) / ((Sqr((X) ^ 2 + (Y) ^ 2)) - (dtDestino.ER * ((dtDestino.ER ^ 2 - dtDestino.ERmenor ^ 2) / dtDestino.ER ^ 2)) * (Cos((Atn(((Z) / (Sqr((X) ^ 2 + (Y) ^ 2))) * (dtDestino.ER / dtDestino.ERmenor))))) ^ 3))))
    altura_destino = ((Sqr((X) ^ 2 + (Y) ^ 2))) / Cos(Latitude) - ((dtDestino.ER) / (1 - ((dtDestino.ER ^ 2 - dtDestino.ERmenor ^ 2) / dtDestino.ER ^ 2) * Sin((((Atn((Z + (dtDestino.ERmenor * ((dtDestino.ER ^ 2 - dtDestino.ERmenor ^ 2) / dtDestino.ERmenor ^ 2)) * (Sin((Atn(((Z) / (Sqr((X) ^ 2 + (Y) ^ 2))) * (dtDestino.ER / dtDestino.ERmenor))))) ^ 3) / ((Sqr((X) ^ 2 + (Y) ^ 2)) - (dtDestino.ER * ((dtDestino.ER ^ 2 - dtDestino.ERmenor ^ 2) / dtDestino.ER ^ 2)) * (Cos((Atn(((Z) / (Sqr((X) ^ 2 + (Y) ^ 2))) * (dtDestino.ER / dtDestino.ERmenor))))) ^ 3)))))) ^ 2) ^ (1 / 2))
    
    longitude_destino = (Longitude / deg2rad) * sgnLat
    latitude_destino = (Latitude / deg2rad) * sgnLon

'Fonte:
'www.topografia.com.br/br/downloads/UTMGEO_GEOUTM.xls
End Sub


Public Function AnguloCoord(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Double
'Calcula o angulo entre duas coordenadas num plano cartesiano

    If X1 = X2 And Y1 = Y2 Then
        AnguloCoord = 0
        Exit Function
    End If
    
    If Y2 > Y1 Then
        If X2 > X1 Then
             AnguloCoord = (Pi / 2) - Atn(Abs(Y2 - Y1) / Abs(X2 - X1))
        Else
             If X1 <> X2 Then
                AnguloCoord = Atn(Abs(Y2 - Y1) / Abs(X1 - X2)) + (3 * (Pi / 2))
             Else
                AnguloCoord = 0
             End If
        End If
    Else
        If X2 > X1 Then
            AnguloCoord = Atn(Abs(Y1 - Y2) / Abs(X2 - X1)) + (Pi / 2)
        Else
            If X1 <> X2 Then
                AnguloCoord = (3 * (Pi / 2)) - Atn(Abs(Y1 - Y2) / Abs(X1 - X2))
            Else
                AnguloCoord = Pi
            End If
        End If
    End If
End Function


Public Function DistM(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Double
'Calcula a distancia entre dois pontos (Pitagoras)
    DistM = Sqr(((X1 - X2) ^ 2) + ((Y1 - Y2) ^ 2))
End Function


Public Sub CalculaAreaPerim()
Dim i As Integer
Dim pArea As Double, pPerim As Double
Dim p1X As Double, p1Y As Double
Dim p2X As Double, p2Y As Double
    
    pPerim = 0
    pArea = 0
    
    For i = 1 To pImovel.NumCoords
        p1X = pImovel.dPoligono(i).X
        p1Y = pImovel.dPoligono(i).Y
        
        If i < pImovel.NumCoords Then
            p2X = pImovel.dPoligono(i + 1).X
            p2Y = pImovel.dPoligono(i + 1).Y
        Else
            p2X = pImovel.dPoligono(1).X
            p2Y = pImovel.dPoligono(1).Y
        End If
                   
        'Fórmulas de area e perimetro
        pPerim = pPerim + Round(Sqr((Abs(p1Y - p2Y) ^ 2) + (Abs(p1X - p2X) ^ 2)), 2)
        pArea = pArea + Round(((p1X * p2Y) - (p1Y * p2X)), 4)
    Next i
        
    'Correcao da area e retorno dos valores
    Dim strArea As String, strPerimetro As String, strNumCoords As String
    strArea = decToStr(Abs(pArea / 20000), 4)
    strPerimetro = decToStr(pPerim, 2)
    strNumCoords = CStr(pImovel.NumCoords)
    
    WriteParamData "CAMPO_AREAMED", strArea
    WriteParamData "CAMPO_PERIMETRO", strPerimetro
    WriteParamData "CAMPO_NUMPONTOS", strNumCoords
    
    form1.Caption = "Área: " + strArea + " ha, Perímetro " + strPerimetro + " m, Vértices: " + strNumCoords
    
End Sub


Public Sub CalculaLimitesFolha()
Dim i As Integer

    pImovel.MaxFolha.X = pImovel.dPoligono(1).X
    pImovel.MaxFolha.Y = pImovel.dPoligono(1).Y
    pImovel.MinFolha.X = pImovel.dPoligono(1).X
    pImovel.MinFolha.Y = pImovel.dPoligono(1).Y
    
    For i = 2 To pImovel.NumCoords
        If pImovel.dPoligono(i).X < pImovel.MinFolha.X Then pImovel.MinFolha.X = pImovel.dPoligono(i).X
        If pImovel.dPoligono(i).Y < pImovel.MinFolha.Y Then pImovel.MinFolha.Y = pImovel.dPoligono(i).Y
        If pImovel.dPoligono(i).X > pImovel.MaxFolha.X Then pImovel.MaxFolha.X = pImovel.dPoligono(i).X
        If pImovel.dPoligono(i).Y > pImovel.MaxFolha.Y Then pImovel.MaxFolha.Y = pImovel.dPoligono(i).Y
    Next i
    
    pImovel.CentroFolha.X = (pImovel.MinFolha.X + pImovel.MaxFolha.X) / 2
    pImovel.CentroFolha.Y = (pImovel.MinFolha.Y + pImovel.MaxFolha.Y) / 2
End Sub
