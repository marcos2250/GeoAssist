Attribute VB_Name = "data"
Option Explicit

'Public Const pMaxCount = 100

Public Type dCoord2d
    X As Double
    Y As Double
End Type

Public Type dAptDefs
    Codigo As Byte
    CodigoIncra As String
    Nome As String
End Type

Public Type dMarcoImovel
    X As Double
    Y As Double
    Z As Double
    dN As Double
    dE As Double
    dH As Double
    Marco As String
    Confrontante As String
    TipoLimite As Byte
    MetodoLevantamento As Byte
End Type

Public Type gImovelRural
    NumCoords As Integer
    Dados(1 To 100, 1 To 3) As String
    dPoligono() As dMarcoImovel
    CentroFolha As dCoord2d
    MinFolha As dCoord2d
    MaxFolha As dCoord2d
End Type


Public AptDefs(0 To 30) As dAptDefs

Public cfgShowGrid As Boolean
Public cfgMarkName As Boolean
Public cfgPrintView As Boolean
Public cfgUnsaved As Boolean

Public NumParamsGeral As Integer
Public prjDir As String
Public prjFileName As String

Public pImovel As gImovelRural




Public Sub ResetWork()
    LoadConfig
    LoadDefault
    prjDir = ""
    prjFileName = ""
    frmMain.Caption = "Geoassist"
    pImovel.NumCoords = 0
    
    DefineDatum
    
    cfgShowGrid = True
    cfgPrintView = False
    cfgMarkName = True
    cfgUnsaved = False
    form1.Escala = 100 * Screen.TwipsPerPixelX
    form1.origemX = 0
    form1.origemY = 0
    form1.marcoSel = 0
    form1.AtualizaGrafico
    
    frmCoordImovel.AtualizaTabela
    frmDadosImovel.AtualizaTabela
End Sub


Public Sub LoadConfig()
Dim s1 As String, s2 As String, s3 As String, s4 As String
Dim i As Integer
    
  NumParamsGeral = 0
  Open App.Path & "\dados.txt" For Input As #1
     Do While Not EOF(1)
        Input #1, s1, s2
        
        If s1 <> "" And s2 <> "" Then
            NumParamsGeral = NumParamsGeral + 1
            pImovel.Dados(NumParamsGeral, 1) = s1
            pImovel.Dados(NumParamsGeral, 2) = s2
            pImovel.Dados(NumParamsGeral, 3) = ""
        End If
     Loop
  Close #1
  
  
  i = 1
  Open App.Path & "\aptdefs.txt" For Input As #1
     Do While Not EOF(1)
        Input #1, s1, s2, s3, s4
        i = i + 1
        AptDefs(i).Codigo = CInt(s2)
        AptDefs(i).CodigoIncra = s3
        AptDefs(i).Nome = s4
     Loop
  Close #1

End Sub

Public Sub LoadDefault()
Dim s1 As String, s2 As String
  
  Open App.Path & "\default.txt" For Input As #1
     Do While Not EOF(1)
        Input #1, s1, s2
        
        If s1 <> "" And s2 <> "" Then
            WriteParamData s1, s2
        End If
     Loop
  Close #1
  
End Sub


Public Sub DefineDatum()
    Dim ndatum As String, nzona As String
    
    ndatum = GetParamData("CAMPO_DATUM")
    If ndatum = "" Then
        WriteParamData "CAMPO_DATUM", "SIRGAS"
        ndatum = "SIRGAS"
    End If
    
    nzona = GetParamData("CAMPO_ZONAUTM")
    If nzona = "" Then
        WriteParamData "CAMPO_ZONAUTM", "23L"
        nzona = "23L"
    End If
    
    datum = SetDatum(ndatum, CInt(Left(nzona, 2)), Right(nzona, 1))
End Sub

Public Sub SaveFile(file As String)
Dim i As Integer
    If Not frmCoordImovel.AtualizaBanco Then Exit Sub
    frmDadosImovel.SalvaDados
    
    Open file For Output As #1
        Print #1, "GEOASSIST201002"
        
        Print #1, "Params," & NumParamsGeral
        
        For i = 1 To NumParamsGeral
            Print #1, Chr(34) & CStr(pImovel.Dados(i, 1)) & Chr(34) & "," & Chr(34) & pImovel.Dados(i, 3) & Chr(34)
        Next i
        
        Print #1, "Coords, " & CStr(pImovel.NumCoords)
        For i = 1 To pImovel.NumCoords
            Print #1, Chr(34) & pImovel.dPoligono(i).Marco & Chr(34) & "," & _
                      Chr(34) & pImovel.dPoligono(i).Confrontante & Chr(34) & "," & _
                      Chr(34) & pImovel.dPoligono(i).X & Chr(34) & "," & _
                      Chr(34) & pImovel.dPoligono(i).Y & Chr(34) & "," & _
                      Chr(34) & pImovel.dPoligono(i).Z & Chr(34) & "," & _
                      Chr(34) & pImovel.dPoligono(i).dE & Chr(34) & "," & _
                      Chr(34) & pImovel.dPoligono(i).dN & Chr(34) & "," & _
                      Chr(34) & pImovel.dPoligono(i).dH & Chr(34) & "," & _
                      Chr(34) & pImovel.dPoligono(i).MetodoLevantamento & Chr(34) & "," & _
                      Chr(34) & pImovel.dPoligono(i).TipoLimite & Chr(34)
        Next i
    Close #1
    
    prjDir = LeftLastChr(file, "\")
    prjFileName = RightLastChr(file, "\")
    cfgUnsaved = False
    frmMain.Caption = "Geoassist - " & prjFileName
End Sub


Public Sub loadFile(file As String)
Dim i As Integer, n As Integer
Dim s(1 To 8) As String
Dim Versao As Byte

    ResetWork
    
    Open file For Input As #1
     
     Input #1, s(1)
     If s(1) = "GEOASSIST201001" Then
        Versao = 1
     ElseIf (s(1) = "GEOASSIST201002") Then
        Versao = 2
     Else
        Exit Sub
     End If
     
     Input #1, s(1), s(2)
     If s(1) = "Params" Then
        n = CInt(s(2))
     Else
        GoTo erro
     End If
          
     For i = 1 To n
        Input #1, s(1), s(2)
        WriteParamData s(1), s(2)
     Next i
    
     Input #1, s(1), s(2)
     If s(1) = "Coords" Then
        pImovel.NumCoords = CInt(s(2))
     Else
        GoTo erro
     End If
          
     ReDim pImovel.dPoligono(pImovel.NumCoords)
          
     For i = 1 To pImovel.NumCoords
        If Versao = 2 Then
            Input #1, pImovel.dPoligono(i).Marco, pImovel.dPoligono(i).Confrontante, s(1), s(2), s(3), s(4), s(5), s(6), s(7), s(8)
        Else
            Input #1, pImovel.dPoligono(i).Marco, pImovel.dPoligono(i).Confrontante, s(1), s(2), s(3)
        End If
        
        pImovel.dPoligono(i).X = CDec(s(1))
        pImovel.dPoligono(i).Y = CDec(s(2))
        pImovel.dPoligono(i).Z = CDec(s(3))
        
        If Versao = 2 Then
            pImovel.dPoligono(i).dE = CDec(s(4))
            pImovel.dPoligono(i).dN = CDec(s(5))
            pImovel.dPoligono(i).dH = CDec(s(6))
            pImovel.dPoligono(i).MetodoLevantamento = CDec(s(7))
            pImovel.dPoligono(i).TipoLimite = CDec(s(8))
        End If
     Next i
    
    
    Close #1
    
    prjDir = LeftLastChr(file, "\")
    prjFileName = RightLastChr(file, "\")
    cfgUnsaved = False
    CalculaLimitesFolha
    frmMain.Caption = "Geoassist - " & prjFileName
    form1.fitView
    DefineDatum
    frmCoordImovel.AtualizaTabela
    frmDadosImovel.AtualizaTabela

    Exit Sub
erro:
    Close #1
    MsgBox "Erro na abertura do arquivo " & file, vbCritical
End Sub

Public Function LoadFileToString(file As String)
    On Error GoTo erro
    Dim dataStream As String
    Open file For Input As #1
        dataStream = Input(LOF(1), 1)
    Close #1
    LoadFileToString = dataStream
    Exit Function
erro:
    MsgBox "Erro ao carregar o arquivo" & vbCrLf & file, vbExclamation, "Erro"
End Function

Public Sub WriteStringToFile(dataStream As String, file As String)
    On Error GoTo erro
    Open file For Output As #1
    Print #1, dataStream
    Close #1
    Exit Sub
erro:
    MsgBox "Erro ao salvar o arquivo" & vbCrLf & file & vbCrLf & _
    "Certifique-se que o arquivo não esteja bloqueado para somente leitura.", vbExclamation, "Erro"
End Sub

Public Function GetParamData(param As String) As String
Dim i As Integer
    For i = 1 To NumParamsGeral
        If pImovel.Dados(i, 1) = param Then
            GetParamData = pImovel.Dados(i, 3)
            Exit For
        End If
    Next i
End Function

Public Function GetParamName(param As String) As String
Dim i As Integer
    For i = 1 To NumParamsGeral
        If pImovel.Dados(i, 1) = param Then
            GetParamName = pImovel.Dados(i, 2)
            Exit For
        End If
    Next i
End Function

Public Sub WriteParamData(param As String, dado As String)
Dim i As Integer
    For i = 1 To NumParamsGeral
        If pImovel.Dados(i, 1) = param Then
            pImovel.Dados(i, 3) = dado
            Exit For
        End If
    Next i
End Sub


Public Function getAptDefIndex(busca As String) As Integer
Dim i As Integer
    For i = 1 To 30
        If AptDefs(i).CodigoIncra = busca Or CStr(AptDefs(i).Codigo) = Trim(busca) Or AptDefs(i).Nome = busca Then
            getAptDefIndex = i
            Exit For
        End If
    Next i
End Function

