VERSION 5.00
Begin VB.Form form1 
   BackColor       =   &H00000000&
   Caption         =   "Visualização"
   ClientHeight    =   6795
   ClientLeft      =   210
   ClientTop       =   780
   ClientWidth     =   9480
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   Begin VB.PictureBox pa 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   1200
      ScaleHeight     =   2175
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   1020
      Width           =   3015
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Escala As Single  'Twips por metro  (1 pixel = 15 twips)
Public dragging As Boolean ', zooming As Boolean
Public origemX As Single, origemY As Single

Public marcoSel As Integer

Dim origemClickX As Single, origemClickY As Single
Dim mouseX As Single, mouseY As Single
Dim mouseClickX As Single, mouseClickY As Single
Dim minX As Long, maxX As Long, minY As Long, maxY As Long
Dim aux As Single

Dim gradeSpc As Single
Dim i As Long


Sub toggleGrid()
    If cfgShowGrid Then
        cfgShowGrid = False
    Else
        cfgShowGrid = True
    End If
    form1.AtualizaGrafico
End Sub


Sub toggleMarkName()
    If cfgMarkName Then
        cfgMarkName = False
    Else
        cfgMarkName = True
    End If
    form1.AtualizaGrafico
End Sub

Sub togglePrintView()
    If cfgPrintView Then
        cfgPrintView = False
    Else
        cfgPrintView = True
    End If
    form1.AtualizaGrafico
End Sub





Private Sub Form_GotFocus()
    AtualizaGrafico
End Sub

Private Sub Form_Paint()
    AtualizaGrafico
End Sub

Private Sub Form_Resize()
    If form1.WindowState = vbMinimized Then Exit Sub
    pa.Top = 0
    pa.Left = 0
    pa.Height = form1.Height - (35 * 15)
    pa.Width = form1.Width - (15 * 15)
    AtualizaGrafico
End Sub


Public Sub gPrintf(texto As String, X As Long, Y As Long)
        If X < pa.Width And Y < pa.Height And X > 0 And Y > 0 Then
            pa.CurrentX = X
            pa.CurrentY = Y
            pa.Print texto
        End If
End Sub



Private Sub quickhelp()
        gPrintf "         Ajuda da Visualização:" & vbCrLf & vbCrLf & _
                "               - Teclas direcionais: Mover" & vbCrLf & _
                "               - Botões do Mouse: Mover" & vbCrLf & _
                "               - Teclas + e -: Zoom rapido" & vbCrLf & _
                "               - F12: Enquadra a área na tela" & vbCrLf & _
                "               - S: Zoom definindo a escala manualmente" & vbCrLf & _
                "               - F: Localizar coordenada" & vbCrLf & _
                "               - G: Liga/desliga Grade de coordenadas" & vbCrLf & _
                "               - M: Liga/desliga Nomes dos marcos" & vbCrLf & _
                "               - P: Liga/desliga Modo de Impressão (fundo branco)" & vbCrLf & _
                "", 100, 1500
End Sub

Sub gZoom(s As Single)
 If Escala < 0.01 And s < 1 Then Exit Sub
 If Escala > 6000 And s > 1 Then Exit Sub
 If s = 0 Then Exit Sub
   
    Dim X As Single, Y As Single

    X = (pa.Width / (Escala * s)) - (pa.Width / Escala)
    Y = (pa.Height / (Escala * s)) - (pa.Height / Escala)
    origemX = origemX - (X / 2)
    origemY = origemY - (Y / 2)
    Escala = Escala * s

    AtualizaGrafico
End Sub

Sub gPan(X As Integer, Y As Integer)
    origemX = origemX + (X / Escala) * 1000
    origemY = origemY + (Y / Escala) * 1000
    AtualizaGrafico
End Sub


Sub customScale()
    On Error Resume Next
    Dim n As Single
    n = CDec(InputBox("Digite a escala em Pixels/metro (de 0,1 a 400): ", "Alterar Escala", 10)) * Screen.TwipsPerPixelX
    If n >= 1 And n <= 100000000 Then
        Escala = n
    End If
    AtualizaGrafico
End Sub

Sub fitView()
    If pImovel.NumCoords < 2 Then
        MsgBox "Não há elementos para serem exibidos."
        Exit Sub
    End If
        
    If pImovel.MaxFolha.X + pImovel.MinFolha.X = 0 Or pImovel.MaxFolha.Y + pImovel.MinFolha.Y = 0 Then Exit Sub
                
    If (Abs(pImovel.MaxFolha.X - pImovel.MinFolha.X) / Abs(pImovel.MaxFolha.Y - pImovel.MinFolha.Y)) > (pa.Width / pa.Height) Then
        Escala = ((pa.Width / Abs(pImovel.MaxFolha.X - pImovel.MinFolha.X))) * 0.9
    Else
        Escala = ((pa.Height / Abs(pImovel.MaxFolha.Y - pImovel.MinFolha.Y))) * 0.9
    End If
    
    If Escala <> 0 Then
        origemX = ((pImovel.MaxFolha.X + pImovel.MinFolha.X) / 2) - (pa.Width / 2) / Escala
        origemY = ((pImovel.MaxFolha.Y + pImovel.MinFolha.Y) / 2) - (pa.Height / 2) / Escala
    Else
        Escala = 1
    End If
    
    AtualizaGrafico
End Sub

Sub customOrigin()
    On Error Resume Next
    Dim T As String
    T = InputBox("Digite a coordenada no campo abaixo" & vbCrLf & "Sintaxe: X,Y" & vbCrLf & "Nao utilize valores decimais", "Buscar coordenada", "0,0")
    
    For i = 1 To Len(T)
        If Mid(T, i, 1) = "," Then
            origemX = Int(Mid(T, 1, i - 1))
            origemY = Int(Mid(T, i + 1, Len(T) - i))
        End If
    Next i
    
    origemX = origemX - (pa.Width / 2) / Escala
    origemY = origemY - (pa.Height / 2) / Escala
    
    AtualizaGrafico
End Sub




Sub AtualizaGrafico()
    
    If cfgPrintView Then
        pa.BackColor = vbWhite
        pa.ForeColor = vbBlack
    Else
        pa.BackColor = vbBlack
        pa.ForeColor = vbWhite
    End If
    
    pa.Cls
        
    If dragging = True Then
        origemX = origemClickX - ((mouseX - mouseClickX) / Escala)
        origemY = origemClickY + ((mouseY - mouseClickY) / Escala)
    End If
    
    'If zooming = True Then
        'gZoom 1 + ((mouseX - mouseClickX) / 1000000)
    'End If
    
    
    'AREA DE TRABALHO
    '----------------
    
    If cfgShowGrid Then
         minX = Int(origemX)
         maxX = Int(origemX + (pa.Width / Escala))
         minY = Int(origemY)
         maxY = Int(origemY + (pa.Height / Escala))
         
         If Escala > 0.5 Then
            gradeSpc = (10 ^ (5 - Len(CStr(Int(Escala))))) / 10
         Else
            gradeSpc = Int(1 / Escala) * 1000
         End If
         If gradeSpc < 1 Then gradeSpc = 1
                  
         For i = Fix(origemX / gradeSpc) To (Fix(origemX / gradeSpc) + Fix((pa.Width / Escala) / gradeSpc) + 1)
            pa.Line ((-origemX + (i * gradeSpc)) * Escala, 0)-((-origemX + (i * gradeSpc)) * Escala, pa.Height), RGB(100, 100, 100)
            gPrintf CStr(i * gradeSpc), (-origemX + (i * gradeSpc)) * Escala, pa.Height * 0.96
         Next i
                
         For i = Fix(origemY / gradeSpc) To (Fix(origemY / gradeSpc) + Fix((pa.Height / Escala) / gradeSpc) + 1)
            pa.Line (0, pa.Height - (-origemY + (i * gradeSpc)) * Escala)-(pa.Width, pa.Height - (-origemY + (i * gradeSpc)) * Escala), RGB(100, 100, 100)
            gPrintf CStr(i * gradeSpc), 50, pa.Height - (-origemY + (i * gradeSpc)) * Escala
         Next i
    End If
        
    If Not cfgPrintView Then
        If Escala > 15 Then
            gPrintf Int(Escala / 15) & " px/m", pa.Width * 0.9, pa.Height * 0.9
        Else
            If Escala <> 0 Then gPrintf Int(1 / (Escala / 15)) & " m/px", pa.Width * 0.9, pa.Height * 0.9
        End If
    Else
        If Printer.Width > Printer.Height Then
            aux = (pa.Width / Escala) / 0.297
        Else
            aux = (pa.Width / Escala) / 0.21
        End If
           
        gPrintf "Escala = 1:" & Round(aux) & "", pa.Width * 0.85, pa.Height * 0.85
        gPrintf "Folha A4", pa.Width * 0.85, (pa.Height * 0.85) + 300
    End If
    
    
    'OBJETOS
    '-------
    
    
    
    
    If pImovel.NumCoords < 2 Then Exit Sub
    For i = 1 To pImovel.NumCoords
            minX = Fix((pImovel.dPoligono(i).X - origemX) * Escala)
            minY = Fix(pa.Height - ((pImovel.dPoligono(i).Y - origemY) * Escala))
            
            If i < pImovel.NumCoords Then
                pa.Line (minX, minY) _
                    -(((pImovel.dPoligono(i + 1).X - origemX) * Escala), _
                    (pa.Height - ((pImovel.dPoligono(i + 1).Y - origemY) * Escala))), pa.ForeColor
            Else
                pa.Line (((pImovel.dPoligono(pImovel.NumCoords).X - origemX) * Escala), _
                    (pa.Height - ((pImovel.dPoligono(pImovel.NumCoords).Y - origemY) * Escala))) _
                    -(((pImovel.dPoligono(1).X - origemX) * Escala), _
                    (pa.Height - ((pImovel.dPoligono(1).Y - origemY) * Escala))), pa.ForeColor
            End If
    
            pa.Circle (minX, minY), 30, pa.ForeColor
            
            If cfgMarkName Then gPrintf pImovel.dPoligono(i).Marco, CLng(minX), CLng(minY)
                                    
            If pImovel.dPoligono(i).Confrontante <> "" Then
                aux = AnguloCoord(pImovel.dPoligono(i).X, pImovel.dPoligono(i).Y, CDbl(pImovel.CentroFolha.X), CDbl(pImovel.CentroFolha.Y))
                pa.Line (minX, minY)-(minX - Sin(aux) * 1500, minY + Cos(aux) * 1500), ColorPreview(vbYellow)
                
                If aux < 1.57 Or aux > 4.71 Then
                    gPrintf pImovel.dPoligono(i).Confrontante, CLng(minX - (Sin(aux) * 1500) - (Len(pImovel.dPoligono(i).Confrontante) * 100)), CLng(minY + Cos(aux) * 1500)
                Else
                    gPrintf pImovel.dPoligono(i).Confrontante, CLng(minX - Sin(aux) * 1500), CLng(minY + Cos(aux) * 1500)
                End If
            End If
    Next i
    
    If form1.marcoSel <> 0 Then
        pa.Circle (Fix((pImovel.dPoligono(form1.marcoSel).X - origemX) * Escala), Fix(pa.Height - ((pImovel.dPoligono(form1.marcoSel).Y - origemY) * Escala))), 120, vbRed
    End If
    
    DoEvents

End Sub


Private Function ColorPreview(iColor As Long) As Long
    If cfgPrintView Then
        If iColor = vbWhite Then
            ColorPreview = vbBlack
        ElseIf iColor = vbBlack Then
            ColorPreview = vbWhite
        ElseIf iColor = vbYellow Then
            ColorPreview = vbRed
        Else
            ColorPreview = iColor
        End If
    Else
        ColorPreview = iColor
    End If
End Function



Private Sub pa_GotFocus()
    AtualizaGrafico
End Sub

Private Sub pa_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then gPan 0, 1
    If KeyCode = vbKeyDown Then gPan 0, -1
    If KeyCode = vbKeyLeft Then gPan -1, 0
    If KeyCode = vbKeyRight Then gPan 1, 0
    
    If KeyCode = 107 Or KeyCode = vbKeyPageUp Then gZoom 1.1
    If KeyCode = 109 Or KeyCode = vbKeyPageDown Then gZoom 0.9
        
    If KeyCode = vbKeyP Then togglePrintView
    If KeyCode = vbKeyG Then toggleGrid
    If KeyCode = vbKeyM Then toggleMarkName
    If KeyCode = vbKeyS Then customScale
    If KeyCode = vbKeyF Then customOrigin
    If KeyCode = vbKeyF1 Then quickhelp

End Sub

Private Sub pa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mouseClickX = X
    mouseClickY = Y
    origemClickX = origemX
    origemClickY = origemY
    
    If Button > 0 Then dragging = True
    'If Button = 2 Then zooming = True
End Sub
Private Sub pa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mouseX = X
    mouseY = Y
    
    If dragging Then AtualizaGrafico
    'If zooming Then AtualizaGrafico
End Sub
Private Sub pa_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button > 0 Then dragging = False
    'If Button = 2 Then zooming = False
End Sub





