VERSION 5.00
Object = "{1E6AA28E-85F9-4256-8088-C3AD8E29B80E}#1.0#0"; "SuperGrid.ocx"
Begin VB.Form frmConvCoordPolar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conversor de Coordenadas Polares"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   10200
   Begin VB.CommandButton Command3 
      Caption         =   "Importar do Levantamento"
      Height          =   615
      Left            =   3720
      TabIndex        =   10
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Coordenada Inicial UTM"
      Height          =   855
      Left            =   5520
      TabIndex        =   3
      Top             =   120
      Width           =   4455
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2880
         TabIndex        =   7
         Top             =   310
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Top             =   310
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Norte:"
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Este:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
   End
   Begin SuperGrid.SFGrid gridC 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5106
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Polar > Cartesiana <<<<"
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cartesiana > Polar   >>>>"
      Height          =   615
      Left            =   3720
      TabIndex        =   0
      Top             =   2280
      Width           =   1575
   End
   Begin SuperGrid.SFGrid grid 
      Height          =   2895
      Left            =   5400
      TabIndex        =   9
      Top             =   1080
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5106
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "frmConvCoordPolar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub Command2_Click()
    'Polar para Cartesiano
    Dim actX As Double, actY As Double, ang As Double
    clearTableC
    
    On Error GoTo erro
    
    actX = strToDec(Text1.Text)
    actY = strToDec(Text2.Text)
            
    For i = 1 To grid.NumRows
        If (grid.value(i, 1) = "" Or grid.value(i, 4) = "") Then Exit Sub
        
        gridC.value(i, 1) = decToStr(actX, 2)
        gridC.value(i, 2) = decToStr(actY, 2)
        
        ang = strToDec(grid.value(i, 1)) + (strToDec(grid.value(i, 2)) / 60) + (strToDec(grid.value(i, 3)) / 3600)
        actX = actX + (Sin(Round(ang, 5) * deg2rad) * strToDec(grid.value(i, 4)))
        actY = actY + (Cos(Round(ang, 5) * deg2rad) * strToDec(grid.value(i, 4)))
    Next i
    'A ultima é a primeira
    Exit Sub
erro:
    MsgBox "Problema na conversão das coordenadas", vbExclamation, "Erro"
End Sub

Private Sub Command1_Click()
    'Cartesiano para Polar
    Dim dx1 As Double, dx2 As Double, dy1 As Double, dy2 As Double
    Dim val As Double
    clearTable
    
    On Error GoTo erro
    
    Text1.Text = gridC.value(1, 1)
    Text2.Text = gridC.value(1, 2)
    
    For i = 1 To gridC.NumRows
        If (gridC.value(i, 1) = "" Or gridC.value(i, 2) = "") Then Exit Sub
        
        grid.value(i, 0) = CStr(i)
        
        dx1 = strToDec(gridC.value(i, 1))
        dy1 = strToDec(gridC.value(i, 2))
        
        If (gridC.value(i + 1, 1) <> "" Or gridC.value(i + 1, 2) <> "") Then
            dx2 = strToDec(gridC.value(i + 1, 1))
            dy2 = strToDec(gridC.value(i + 1, 2))
        Else
            dx2 = strToDec(gridC.value(1, 1))
            dy2 = strToDec(gridC.value(1, 2))
        End If
        
        val = Round(AnguloCoord(dx1, dy1, dx2, dy2) / deg2rad, 5)
        grid.value(i, 1) = CStr(Fix(val))
        grid.value(i, 2) = CStr(Abs(Fix(((val - Fix(val)) * 60))))
        grid.value(i, 3) = decToStr(Round(((val * 60) - Fix(val * 60)) * 60, 2), 2)
        grid.value(i, 4) = decToStr(Sqr((Abs(dy1 - dy2) ^ 2) + (Abs(dx1 - dx2) ^ 2)), 2)
    Next i
    Exit Sub
erro:
    MsgBox "Problema na conversão das coordenadas", vbExclamation, "Erro"
End Sub

Private Sub clearTable()
    grid.Clear
    grid.FixCols = 1
    grid.NumCols = 5
    grid.NumRows = 100
    grid.ColWidth(0) = 330
    grid.ColWidth(1) = 960
    grid.ColWidth(2) = 900
    grid.ColWidth(3) = 900
    grid.ColWidth(4) = 1030
    grid.value(0, 0) = "Nº"
    grid.value(0, 1) = "Azimute (º)"
    grid.value(0, 2) = "Minutos"
    grid.value(0, 3) = "Segundos"
    grid.value(0, 4) = "Distância (m)"
    
    For i = 1 To 99
        grid.value(i, 0) = i
    Next i
End Sub

Private Sub clearTableC()
    gridC.Clear
    gridC.FixCols = 1
    gridC.NumCols = 3
    gridC.NumRows = 100
    gridC.ColWidth(0) = 330
    gridC.ColWidth(1) = 1080
    gridC.ColWidth(2) = 1080
    gridC.value(0, 0) = "Nº"
    gridC.value(0, 1) = "Este"
    gridC.value(0, 2) = "Norte"
    
    For i = 1 To 99
        gridC.value(i, 0) = i
    Next i
End Sub

Private Sub Command3_Click()
    clearTableC
    For i = 1 To pImovel.NumCoords
        gridC.value(i, 0) = CStr(i)
        gridC.value(i, 1) = pImovel.dPoligono(i).X
        gridC.value(i, 2) = pImovel.dPoligono(i).Y
    Next i
End Sub

Private Sub Form_Load()
    Label3.Caption = "Converta aqui coordenadas cartesianas em polares e vice-e-versa." & vbCrLf & _
                        "  - Tabela da esquerda: Cartesiana/UTM" & vbCrLf & _
                        "  - Tabela da direita: Polar com coordenada inicial"
    clearTable
    clearTableC
End Sub
