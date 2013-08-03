VERSION 5.00
Begin VB.Form frmDadosCartograficos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dados cartograficos"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7050
   Begin VB.TextBox txtSaida 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   3360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "frmFichaCartografica.frx":0000
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calcular"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Coordenadas iniciais"
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      Begin VB.TextBox txtUTMy 
         Height          =   285
         Left            =   600
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtUTMx 
         Height          =   285
         Left            =   600
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Y:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   885
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "X:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   525
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmDadosCartograficos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim cX As Double, cY As Double
    Dim auxStr As String, txt As String
    On Error GoTo erro
    
    If Len(txtUTMx.Text) < 6 Then GoTo erro
    If Len(txtUTMy.Text) < 7 Then GoTo erro
    
    UTM2LL cX, cY, strToDec(txtUTMx.Text), strToDec(txtUTMy.Text)
    
    CalculaAreaPerim
    
    GetParamData ""
    

    txtSaida.Text = "Latitude: " & vbCrLf & dec2dms(cX, 1) & vbCrLf & _
        vbCrLf & "Longitude: " & vbCrLf & dec2dms(cY, 2) & vbCrLf & _
        vbCrLf & "Convergência Meridiana: " & vbCrLf & dec2dms(ConvMeridiana(cX, cY), 4) & vbCrLf & _
        vbCrLf & "Fator de Escala K: " & vbCrLf & "K = " & decToStr(FatorK_UTM(strToDec(txtUTMx.Text), strToDec(txtUTMy.Text)), 8) & vbCrLf & _
        vbCrLf & "Área medida: " & vbCrLf & GetParamData("CAMPO_AREAMED") & " ha " & vbCrLf & _
        vbCrLf & "Perímetro: " & vbCrLf & GetParamData("CAMPO_PERIMETRO") & " m " & vbCrLf

        'vbCrLf & "K = " & decToStr(FatorK(cX, cY), 8)
        'vbCrLf & "Datum: " & vbCrLf & Datum.Nome & " - Zona " & Datum.Zona & Datum.Letra
    
    txtSaida.SetFocus
    Exit Sub
erro:
    MsgBox "Problema nos valores de entrada!", vbExclamation, "Calculadora de Georeferências"

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Label1.Caption = "Coordenadas do centro do mapa." & vbCrLf & vbCrLf & "Datum: " & datum.Nome & vbCrLf & _
            "Zona UTM: " & CStr(datum.Zona) & "/" & datum.Letra
    txtSaida.Text = ""
    txtUTMx.Text = decToStr(pImovel.CentroFolha.X, 0)
    txtUTMy.Text = decToStr(pImovel.CentroFolha.Y, 0)
End Sub
