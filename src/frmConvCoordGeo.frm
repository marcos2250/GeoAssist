VERSION 5.00
Object = "{1E6AA28E-85F9-4256-8088-C3AD8E29B80E}#1.0#0"; "SuperGrid.ocx"
Begin VB.Form frmConvCoordGeo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conversor de coordenadas geográficas"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   8790
   Begin VB.CommandButton Command3 
      Caption         =   "Datum"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<<"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>>"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   1920
      Width           =   735
   End
   Begin SuperGrid.SFGrid gUTM 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   6165
   End
   Begin SuperGrid.SFGrid gGeo 
      Height          =   3495
      Left            =   4320
      TabIndex        =   1
      Top             =   600
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6165
   End
   Begin VB.Label Label2 
      Caption         =   "Coordenadas Geográficas"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Coordenadas UTM"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmConvCoordGeo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub clearTableGeo()
    gGeo.Clear
    gGeo.FixCols = 1
    gGeo.NumCols = 3
    gGeo.NumRows = 100
    gGeo.ColWidth(0) = 330
    gGeo.ColWidth(1) = 1700
    gGeo.ColWidth(2) = 1700
    gGeo.value(0, 0) = "Nº"
    gGeo.value(0, 1) = "Latitude"
    gGeo.value(0, 2) = "Longitude"
    
    For i = 1 To 99
        gGeo.value(i, 0) = i
    Next i
End Sub

Private Sub clearTableUTM()
    gUTM.Clear
    gUTM.FixCols = 1
    gUTM.NumCols = 3
    gUTM.NumRows = 100
    gUTM.ColWidth(0) = 330
    gUTM.ColWidth(1) = 1200
    gUTM.ColWidth(2) = 1200
    gUTM.value(0, 0) = "Nº"
    gUTM.value(0, 1) = "Este"
    gUTM.value(0, 2) = "Norte"
    
    For i = 1 To 99
        gUTM.value(i, 0) = i
    Next i
End Sub


Private Sub Command1_Click()
    'UTM para GEO
    Dim nx As Double, ny As Double
    clearTableGeo
    
    On Error GoTo erro
    For i = 1 To gUTM.NumRows
        If (gUTM.value(i, 1) = "" Or gUTM.value(i, 2) = "") Then Exit Sub
        
        gGeo.value(i, 0) = CStr(i)
        UTM2LL nx, ny, strToDec(gUTM.value(i, 1)), strToDec(gUTM.value(i, 2))
        gGeo.value(i, 1) = dec2dms(nx, 1)
        gGeo.value(i, 2) = dec2dms(ny, 2)
    Next i
    Exit Sub
erro:
    MsgBox "Problema na conversão das coordenadas", vbExclamation, "Erro"
End Sub

Private Sub Command2_Click()
    'GEO para UTM
    Dim nx As Double, ny As Double, lx As Double, ly As Double
    clearTableUTM
    
    On Error GoTo erro
    
    For i = 1 To gUTM.NumRows
        If (gGeo.value(i, 1) = "" Or gGeo.value(i, 2) = "") Then Exit Sub
        
        gGeo.value(i, 0) = CStr(i)
        If isNumerico(gGeo.value(i, 1)) And isNumerico(gGeo.value(i, 2)) Then
            lx = strToDec(gGeo.value(i, 1))
            ly = strToDec(gGeo.value(i, 2))
        Else
            lx = dms2dec(gGeo.value(i, 1))
            ly = dms2dec(gGeo.value(i, 2))
        End If
        
        LL2UTM nx, ny, lx, ly
        gUTM.value(i, 1) = decToStr(nx, 2)
        gUTM.value(i, 2) = decToStr(ny, 2)
    Next i
    Exit Sub
erro:
    MsgBox "Problema na conversão das coordenadas", vbExclamation, "Erro"
    
End Sub

Private Sub Command3_Click()
    frmDatum.Show
End Sub

Private Sub Form_Load()
    clearTableGeo
    clearTableUTM
End Sub
