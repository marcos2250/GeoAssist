VERSION 5.00
Object = "{1E6AA28E-85F9-4256-8088-C3AD8E29B80E}#1.0#0"; "SuperGrid.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin SuperGrid.SFGrid SFGrid1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6588
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    SFGrid1.Value(1, 1) = "aaa"
    MsgBox SFGrid1.Value(1, 1)
    SFGrid1.ColWidth(1) = 300
End Sub
