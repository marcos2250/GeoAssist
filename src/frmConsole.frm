VERSION 5.00
Begin VB.Form frmConsole 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Console"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   9495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TakeKeyIn Text2.Text
        Text1.Text = Text1.Text & vbCrLf & Text2.Text
        Text1.SelStart = Len(Text1.Text)
        Text2.Text = ""
        form1.AtualizaGrafico
    End If
End Sub
