VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDocumentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de Documentos"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10665
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   10665
   Begin VB.Frame Frame5 
      Height          =   1935
      Index           =   3
      Left            =   6600
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CheckBox chkReqIncra 
         Caption         =   "Requerimento p/ INCRA (Trabalhos de Georef.)"
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CheckBox chkDL 
         Caption         =   "Descrições Confrontações"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   2295
      End
      Begin VB.CheckBox chkOpcDl 
         Caption         =   "Estilo Tabela"
         Enabled         =   0   'False
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1935
      Index           =   2
      Left            =   3480
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CheckBox chkGEarth 
         Caption         =   "Exportar e exibir planta no Google Earth (imagem)"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   2295
      End
      Begin VB.CheckBox chkAutoCAD 
         Caption         =   "Gerar planta no AutoCAD"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   2295
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1935
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CheckBox chkCalcXLS 
         Caption         =   "Planilha de Coordenadas Geográficas"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CheckBox chkMemoDescr 
         Caption         =   "Memorial Descritivo (Escritura para Cartório)"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1935
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   3015
      Begin VB.CheckBox chkResArquivo 
         Caption         =   "Arquivo de Resultados"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CheckBox chkResPlanilha 
         Caption         =   "Planilha de Resultados"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   2295
      End
      Begin VB.CheckBox chkDeclara 
         Caption         =   "Declaração de Confrontantes"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   2535
      End
      Begin VB.CheckBox chkRelatorio 
         Caption         =   "Relatório Técnico"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
      Begin VB.CheckBox chkMemoDescrIncra 
         Caption         =   "Memorial Descritivo"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Gerar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   4260
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Georef."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Relatórios"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Graficos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Outros"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDL_Click()
    chkOpcDl.Enabled = CBool(chkDL.value)
End Sub

Private Function getSavePath(sFilter As String, sInitDir As String, sFileName As String) As String
   With frmMain.cmD1
      .FileName = sFileName
      .InitDir = sInitDir
      .Filter = sFilter
      .CancelError = True
      On Error GoTo cancel
      .ShowSave
      If .FileName <> "" Then
            If Dir$(.FileName) <> "" Then
                If MsgBox("O arquivo já existe. Deseja salvar assim mesmo?", vbQuestion + vbYesNo) = vbYes Then getSavePath = .FileName
            Else
                getSavePath = .FileName
            End If
      End If
      Exit Function
cancel:
   getSavePath = ""
   End With
End Function


Private Sub Command2_Click()
    Dim projName As String, strAux As String
    Dim nadaafazer As Boolean
    
    Dim listaComandos(0 To 10) As String
    Dim iCmd As Integer
    
    nadaafazer = True
    projName = LeftLastChr(prjFileName, ".")
        
    DefineDatum
    
    If prjDir = "" Or cfgUnsaved Then
        MsgBox "Salve o projeto antes!", vbExclamation
        Exit Sub
    End If
    
    iCmd = 0
        
    If CBool(chkMemoDescr.value) Then
        strAux = getSavePath("Documento do Word (*.doc)|*.doc", prjDir, "Memorial Escritura " & projName & ".doc")
        GeraRelatorio 1, strAux
        If strAux <> "" Then
            listaComandos(iCmd) = strAux
            iCmd = iCmd + 1
            nadaafazer = False
        End If
    End If
    
    If CBool(chkMemoDescrIncra.value) Then
        strAux = getSavePath("Documento do Word (*.doc)|*.doc", prjDir, "Memorial Georeferenciamento " & projName & ".doc")
        GeraMemoDescrIncra strAux
        If strAux <> "" Then
            listaComandos(iCmd) = strAux
            iCmd = iCmd + 1
            nadaafazer = False
        End If
    End If
        
    If CBool(chkGEarth.value) Then
        strAux = getSavePath("Google Earth (*.kml)|*.kml", prjDir, projName & ".kml")
        ExpGoogleEarthKml strAux
        If strAux <> "" Then
            listaComandos(iCmd) = strAux
            iCmd = iCmd + 1
            nadaafazer = False
        End If
    End If
    
    If CBool(chkAutoCAD.value) Then
        strAux = getSavePath("AutoCAD (*.dxf)|*.dxf", prjDir, projName & ".dxf")
        ExpAutoCadDXF strAux
        If strAux <> "" Then
            listaComandos(iCmd) = strAux
            iCmd = iCmd + 1
            nadaafazer = False
        End If
    End If
    
    If CBool(chkDL.value) Then
        strAux = getSavePath("Documento do Word (*.doc)|*.doc", prjDir, "Confrontacoes " & projName & ".doc")
        If CBool(chkOpcDl.value) Then
           'GeraRelatorio 4, prjDir
           GeraTabelaDL strAux
        Else
           GeraRelatorio 2, strAux
        End If
        If strAux <> "" Then
            listaComandos(iCmd) = strAux
            iCmd = iCmd + 1
            nadaafazer = False
        End If
    End If

    
    If CBool(chkReqIncra.value) Then
        strAux = getSavePath("Documento do Word (*.doc)|*.doc", prjDir, "RequerimentoINCRA.doc")
        GeraDocSimples App.Path & "\@Requerimento.xml", strAux
        If strAux <> "" Then
            listaComandos(iCmd) = strAux
            iCmd = iCmd + 1
            nadaafazer = False
        End If
    End If
    
    If CBool(chkRelatorio.value) Then
        strAux = getSavePath("Documento do Word (*.doc)|*.doc", prjDir, "RelatorioTecnico.doc")
        GeraDocSimples App.Path & "\@RelatorioTecnico.xml", strAux
        If strAux <> "" Then
            listaComandos(iCmd) = strAux
            iCmd = iCmd + 1
            nadaafazer = False
        End If
    End If

    If CBool(chkDeclara.value) Then
        strAux = getSavePath("Documento do Word (*.doc)|*.doc", prjDir, "DeclaracaoLimite.doc")
        GeraDocSimples App.Path & "\@DeclaracaoLimite.xml", strAux
        If strAux <> "" Then
            listaComandos(iCmd) = strAux
            iCmd = iCmd + 1
            nadaafazer = False
        End If
    End If

    If CBool(chkCalcXLS.value) Then
        strAux = getSavePath("Documento do Excel (*.xls)|*.xls", prjDir, "CalculoAnalitico.xls")
        CalculoAnalitico strAux, 1
        If strAux <> "" Then
            listaComandos(iCmd) = strAux
            iCmd = iCmd + 1
            nadaafazer = False
        End If
    End If
    
    If CBool(chkResPlanilha.value) Then
        strAux = getSavePath("Documento do Excel (*.xls)|*.xls", prjDir, "Planilha_" & GetParamData("CAMPO_NUMCPFCNPJ") & ".xls")
        CalculoAnalitico strAux, 2
        If strAux <> "" Then
            listaComandos(iCmd) = strAux
            iCmd = iCmd + 1
            nadaafazer = False
        End If
    End If
    
    If CBool(chkResArquivo.value) Then
        strAux = getSavePath("Arquivo Texto (*.txt)|*.txt", prjDir, "Planilha_" & GetParamData("CAMPO_NUMCPFCNPJ") & ".txt")
        ArquivoResultado strAux
        If strAux <> "" Then
            listaComandos(iCmd) = strAux
            iCmd = iCmd + 1
            nadaafazer = False
        End If
    End If
    
    
    If nadaafazer Then
        MsgBox "Nenhum documento foi gerado", vbExclamation
    Else
        If MsgBox("Processamento concluido" & vbCrLf & "Deseja abrir os arquivos gerados?", vbQuestion + vbYesNo) = vbYes Then
            Dim i As Integer
            For i = 0 To iCmd
             ShellExecute 0&, vbNullString, Chr(34) & listaComandos(i) & Chr(34), vbNullString, vbNullString, 0
            Next i
        End If
    End If

End Sub

Private Sub Form_Load()
    Frame5(1).Visible = False
    Frame5(2).Visible = False
    Frame5(3).Visible = False
    Frame5(1).Left = Frame5(0).Left
    Frame5(2).Left = Frame5(0).Left
    Frame5(3).Left = Frame5(0).Left
    Frame5(1).Top = Frame5(0).Top
    Frame5(2).Top = Frame5(0).Top
    Frame5(3).Top = Frame5(0).Top
    frmDocumentos.Height = 3015
    frmDocumentos.Width = 4530
End Sub

Private Sub TabStrip1_Click()
    chkMemoDescr.value = 0
    chkMemoDescrIncra.value = 0
    chkGEarth.value = 0
    chkAutoCAD.value = 0
    chkDL.value = 0
    chkOpcDl.value = 0
    chkReqIncra.value = 0
    chkRelatorio.value = 0
    chkDeclara.value = 0
    chkCalcXLS.value = 0
    chkResPlanilha.value = 0
    chkResArquivo.value = 0
    
    Frame5(0).Visible = False
    Frame5(1).Visible = False
    Frame5(2).Visible = False
    Frame5(3).Visible = False
    Frame5(TabStrip1.SelectedItem.Index - 1).Visible = True
End Sub
