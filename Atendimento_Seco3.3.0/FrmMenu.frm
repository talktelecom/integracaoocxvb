VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{0DEF929B-0FAF-4D7F-BBE5-43AEDE022B88}#1.0#0"; "Atendimento_Control.ocx"
Begin VB.Form FrmMenu 
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11625
   Icon            =   "FrmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   11625
   Begin Atendimento_Control.Atendimento_Button Atendimento_Button1 
      Height          =   615
      Left            =   120
      TabIndex        =   47
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmMenu.frx":0442
      Alignment       =   0
      Caption         =   "At_Comm"
   End
   Begin VB.CommandButton Command6 
      Caption         =   "3"
      Height          =   495
      Left            =   3120
      TabIndex        =   46
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "2"
      Height          =   495
      Left            =   1560
      TabIndex        =   45
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "1"
      Height          =   495
      Left            =   120
      TabIndex        =   44
      Top             =   7800
      Width           =   1215
   End
   Begin VB.TextBox TxtRamal 
      Height          =   285
      Left            =   7440
      TabIndex        =   43
      Text            =   "218"
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton cmd_Recado_Especific 
      Caption         =   "Recado_Especific"
      Height          =   495
      Left            =   5280
      Picture         =   "FrmMenu.frx":045E
      TabIndex        =   42
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ShowAbout"
      Height          =   495
      Left            =   10200
      TabIndex        =   41
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox TxtCanal 
      Height          =   285
      Left            =   7440
      TabIndex        =   39
      Text            =   "PA_??"
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton cmdAtender_Especfic 
      Caption         =   "Atender_Especfic"
      Height          =   495
      Left            =   5280
      Picture         =   "FrmMenu.frx":08A0
      TabIndex        =   40
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MERDAA"
      Height          =   495
      Left            =   120
      TabIndex        =   38
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton CmdTransferencia_Agenda 
      Caption         =   "Transferencia_Agenda"
      Height          =   495
      Left            =   120
      Picture         =   "FrmMenu.frx":0CE2
      TabIndex        =   37
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton CmdDiscar_Agenda 
      Caption         =   "Discar Agenda"
      Height          =   495
      Left            =   2520
      TabIndex        =   36
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Ligação Corrente Externo"
      Height          =   255
      Left            =   2280
      TabIndex        =   35
      Top             =   6000
      Value           =   1  'Checked
      Width           =   3375
   End
   Begin VB.TextBox TxtCC 
      Height          =   285
      Left            =   1440
      TabIndex        =   34
      Text            =   "1216"
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton CmdCentroCusto 
      Caption         =   "Centro Custo"
      Height          =   495
      Left            =   120
      TabIndex        =   33
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Informacao Estatistica"
      Height          =   495
      Left            =   10200
      TabIndex        =   32
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txt2 
      Height          =   285
      Left            =   2280
      TabIndex        =   31
      Text            =   "1234"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Left            =   1440
      TabIndex        =   30
      Text            =   "218"
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton CmdDesligaExpecif 
      Caption         =   "Desliga_Especif"
      Height          =   495
      Left            =   10200
      TabIndex        =   29
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CheckBox ChkSem_Operadora 
      Caption         =   "Sem Operadora"
      Height          =   255
      Left            =   10080
      TabIndex        =   27
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CommandButton CmdCadOperadora 
      Caption         =   "CadOperadora"
      Height          =   495
      Left            =   10200
      TabIndex        =   26
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton CmdMessager 
      Caption         =   "Messager"
      Height          =   495
      Left            =   10200
      TabIndex        =   25
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton CmdLiberaPausa 
      Caption         =   "Libera Pausa"
      Height          =   495
      Left            =   10200
      TabIndex        =   24
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton CmdAgenteNaoDisponivel 
      Caption         =   "Agente Nao Disponivel"
      Height          =   495
      Left            =   10200
      TabIndex        =   23
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton CmdOperadora 
      Caption         =   "Operadora"
      Height          =   495
      Left            =   10200
      TabIndex        =   22
      Top             =   1440
      Width           =   1215
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   1  'Align Top
      Height          =   240
      Left            =   0
      TabIndex        =   17
      Top             =   255
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   423
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   10
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "0"
            TextSave        =   "0"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Text            =   "0"
            TextSave        =   "0"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1588
            MinWidth        =   1588
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1588
            MinWidth        =   1588
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel9 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel10 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "14:25"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar2 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   10
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "Ramal"
            TextSave        =   "Ramal"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1235
            MinWidth        =   1235
            Text            =   "ID"
            TextSave        =   "ID"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Último Status"
            TextSave        =   "Último Status"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "Usuário"
            TextSave        =   "Usuário"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1060
            MinWidth        =   1060
            Text            =   "Fila"
            TextSave        =   "Fila"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Text            =   "N\Atendidas"
            TextSave        =   "N\Atendidas"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Para ver suas ligações atendidas e não atendidas clique aqui"
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1588
            MinWidth        =   1588
            Text            =   "Recados"
            TextSave        =   "Recados"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Para recuperar recados clique aqui"
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1588
            MinWidth        =   1588
            Text            =   "Diálogos"
            TextSave        =   "Diálogos"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Para recuperar diálogos clique aqui"
         EndProperty
         BeginProperty Panel9 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "Siga - me"
            TextSave        =   "Siga - me"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Para transferir chamadas clique aqui"
         EndProperty
         BeginProperty Panel10 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Hora"
            TextSave        =   "Hora"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3735
      Left            =   1440
      TabIndex        =   6
      Top             =   960
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ForeColor       =   255
      BackColorBkg    =   -2147483648
      GridColor       =   16777215
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   3735
      Left            =   6240
      TabIndex        =   7
      Top             =   960
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ForeColor       =   -2147483646
      BackColorBkg    =   -2147483648
      GridColor       =   16777215
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
      Height          =   1815
      Left            =   1440
      TabIndex        =   21
      Top             =   2880
      Visible         =   0   'False
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      ForeColor       =   8421376
      BackColorBkg    =   -2147483648
      GridColor       =   16777215
      Appearance      =   0
   End
   Begin VB.CommandButton CmdGravaOn 
      Caption         =   "&Grava/Off"
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdVoltar 
      Caption         =   "Voltar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9000
      Picture         =   "FrmMenu.frx":1124
      TabIndex        =   14
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdAgenda 
      Caption         =   "Agenda"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4920
      Picture         =   "FrmMenu.frx":1566
      TabIndex        =   13
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdAtualiza 
      Caption         =   "Atualiza"
      Height          =   495
      Left            =   3720
      Picture         =   "FrmMenu.frx":19A8
      TabIndex        =   12
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdTransfere 
      Caption         =   "Transfere"
      Height          =   495
      Left            =   2520
      Picture         =   "FrmMenu.frx":1CB2
      TabIndex        =   11
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdConsulta 
      Caption         =   "Consulta"
      Height          =   495
      Left            =   1320
      Picture         =   "FrmMenu.frx":29A4
      TabIndex        =   10
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCaptura 
      Caption         =   "Captura"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      Picture         =   "FrmMenu.frx":2DE6
      TabIndex        =   5
      Top             =   4440
      Width           =   1100
   End
   Begin VB.CommandButton CmdDiscagem 
      Caption         =   "Discagem"
      Height          =   495
      Left            =   120
      Picture         =   "FrmMenu.frx":3228
      TabIndex        =   4
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton CmdDesliga 
      Caption         =   "Desliga"
      Height          =   495
      Left            =   120
      Picture         =   "FrmMenu.frx":366A
      TabIndex        =   3
      Top             =   3240
      Width           =   1100
   End
   Begin VB.CommandButton CmdEspera 
      Caption         =   "Espera"
      Height          =   495
      Left            =   120
      Picture         =   "FrmMenu.frx":3AAC
      TabIndex        =   2
      Top             =   2640
      Width           =   1100
   End
   Begin VB.CommandButton CmdTransferencia 
      Caption         =   "Transferencia"
      Height          =   495
      Left            =   120
      Picture         =   "FrmMenu.frx":3EF6
      TabIndex        =   1
      Top             =   2040
      Width           =   1100
   End
   Begin VB.CommandButton CmdIntervalo 
      Caption         =   "Intervalo"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      Picture         =   "FrmMenu.frx":4338
      TabIndex        =   0
      Top             =   1440
      Width           =   1100
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Height          =   1815
      Left            =   1440
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      ForeColor       =   8421376
      BackColorBkg    =   -2147483648
      GridColor       =   16777215
      Appearance      =   0
   End
   Begin VB.CommandButton CmdDiscar 
      Caption         =   "Discar"
      Height          =   495
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CheckBox chkinterno 
      Caption         =   "Interno"
      Height          =   195
      Left            =   6120
      TabIndex        =   9
      Top             =   5400
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.ComboBox Cmbdisca 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FrmMenu.frx":477A
      Left            =   6120
      List            =   "FrmMenu.frx":477C
      TabIndex        =   8
      Text            =   "34710016"
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label LblID_LIGACAO 
      Caption         =   "ID_LIGACAO"
      Height          =   255
      Left            =   1440
      TabIndex        =   28
      Top             =   4800
      Width           =   3495
   End
   Begin VB.Image Image3 
      Height          =   480
      Index           =   0
      Left            =   7800
      Picture         =   "FrmMenu.frx":477E
      Top             =   4560
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Talk Telecom"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   8520
      TabIndex        =   19
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Menu MnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu MnuSair 
         Caption         =   "&Sair"
      End
   End
   Begin VB.Menu MnuConfiguracoes 
      Caption         =   "Configurações"
      Visible         =   0   'False
      Begin VB.Menu MnuConferencia 
         Caption         =   "Confêrencia"
         Begin VB.Menu MnuArquivos 
            Caption         =   "Arquivos"
         End
      End
   End
   Begin VB.Menu MnuVolume 
      Caption         =   "Volume"
      Begin VB.Menu MnuControleVolume 
         Caption         =   "Controle de Volume"
      End
   End
   Begin VB.Menu MnuVoice_Mail 
      Caption         =   "Voice-Mail"
      Visible         =   0   'False
      Begin VB.Menu MnuMudar_Saudacao 
         Caption         =   "Mudar_Saudacao"
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEscutar_Saudacao 
         Caption         =   "Escutar Saudação"
      End
   End
   Begin VB.Menu MnuSiga_me 
      Caption         =   "Siga - me"
      Begin VB.Menu MnuTranferirChamadas 
         Caption         =   "Tranferir Chamadas"
      End
   End
   Begin VB.Menu MnuTalkMessager 
      Caption         =   "TalkMessager"
      Visible         =   0   'False
   End
   Begin VB.Menu MnuFax 
      Caption         =   "Fax"
      Visible         =   0   'False
      Begin VB.Menu MnuEnviarFax 
         Caption         =   "Enviar Fax"
      End
   End
   Begin VB.Menu MnuAjuda 
      Caption         =   "&Ajuda"
      Begin VB.Menu MnuSobre 
         Caption         =   "Sobre"
      End
   End
End
Attribute VB_Name = "FrmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iContarPa As Integer
Dim iContarGr As Integer
Dim iFlgTransferencia As Integer
Dim iContAtendeuPA As Integer
Dim iContInFilaGrupo As Integer
Dim iiCanal As String
Dim i As Integer
Dim j As Integer
Dim iFlg As Integer

Private Sub Atendimento_Button1_Click()
  
'  x = Atendimento_Button1.met_Ultimo_Nome_VOX
'   MsgBox x
  If Atendimento_Button1.Caption = "&Login" Then
    'Call Atendimento_Button1.met_Chamar_Tela_Logon("1")
    Call Atendimento_Button1.met_Logon_Automatico("227", "33", "1234", "192.168.1.93", "44900")
    'Call Atendimento_Button1.met_Logon_Automatico_Simple(txt1, txt2)
  Else
    Call Atendimento_Button1.met_Logoff
  End If
End Sub

Private Sub Atendimento_Button1_iRetornoAtendeuPA(iCanal As String, iNome As String, iFone As Variant, iLigouPara As Variant, iStatus As Variant, iDuracao As Variant, iQtde As Integer)
    
  If iFlgTransferencia = 0 Then
'    If FrmMenu.WindowState = 0 Then
'      FrmMenu.Height = 3800
'    End If
    FrmMenu.MSFlexGrid4.Visible = True
    FrmMenu.MSFlexGrid3.Visible = True
    FrmMenu.MSFlexGrid2.Visible = False
    FrmMenu.MSFlexGrid1.Visible = False
    iContAtendeuPA = iQtde
    
    Debug.Print iCanal, iNome, iFone, iLigouPara, iStatus, iDuracao, iQtde
    For i = 1 To iContAtendeuPA
      FrmMenu.MSFlexGrid3.Rows = iContAtendeuPA + 1
      If FrmMenu.MSFlexGrid3.TextMatrix(i, 2) = "" Then
        FrmMenu.MSFlexGrid3.TextMatrix(i, 0) = iCanal
        FrmMenu.MSFlexGrid3.TextMatrix(i, 1) = iNome
        FrmMenu.MSFlexGrid3.TextMatrix(i, 2) = iFone
        FrmMenu.MSFlexGrid3.TextMatrix(i, 3) = iLigouPara
        FrmMenu.MSFlexGrid3.TextMatrix(i, 4) = iStatus
        FrmMenu.MSFlexGrid3.TextMatrix(i, 5) = iDuracao
        Exit For
      Else
        FrmMenu.MSFlexGrid3.TextMatrix(FrmMenu.MSFlexGrid3.Rows - 1, 3) = iLigouPara
        FrmMenu.MSFlexGrid3.TextMatrix(FrmMenu.MSFlexGrid3.Rows - 1, 4) = iStatus
        FrmMenu.MSFlexGrid3.TextMatrix(FrmMenu.MSFlexGrid3.Rows - 1, 5) = iDuracao
        Exit For
      End If
    Next
    'CmdDiscagem.Enabled = False
    CmdDesliga.Enabled = True
    CmdTransferencia.Enabled = True
  End If
  If iFlgTransferencia = 1 Then
    FrmMenu.MSFlexGrid4.Visible = False
    FrmMenu.MSFlexGrid3.Visible = False
    FrmMenu.MSFlexGrid2.Visible = True
    FrmMenu.MSFlexGrid1.Visible = True
    If MSFlexGrid1.Rows = 2 Then
      Atendimento_Button1.met_Atualizar_Ramais
    End If
  End If
End Sub

Private Sub Atendimento_Button1_iRetornoBotoes(iCapLogoff As Variant)
  Atendimento_Button1.Caption = iCapLogoff
End Sub

Private Sub Atendimento_Button1_iRetornoChamadasNaoAtendidas(iNome As String, iInterno As String, iFone As String, iLigouPara As String, iData As String, iHora As String)
  FrmnAtendidas.MSFlexGrid1.Rows = FrmnAtendidas.MSFlexGrid1.Rows + 1
  FrmnAtendidas.MSFlexGrid1.TextMatrix(FrmnAtendidas.MSFlexGrid1.Rows - 2, 0) = Trim(iNome)
  FrmnAtendidas.MSFlexGrid1.TextMatrix(FrmnAtendidas.MSFlexGrid1.Rows - 2, 1) = Trim(iInterno)
  FrmnAtendidas.MSFlexGrid1.TextMatrix(FrmnAtendidas.MSFlexGrid1.Rows - 2, 2) = Trim(iFone)
  FrmnAtendidas.MSFlexGrid1.TextMatrix(FrmnAtendidas.MSFlexGrid1.Rows - 2, 3) = Trim(iLigouPara)
  FrmnAtendidas.MSFlexGrid1.TextMatrix(FrmnAtendidas.MSFlexGrid1.Rows - 2, 4) = Trim(iData)
  FrmnAtendidas.MSFlexGrid1.TextMatrix(FrmnAtendidas.MSFlexGrid1.Rows - 2, 5) = iHora
  j = j + 1
  FrmnAtendidas.Command2.Caption = "Não Atendidas " + Str(j)
  
End Sub

Private Sub Atendimento_Button1_iRetornoChamadasAtendidas(iNome As String, iInterno As String, iFone As String, iLigouPara As String, iData As String, iHora As String)
  FrmnAtendidas.MSFlexGrid2.Rows = FrmnAtendidas.MSFlexGrid2.Rows + 1
  FrmnAtendidas.MSFlexGrid2.TextMatrix(FrmnAtendidas.MSFlexGrid2.Rows - 2, 0) = Trim(iNome)
  FrmnAtendidas.MSFlexGrid2.TextMatrix(FrmnAtendidas.MSFlexGrid2.Rows - 2, 1) = Trim(iInterno)
  FrmnAtendidas.MSFlexGrid2.TextMatrix(FrmnAtendidas.MSFlexGrid2.Rows - 2, 2) = Trim(iFone)
  FrmnAtendidas.MSFlexGrid2.TextMatrix(FrmnAtendidas.MSFlexGrid2.Rows - 2, 3) = Trim(iLigouPara)
  FrmnAtendidas.MSFlexGrid2.TextMatrix(FrmnAtendidas.MSFlexGrid2.Rows - 2, 4) = Trim(iData)
  FrmnAtendidas.MSFlexGrid2.TextMatrix(FrmnAtendidas.MSFlexGrid2.Rows - 2, 5) = iHora
  i = i + 1
  FrmnAtendidas.Command1.Caption = "Atendidas " + Str(i)
End Sub

Private Sub Atendimento_Button1_iRetornoConsultaCliente(iTipo As Integer, iMessagem As String)
  MsgBox (iTipo & " " & iMessagem)
End Sub

Private Sub Atendimento_Button1_iRetornoDadosPower(Cod_Cli_Nome As Variant, Contato_Endereco As Variant, Cidade As Variant, Estado As Variant, CEP As Variant, Numero_Servico As Variant, Numero_Cliente As Variant, Cod_Campanha As Variant)
  Atendimento_Button1.met_Agente_Nao_Disponivel
End Sub

Private Sub Atendimento_Button1_iRetornoDesligado()
  Call Carrega_Flex
  FrmMenu.MSFlexGrid4.Visible = False
  FrmMenu.MSFlexGrid3.Visible = False
  FrmMenu.MSFlexGrid2.Visible = True
  FrmMenu.MSFlexGrid1.Visible = True
  'FrmMenu.Cmbdisca.Text = ""
  'If FrmMenu.WindowState = 0 Then
  '  FrmMenu.Height = 2220
  'End If
  iFlgTransferencia = 0
  CmdTransferencia.Enabled = False
  CmdDesliga.Enabled = False
  CmdDiscagem.Enabled = True
  'CmdVoltar.Enabled = False
  iFlg = 0
  Atendimento_Button1.met_Libera_Pausa
End Sub

Private Sub Atendimento_Button1_iRetornoDialSliderMax(iMaximoSlider As Integer)
  FrmEscutaDial.Slider1.Max = iMaximoSlider
End Sub

Private Sub Atendimento_Button1_iRetornoDialSliderPos(iPosicaoSlider As Integer)
  FrmEscutaDial.Slider1.Value = iPosicaoSlider
End Sub

Private Sub Atendimento_Button1_iRetornoEncerrarConf(idTipo As Integer)
  Dim v As Integer
  v = 0
End Sub

Private Sub Atendimento_Button1_iRetornoEspera(iStatus As Integer)
  If iStatus = 1 Then
    CmdEspera.Caption = "Tira Espera"
  End If
  If iStatus = 0 Then
    CmdEspera.Caption = "Espera"
  End If
End Sub

Private Sub Atendimento_Button1_iRetornoGravaOnOff(iStatus As Integer)
  If iStatus = 1 Then
    'Gravaçao Habilitada
    CmdGravaOn.Caption = "&Grava/On"
  End If
End Sub

Private Sub Atendimento_Button1_iRetornoIDLIGACAO(iID_LIGACAO As String)
  LblID_LIGACAO.Caption = iID_LIGACAO
End Sub

Private Sub Atendimento_Button1_iRetornoInFilaGrupo(iCanal As String, iNome As String, iFone As Variant, iLigouPara As Variant, iStatus As Variant, iDuracao As Variant, iQtde As Integer)
  'If FrmMenu.WindowState = 0 Then FrmMenu.Height = 5250
  
  FrmMenu.MSFlexGrid1.Visible = False
  FrmMenu.MSFlexGrid2.Visible = False
  FrmMenu.MSFlexGrid3.Visible = True
  FrmMenu.MSFlexGrid4.Visible = True
  iContInFilaGrupo = iQtde
  For i = 1 To iContInFilaGrupo
    FrmMenu.MSFlexGrid4.Rows = iContInFilaGrupo + 1
    If FrmMenu.MSFlexGrid4.TextMatrix(i, 3) = "" Then
      FrmMenu.MSFlexGrid4.TextMatrix(i, 0) = iCanal
      FrmMenu.MSFlexGrid4.TextMatrix(i, 1) = iNome
      FrmMenu.MSFlexGrid4.TextMatrix(i, 2) = iFone
      FrmMenu.MSFlexGrid4.TextMatrix(i, 3) = iLigouPara
      FrmMenu.MSFlexGrid4.TextMatrix(i, 4) = iStatus
      FrmMenu.MSFlexGrid4.TextMatrix(i, 5) = iDuracao
      Exit For
    End If
    If FrmMenu.MSFlexGrid4.TextMatrix(i, 3) = iLigouPara Then
      FrmMenu.MSFlexGrid4.TextMatrix(i, 4) = iStatus
      FrmMenu.MSFlexGrid4.TextMatrix(i, 5) = iDuracao
      Exit For
    End If
  Next
End Sub


Private Sub Atendimento_Button1_iRetornoInfoDial(iRamal As String, iDataHora As String, iDuracao As String, iFone As String, iLigouPara As String, iMensagem As String, iComentario As String)
  FrmEscutaDial.MSFlexGrid1.Rows = FrmEscutaDial.MSFlexGrid1.Rows + 1
  FrmEscutaDial.MSFlexGrid1.TextMatrix(FrmEscutaDial.MSFlexGrid1.Rows - 2, 0) = Trim(iRamal)
  FrmEscutaDial.MSFlexGrid1.TextMatrix(FrmEscutaDial.MSFlexGrid1.Rows - 2, 1) = Trim(iDataHora)
  FrmEscutaDial.MSFlexGrid1.TextMatrix(FrmEscutaDial.MSFlexGrid1.Rows - 2, 2) = Trim(iDuracao)
  FrmEscutaDial.MSFlexGrid1.TextMatrix(FrmEscutaDial.MSFlexGrid1.Rows - 2, 3) = Trim(iFone)
  FrmEscutaDial.MSFlexGrid1.TextMatrix(FrmEscutaDial.MSFlexGrid1.Rows - 2, 4) = Trim(iLigouPara)
  FrmEscutaDial.MSFlexGrid1.TextMatrix(FrmEscutaDial.MSFlexGrid1.Rows - 2, 5) = iMensagem
  FrmEscutaDial.MSFlexGrid1.TextMatrix(FrmEscutaDial.MSFlexGrid1.Rows - 2, 6) = iComentario
  'M0126_16042002_113209.vox
End Sub

Private Sub Atendimento_Button1_iRetornoInfoRec(iMensagem As String, iData As String, iHora As String, iDuracao As String, iNumero As String)
  FrmEscutaRec.MSFlexGrid1.Rows = FrmEscutaRec.MSFlexGrid1.Rows + 1
  FrmEscutaRec.MSFlexGrid1.TextMatrix(FrmEscutaRec.MSFlexGrid1.Rows - 2, 0) = iMensagem
  FrmEscutaRec.MSFlexGrid1.TextMatrix(FrmEscutaRec.MSFlexGrid1.Rows - 2, 1) = Trim(iData)
  FrmEscutaRec.MSFlexGrid1.TextMatrix(FrmEscutaRec.MSFlexGrid1.Rows - 2, 2) = Trim(iHora)
  FrmEscutaRec.MSFlexGrid1.TextMatrix(FrmEscutaRec.MSFlexGrid1.Rows - 2, 3) = Trim(iDuracao)
  FrmEscutaRec.MSFlexGrid1.TextMatrix(FrmEscutaRec.MSFlexGrid1.Rows - 2, 4) = Trim(iNumero)
End Sub

Private Sub Atendimento_Button1_iRetornoInformacaoEstatistica()
  MsgBox "Ok"
End Sub

Private Sub Atendimento_Button1_iRetornoLogon(iStrRamal As Variant, iStrID As Variant, iStrUltimoStatus As Variant, iStrUsuario As Variant, iStrFila As Variant, iStrNAtendidas As Variant, iStrRecados As Variant, iStrDialogos As Variant, iStrSigaMe As Variant)
  FrmMenu.StatusBar1.Panels(1).Text = iStrRamal
  FrmMenu.StatusBar1.Panels(2).Text = iStrID
  FrmMenu.StatusBar1.Panels(3).Text = iStrUltimoStatus
  FrmMenu.StatusBar1.Panels(4).Text = iStrUsuario
  FrmMenu.StatusBar1.Panels(5).Text = iStrFila
  FrmMenu.StatusBar1.Panels(6).Text = iStrNAtendidas
  FrmMenu.StatusBar1.Panels(7).Text = iStrRecados
  FrmMenu.StatusBar1.Panels(8).Text = iStrDialogos
  FrmMenu.StatusBar1.Panels(9).Text = iStrSigaMe
  CmdGravaOn.Caption = "&Grava/Off"
  If iStrUltimoStatus = "Logon Válido" Then
    CmdDiscagem.Enabled = True
  End If
  CmdIntervalo.Enabled = True
End Sub

Private Sub Atendimento_Button1_iRetornoMessager(iRamal As String, iMessagem As String, iData As String, iHora As String)
  'MsgBox (iRamal + " " + iMessagem + " " + iData + " " + iHora)
End Sub


Private Sub Atendimento_Button1_iRetornoNumeroCliente(iTipo As Integer, iID As String, iNome As String, iNumeroA As String, iNumeroB As String)
  If iFlg = 0 Then
    If iTipo = 1 Then Atendimento_Button1.met_Agente_Nao_Disponivel
    iFlg = 1
  End If
End Sub

Private Sub Atendimento_Button1_iRetornoOutFilaGrupo(iRamal As String)
  'If FrmMenu.MSFlexGrid3.Rows = 2 Then
  '  FrmMenu.Height = 2220
  'End If
  For i = 1 To MSFlexGrid4.Rows
    If FrmMenu.MSFlexGrid4.TextMatrix(i, 3) = iRamal Then
      FrmMenu.MSFlexGrid4.TextMatrix(i, 0) = ""
      FrmMenu.MSFlexGrid4.TextMatrix(i, 1) = ""
      FrmMenu.MSFlexGrid4.TextMatrix(i, 2) = ""
      FrmMenu.MSFlexGrid4.TextMatrix(i, 3) = ""
      FrmMenu.MSFlexGrid4.TextMatrix(i, 4) = ""
      FrmMenu.MSFlexGrid4.TextMatrix(i, 5) = ""
      Exit For
    End If
  Next
End Sub

Private Sub Atendimento_Button1_iRetornoOutFilaRamal(iRamal As String)
  'If FrmMenu.MSFlexGrid3.Rows = 2 Then
  '  FrmMenu.Height = 2220
  'End If
End Sub

Private Sub Atendimento_Button1_iRetornoRamaisGR(iRamal As String, iNome As String, iNumeroPA As String)
  FrmMenu.MSFlexGrid2.Rows = FrmMenu.MSFlexGrid2.Rows + 1
  FrmMenu.MSFlexGrid2.TextMatrix(FrmMenu.MSFlexGrid2.Rows - 2, 0) = iRamal
  FrmMenu.MSFlexGrid2.TextMatrix(FrmMenu.MSFlexGrid2.Rows - 2, 1) = iNome
  FrmMenu.MSFlexGrid2.TextMatrix(FrmMenu.MSFlexGrid2.Rows - 2, 2) = iNumeroPA
End Sub

Private Sub Atendimento_Button1_iRetornoRamaisPA(iRamal As String, iOperador As String, iStatus As String)
  FrmMenu.MSFlexGrid1.Rows = FrmMenu.MSFlexGrid1.Rows + 1
  FrmMenu.MSFlexGrid1.TextMatrix(FrmMenu.MSFlexGrid1.Rows - 2, 0) = iRamal
  FrmMenu.MSFlexGrid1.TextMatrix(FrmMenu.MSFlexGrid1.Rows - 2, 1) = iOperador
  FrmMenu.MSFlexGrid1.TextMatrix(FrmMenu.MSFlexGrid1.Rows - 2, 2) = iStatus
  
  CmdVoltar.Enabled = True
 ' CmdDiscagem.Enabled = False
End Sub



Private Sub Atendimento_Button1_iRetornoRecSliderMax(iMaximoSlider As Integer)
  FrmEscutaRec.Slider1.Max = iMaximoSlider
End Sub

Private Sub Atendimento_Button1_iRetornoRecSliderPos(iPosicaoSlider As Integer)
  FrmEscutaRec.Slider1.Value = iPosicaoSlider
End Sub

Private Sub Atendimento_Button1_iRetornoRetornoConsulta()
  Call Carrega_Flex
  FrmMenu.MSFlexGrid4.Visible = False
  FrmMenu.MSFlexGrid3.Visible = False
  FrmMenu.MSFlexGrid2.Visible = True
  FrmMenu.MSFlexGrid1.Visible = True
  FrmMenu.Cmbdisca.Text = ""
  'If FrmMenu.WindowState = 0 Then
  '  FrmMenu.Height = 2220
  'End If
  iFlgTransferencia = 0
  CmdTransferencia.Enabled = True
  CmdDesliga.Enabled = False
  CmdDiscagem.Enabled = True
  'CmdVoltar.Enabled = False
  iFlg = 0
End Sub

Private Sub Atendimento_Button1_iRetornoRetornoForadoGancho()
  MsgBox "Fora do Gancho"
End Sub

Private Sub Atendimento_Button1_iRetornoRetornoNoGancho()
  MsgBox "No Gancho"
End Sub

Private Sub Atendimento_Button1_iRetornoStatusConf(iRamalTelefoneArquivo As Variant, iNome As Variant, iStatus As Variant, iCanal As Variant, iID As Variant)
  Debug.Print iRamalTelefoneArquivo & "," & iNome & "," & iStatus & "," & iCanal & "," & iID
End Sub

Private Sub Atendimento_Button1_iRetornoStatusGeral(iStatus As String)
  
  FrmMenu.StatusBar1.Panels(3).Text = iStatus
  If iStatus = "Livre" Then
    CmdVoltar.Enabled = False
    CmdDiscagem.Enabled = True
  End If
End Sub

Private Sub Atendimento_Button1_iRetornoValueAgenteNaoDisponivel(iValue As Integer)
If iValue = 2 Then
  FrmMenu.StatusBar1.Panels(3).Text = "Não Disponivel"
End If
End Sub

Private Sub Atendimento_Button1_iRetornoValueIntervalo(iValue As Integer)
  If iValue = 0 Then FrmMenu.StatusBar1.Panels(3).Text = "Livre"
  If iValue = 1 Then FrmMenu.StatusBar1.Panels(3).Text = "Intervalo"
End Sub

Private Sub Atendimento_Button2_Click()

End Sub

Private Sub Cmbdisca_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    Call CmdDiscar_Click
  End If
End Sub


Private Sub cmd_Recado_Especific_Click()
  Call Atendimento_Button1.met_Recado_Especific(TxtRamal)
End Sub

Private Sub CmdAgenteNaoDisponivel_Click()
  Atendimento_Button1.met_Agente_Nao_Disponivel
End Sub

Private Sub CmdAtualiza_Click()
  Call Carrega_Flex
  Atendimento_Button1.met_Atualizar_Ramais
End Sub

Private Sub CmdCadOperadora_Click()
  Atendimento_Button1.met_Cadastro_Operadora
End Sub

Private Sub cmdAtender_Especfic_Click()
  Call Atendimento_Button1.met_Atender_Especific(TxtCanal)
End Sub

Private Sub CmdCaptura_Click()
  Atendimento_Button1.met_Captura
End Sub

Private Sub CmdConferencia_Click()
  Call Atendimento_Button1.met_Conferencia(0)
End Sub

Private Sub CmdCentroCusto_Click()
  If (Check1.Value = 1) Then
    'Ligacao Corrente Externo
    'Call Atendimento_Button1.met_Setar_CentroCusto(1, TxtCC.Text)
  Else
    'Ligacao Externa muda o centro de custo do ramal
    'Call Atendimento_Button1.met_Setar_CentroCusto(0, TxtCC.Text)
  End If
End Sub

Private Sub CmdConsulta_Click()
  If CmdConsulta.Caption = "Consulta" Then
    Call Atendimento_Button1.met_Consulta(Cmbdisca.Text, chkinterno.Value)
    CmdConsulta.Caption = "Retorna"
  Else
    Call Atendimento_Button1.met_Retorna(Cmbdisca.Text, chkinterno.Value)
    CmdConsulta.Caption = "Consulta"
  End If
End Sub

Private Sub CmdDesliga_Click()
  Atendimento_Button1.met_Desliga
  Atendimento_Button1.met_Libera_Pausa
  FrmMenu.MSFlexGrid4.Visible = False
  FrmMenu.MSFlexGrid3.Visible = False
  FrmMenu.MSFlexGrid2.Visible = True
  FrmMenu.MSFlexGrid1.Visible = True
  Carrega_Flex
  'FrmMenu.Height = 2220
  CmdDiscagem.Enabled = True
  iFlg = 0
End Sub

Private Sub CmdDesligaExpecif_Click()
  Atendimento_Button1.met_Desliga_Especific (iiCanal)
End Sub

Private Sub CmdDiscagem_Click()
  Atendimento_Button1.met_Discagem
  CmdDiscar.Enabled = True
  'FrmMenu.Height = 6465
End Sub

Private Sub CmdDiscar_Agenda_Click()
  Call Atendimento_Button1.met_Discar_Agenda(Cmbdisca.Text)
End Sub

Private Sub CmdDiscar_Click()
  CmdDiscar.Enabled = False
  FrmMenu.MSFlexGrid4.Visible = True
  FrmMenu.MSFlexGrid3.Visible = True
  FrmMenu.MSFlexGrid2.Visible = False
  FrmMenu.MSFlexGrid1.Visible = False
  'Call Atendimento_Button1.met_Setar_CentroCusto(1, "1111", Empty)
  Call Atendimento_Button1.met_Discar(Cmbdisca.Text, chkinterno.Value)
End Sub

Private Sub CmdEspera_Click()
  Atendimento_Button1.met_Espera
End Sub


Private Sub CmdIntervalo_Click()
  Atendimento_Button1.met_Intervalo
End Sub

Private Sub CmdLiberaPausa_Click()
  Call Atendimento_Button1.met_Libera_Pausa
End Sub

Private Sub CmdMessager_Click()
  Dim iMensagem As String, iRamal As String
  iMensagem = InputBox("Messagem Ramal", "TalkMessager")
  iRamal = InputBox("Ramal Destino", "TalkMessager")
  Call Atendimento_Button1.met_Enviar_Mensagem(iMensagem, iRamal)

End Sub

Private Sub CmdOperadora_Click()
  Dim iRetorno As Integer
  'Sem Operadora
  iRetorno = Atendimento_Button1.met_Cadastro_Operadora_Param("-1", "011", "0", "0", "0", True, False)
  'Com Operadora
  'iRetorno = Atendimento_Button1.met_Cadastro_Operadora_Param("-1", "", "11", "15", "15", False, False)
  
  'Sem Operadora
  'iRetorno = Atendimento_Button1.met_Cadastro_Operadora_Param("-1", "0", "0", "0", "0", True, False)
  
  
  Select Case iRetorno
    Case -1
      LblID_LIGACAO.Caption = "Não Conectado."
    Case 0
      LblID_LIGACAO.Caption = "Falhou."
    Case 1
      LblID_LIGACAO.Caption = "Sucesso."
  End Select
  'Call Atendimento_Button1.met_Cadastro_Operadora

End Sub

Private Sub CmdTransfere_Click()
  iFlgTransferencia = 0
  Call Atendimento_Button1.met_Transfere_Ligacao(Cmbdisca.Text, chkinterno.Value)
  CmdTransfere.Enabled = False
  CmdDesliga.Enabled = False
  CmdVoltar.Enabled = False
  CmdVoltar.Enabled = False
  'FrmMenu.Height = 2220
End Sub

Private Sub CmdTransferencia_Agenda_Click()
  Atendimento_Button1.met_Transferencia_Agenda
  iFlgTransferencia = 1
  CmdTransferencia.Enabled = False
  CmdTransfere.Enabled = True
  CmdVoltar.Enabled = True
  'FrmMenu.Height = 6465

End Sub

Private Sub CmdTransferencia_Click()
  Atendimento_Button1.met_Transferencia
  iFlgTransferencia = 1
  CmdTransferencia.Enabled = False
  CmdTransfere.Enabled = True
  'FrmMenu.Height = 6465
End Sub

Private Sub CmdVoltar_Click()
  'Call Carrega_Flex
  Atendimento_Button1.met_Voltar_Ligacao
  'FrmMenu.Height = 2220
  CmdDiscar.Enabled = False
End Sub


Private Sub Command2_Click()
Call Atendimento_Button1.met_Transferencia_Agenda
Call Atendimento_Button1.met_Consulta("91953907", 0)

End Sub

Private Sub Command3_Click()
  Atendimento_Button1.met_ShowAboutBox
End Sub

Private Sub Command4_Click()
  Call Atendimento_Button1.met_Conferencia(0)
End Sub

Private Sub Command1_Click()
  Call Atendimento_Button1.met_Informacao_Estatistica("11111", "2", "Teste")
End Sub

Private Sub Command5_Click()
  Call Atendimento_Button1.met_Adicionar_Conf("236", 0)
  Call Atendimento_Button1.met_Adicionar_Conf("216", 0)
  Call Atendimento_Button1.met_Iniciar_Conf(0)
End Sub

Private Sub Command6_Click()
  Call Atendimento_Button1.met_Encerrar_Conf(0)
End Sub

Private Sub Form_Load()
  Caption = "Atendimento OCX 1.7.00b " & Format(Now, "DD.MM.YYYY HH:MM:SS")
  Carrega_Flex
  CmdGravaOn.BackColor = &H8000000F
  Atendimento_Button1.Caption = "&Login"
  'FrmMenu.Height = 2220
  'Call Atendimento_Button1.met_ShowAboutBox
End Sub
Sub Carrega_Flex()
  FrmMenu.MSFlexGrid4.Visible = False
  FrmMenu.MSFlexGrid3.Visible = False
  FrmMenu.MSFlexGrid2.Visible = True
  FrmMenu.MSFlexGrid1.Visible = True
  
  FrmMenu.MSFlexGrid1.Clear
  FrmMenu.MSFlexGrid1.TextMatrix(0, 0) = "Ramal"
  FrmMenu.MSFlexGrid1.TextMatrix(0, 1) = "Operador"
  FrmMenu.MSFlexGrid1.TextMatrix(0, 2) = "Status"
  FrmMenu.MSFlexGrid1.ColWidth(0) = 700
  FrmMenu.MSFlexGrid1.ColWidth(1) = 1800
  FrmMenu.MSFlexGrid1.ColWidth(2) = 1800
  FrmMenu.MSFlexGrid1.Rows = 2
  
  FrmMenu.MSFlexGrid2.Clear
  FrmMenu.MSFlexGrid2.TextMatrix(0, 0) = "Ramal"
  FrmMenu.MSFlexGrid2.TextMatrix(0, 1) = "Nome"
  FrmMenu.MSFlexGrid2.TextMatrix(0, 2) = "Nr Pa"
  FrmMenu.MSFlexGrid2.ColWidth(0) = 700
  FrmMenu.MSFlexGrid2.ColWidth(1) = 1800
  FrmMenu.MSFlexGrid2.ColWidth(2) = 800
  FrmMenu.MSFlexGrid2.Rows = 2

  FrmMenu.MSFlexGrid3.Clear
  FrmMenu.MSFlexGrid3.TextMatrix(0, 0) = "Canal"
  FrmMenu.MSFlexGrid3.TextMatrix(0, 1) = "Nome"
  FrmMenu.MSFlexGrid3.TextMatrix(0, 2) = "Fone"
  FrmMenu.MSFlexGrid3.TextMatrix(0, 3) = "LigouPara"
  FrmMenu.MSFlexGrid3.TextMatrix(0, 4) = "Status"
  FrmMenu.MSFlexGrid3.TextMatrix(0, 5) = "Duracao"
  FrmMenu.MSFlexGrid3.ColWidth(0) = 600
  FrmMenu.MSFlexGrid3.ColWidth(1) = 1800
  FrmMenu.MSFlexGrid3.ColWidth(2) = 1800
  FrmMenu.MSFlexGrid3.ColWidth(3) = 1800
  FrmMenu.MSFlexGrid3.ColWidth(4) = 1500
  FrmMenu.MSFlexGrid3.ColWidth(5) = 1000
  FrmMenu.MSFlexGrid3.Rows = 2

  FrmMenu.MSFlexGrid4.Clear
  FrmMenu.MSFlexGrid4.TextMatrix(0, 0) = "Canal"
  FrmMenu.MSFlexGrid4.TextMatrix(0, 1) = "Nome"
  FrmMenu.MSFlexGrid4.TextMatrix(0, 2) = "Fone"
  FrmMenu.MSFlexGrid4.TextMatrix(0, 3) = "LigouPara"
  FrmMenu.MSFlexGrid4.TextMatrix(0, 4) = "Status"
  FrmMenu.MSFlexGrid4.TextMatrix(0, 5) = "Duracao"
  FrmMenu.MSFlexGrid4.ColWidth(0) = 600
  FrmMenu.MSFlexGrid4.ColWidth(1) = 1800
  FrmMenu.MSFlexGrid4.ColWidth(2) = 1800
  FrmMenu.MSFlexGrid4.ColWidth(3) = 1800
  FrmMenu.MSFlexGrid4.ColWidth(4) = 1500
  FrmMenu.MSFlexGrid4.ColWidth(5) = 1000
  FrmMenu.MSFlexGrid4.Rows = 2
  
  FrmEscutaDial.MSFlexGrid1.Clear
  FrmEscutaDial.MSFlexGrid1.TextMatrix(0, 0) = "Ramal"
  FrmEscutaDial.MSFlexGrid1.TextMatrix(0, 1) = "DataHora"
  FrmEscutaDial.MSFlexGrid1.TextMatrix(0, 2) = "Duracao"
  FrmEscutaDial.MSFlexGrid1.TextMatrix(0, 3) = "Fone"
  FrmEscutaDial.MSFlexGrid1.TextMatrix(0, 4) = "LigouPara"
  FrmEscutaDial.MSFlexGrid1.TextMatrix(0, 5) = "Mensagem"
  FrmEscutaDial.MSFlexGrid1.TextMatrix(0, 6) = "Comentário"
  FrmEscutaDial.MSFlexGrid1.ColWidth(0) = 600
  FrmEscutaDial.MSFlexGrid1.ColWidth(1) = 1800
  FrmEscutaDial.MSFlexGrid1.ColWidth(2) = 1000
  FrmEscutaDial.MSFlexGrid1.ColWidth(3) = 1000
  FrmEscutaDial.MSFlexGrid1.ColWidth(4) = 1800
  FrmEscutaDial.MSFlexGrid1.ColWidth(5) = 2300
  FrmEscutaDial.MSFlexGrid1.ColWidth(6) = 1800
  FrmEscutaDial.MSFlexGrid1.Rows = 2
  
  FrmEscutaRec.MSFlexGrid1.Clear
  FrmEscutaRec.MSFlexGrid1.TextMatrix(0, 0) = "Mensagem"
  FrmEscutaRec.MSFlexGrid1.TextMatrix(0, 1) = "Data"
  FrmEscutaRec.MSFlexGrid1.TextMatrix(0, 2) = "Hora"
  FrmEscutaRec.MSFlexGrid1.TextMatrix(0, 3) = "Duracao"
  FrmEscutaRec.MSFlexGrid1.TextMatrix(0, 4) = "Numero"
  FrmEscutaRec.MSFlexGrid1.ColWidth(0) = 2000
  FrmEscutaRec.MSFlexGrid1.ColWidth(1) = 800
  FrmEscutaRec.MSFlexGrid1.ColWidth(2) = 800
  FrmEscutaRec.MSFlexGrid1.ColWidth(3) = 1000
  FrmEscutaRec.MSFlexGrid1.ColWidth(4) = 1000
  FrmEscutaRec.MSFlexGrid1.Rows = 2
  
  FrmnAtendidas.MSFlexGrid1.Clear
  FrmnAtendidas.MSFlexGrid1.TextMatrix(0, 0) = "Nome"
  FrmnAtendidas.MSFlexGrid1.TextMatrix(0, 1) = "Interno"
  FrmnAtendidas.MSFlexGrid1.TextMatrix(0, 2) = "Fone"
  FrmnAtendidas.MSFlexGrid1.TextMatrix(0, 3) = "Ligou Papa"
  FrmnAtendidas.MSFlexGrid1.TextMatrix(0, 4) = "Data"
  FrmnAtendidas.MSFlexGrid1.TextMatrix(0, 5) = "Hora"
  FrmnAtendidas.MSFlexGrid1.ColWidth(0) = 2000
  FrmnAtendidas.MSFlexGrid1.ColWidth(1) = 800
  FrmnAtendidas.MSFlexGrid1.ColWidth(2) = 800
  FrmnAtendidas.MSFlexGrid1.ColWidth(3) = 1000
  FrmnAtendidas.MSFlexGrid1.ColWidth(4) = 1000
  FrmnAtendidas.MSFlexGrid1.ColWidth(5) = 1000
  FrmnAtendidas.MSFlexGrid1.Rows = 2
  
  FrmnAtendidas.MSFlexGrid2.Clear
  FrmnAtendidas.MSFlexGrid2.TextMatrix(0, 0) = "Nome"
  FrmnAtendidas.MSFlexGrid2.TextMatrix(0, 1) = "Interno"
  FrmnAtendidas.MSFlexGrid2.TextMatrix(0, 2) = "Fone"
  FrmnAtendidas.MSFlexGrid2.TextMatrix(0, 3) = "Ligou Papa"
  FrmnAtendidas.MSFlexGrid2.TextMatrix(0, 4) = "Data"
  FrmnAtendidas.MSFlexGrid2.TextMatrix(0, 5) = "Hora"
  FrmnAtendidas.MSFlexGrid2.ColWidth(0) = 2000
  FrmnAtendidas.MSFlexGrid2.ColWidth(1) = 800
  FrmnAtendidas.MSFlexGrid2.ColWidth(2) = 800
  FrmnAtendidas.MSFlexGrid2.ColWidth(3) = 1000
  FrmnAtendidas.MSFlexGrid2.ColWidth(4) = 1000
  FrmnAtendidas.MSFlexGrid2.ColWidth(5) = 1000
  FrmnAtendidas.MSFlexGrid2.Rows = 2
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub MnuSair_Click()
  Unload Me
End Sub


Private Sub MnuSobre_Click()
  Atendimento_Button1.met_ShowAboutBox
End Sub

Private Sub MSFlexGrid1_Click()
  Cmbdisca.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0)
End Sub

Private Sub MSFlexGrid2_Click()
  Cmbdisca.Text = MSFlexGrid2.TextMatrix(MSFlexGrid1.RowSel, 0)
End Sub

Private Sub CmdGravaOn_Click()
  Atendimento_Button1.met_GravarOn
End Sub

Private Sub MSFlexGrid3_Click()

iiCanal = MSFlexGrid3.TextMatrix(MSFlexGrid3.RowSel, 0)

End Sub

Private Sub StatusBar2_PanelClick(ByVal Panel As ComctlLib.Panel)
  Select Case Panel
    Case "N\Atendidas"
      FrmMenu.Carrega_Flex
      Atendimento_Button1.met_ChamadasEfetuadas
      FrmnAtendidas.Show
    Case "Recados"
      FrmMenu.Carrega_Flex
      Atendimento_Button1.met_Recados
      FrmEscutaRec.Show
    Case "Diálogos"
      FrmMenu.Carrega_Flex
      Atendimento_Button1.met_Dialogos
      FrmEscutaDial.Show
    Case "Siga - me"
      'Atendimento_Button1.met_SigaMe
  End Select
End Sub
