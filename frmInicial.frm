VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{56203576-4F20-4A38-BC4F-E9CCA206104E}#1.0#0"; "Atendimento_Control.ocx"
Begin VB.Form frmInicial 
   Caption         =   "Atendimento Exemplo OCX 2.0.57b SIP"
   ClientHeight    =   6420
   ClientLeft      =   1935
   ClientTop       =   1305
   ClientWidth     =   13170
   Icon            =   "frmInicial.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   13170
   Begin Atendimento_Control.Atendimento_Button Atendimento_Button1 
      Height          =   375
      Left            =   120
      TabIndex        =   81
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
      Picture         =   "frmInicial.frx":0442
      Alignment       =   0
      Caption         =   "&Login"
   End
   Begin VB.CommandButton CmdDiscagem 
      Caption         =   "Discagem"
      Height          =   375
      Left            =   120
      Picture         =   "frmInicial.frx":045E
      TabIndex        =   79
      Top             =   4440
      Width           =   1100
   End
   Begin VB.CommandButton CmdDiscagem2 
      Caption         =   "Discagem2"
      Height          =   375
      Index           =   1
      Left            =   10920
      Picture         =   "frmInicial.frx":08A0
      TabIndex        =   80
      Top             =   4200
      Width           =   1100
   End
   Begin VB.TextBox txtOpcao 
      Height          =   285
      Left            =   12000
      TabIndex        =   78
      Text            =   "dds6a6c5"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrupoDialer 
      Caption         =   "GrupoDialer"
      Height          =   615
      Left            =   12120
      TabIndex        =   77
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   375
      Left            =   12480
      TabIndex        =   76
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "N\Disp"
      Height          =   495
      Left            =   12360
      TabIndex        =   75
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   11880
      TabIndex        =   74
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Discar Agenda"
      Height          =   495
      Left            =   12120
      TabIndex        =   73
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Discar"
      Height          =   375
      Left            =   12120
      TabIndex        =   72
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdVersao 
      Caption         =   "Versao"
      Height          =   255
      Left            =   9840
      TabIndex        =   71
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdIntGrupo 
      Caption         =   "Int. Grupo"
      Height          =   375
      Left            =   11280
      TabIndex        =   70
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdLiberaPausa 
      Caption         =   "Libera Pausa"
      Height          =   495
      Left            =   11280
      TabIndex        =   69
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtArquivo 
      Height          =   285
      Left            =   8880
      TabIndex        =   68
      Text            =   "M0121_09042012_171539.vox"
      Top             =   5400
      Width           =   2055
   End
   Begin VB.TextBox txtPathVox 
      Height          =   285
      Left            =   10080
      TabIndex        =   67
      Text            =   "\\talkstorage\epbx\ramais\r1453"
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton cmdCopiaVox 
      Caption         =   "CopiaVox"
      Height          =   375
      Left            =   11040
      TabIndex        =   66
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Frame fraKeyPad 
      Caption         =   "Envia Digitos IP"
      Height          =   1935
      Left            =   10080
      TabIndex        =   53
      Top             =   2160
      Width           =   1935
      Begin VB.CommandButton cmdKey 
         Caption         =   "1"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   65
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "2"
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   64
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "3"
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   63
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "4"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   62
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "5"
         Height          =   375
         Index           =   5
         Left            =   720
         TabIndex        =   61
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "6"
         Height          =   375
         Index           =   6
         Left            =   1200
         TabIndex        =   60
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "7"
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   59
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "8"
         Height          =   375
         Index           =   8
         Left            =   720
         TabIndex        =   58
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "9"
         Height          =   375
         Index           =   9
         Left            =   1200
         TabIndex        =   57
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   240
         TabIndex        =   56
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "0"
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   55
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton cmdKey 
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   1200
         TabIndex        =   54
         Top             =   1440
         Width           =   495
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   7800
      TabIndex        =   52
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdTesteRFC 
      Caption         =   "Teste RFC"
      Height          =   255
      Left            =   6600
      TabIndex        =   51
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdDesliga5 
      Caption         =   "Desl. 5 seg"
      Height          =   375
      Left            =   10080
      TabIndex        =   50
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdDisca5 
      Caption         =   "D. Ag 2 5 seg"
      Height          =   375
      Left            =   10080
      TabIndex        =   49
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtNum5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10080
      TabIndex        =   47
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox chk5seg 
      Caption         =   "Hab. Disc 5 Seg."
      Height          =   255
      Left            =   10080
      TabIndex        =   46
      Top             =   240
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9240
      Top             =   0
   End
   Begin VB.CommandButton cmdDAg2 
      Caption         =   "D.Ag2"
      Height          =   375
      Left            =   3720
      TabIndex        =   45
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton cmdConsultaComRota 
      Caption         =   "Cons. com Rota"
      Height          =   375
      Left            =   2520
      TabIndex        =   43
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdDiscarAgendaRamal 
      Caption         =   "Disc. Agenda Ramal"
      Height          =   375
      Left            =   4440
      TabIndex        =   42
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdTransfComInfo 
      Caption         =   "Transf + Info"
      Height          =   375
      Left            =   3720
      Picture         =   "frmInicial.frx":0CE2
      TabIndex        =   41
      Top             =   4440
      Width           =   1100
   End
   Begin VB.CommandButton cmdComentario 
      Caption         =   "Comentario"
      Height          =   375
      Left            =   2520
      TabIndex        =   40
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton CmdTransferencia_Agenda 
      Caption         =   "Transferencia_Agenda"
      Height          =   375
      Left            =   120
      Picture         =   "frmInicial.frx":19D4
      TabIndex        =   27
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton CmdDiscar_Agenda 
      Caption         =   "Discar Agenda"
      Height          =   375
      Left            =   2160
      TabIndex        =   26
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton CmdGravaOn 
      Caption         =   "&Grava/Off"
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   4920
      Width           =   1100
   End
   Begin VB.CommandButton CmdVoltar 
      Caption         =   "Voltar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8880
      Picture         =   "frmInicial.frx":1E16
      TabIndex        =   24
      Top             =   4920
      Width           =   1100
   End
   Begin VB.CommandButton CmdAgenda 
      Caption         =   "Agenda"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      Picture         =   "frmInicial.frx":2258
      TabIndex        =   23
      Top             =   4440
      Width           =   1100
   End
   Begin VB.CommandButton CmdAtualiza 
      Caption         =   "Atualiza"
      Height          =   375
      Left            =   5160
      Picture         =   "frmInicial.frx":269A
      TabIndex        =   22
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton CmdTransfere 
      Caption         =   "Transfere"
      Height          =   375
      Left            =   4200
      Picture         =   "frmInicial.frx":29A4
      TabIndex        =   21
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton CmdConsulta 
      Caption         =   "Consulta"
      Height          =   375
      Left            =   1320
      Picture         =   "frmInicial.frx":3696
      TabIndex        =   20
      Top             =   4920
      Width           =   1100
   End
   Begin VB.CommandButton CmdCaptura 
      Caption         =   "Captura"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      Picture         =   "frmInicial.frx":3AD8
      TabIndex        =   19
      Top             =   4440
      Width           =   1100
   End
   Begin VB.CommandButton CmdDesliga 
      Caption         =   "Desliga"
      Height          =   375
      Left            =   120
      Picture         =   "frmInicial.frx":3F1A
      TabIndex        =   18
      Top             =   3960
      Width           =   1100
   End
   Begin VB.CommandButton CmdEspera 
      Caption         =   "Espera"
      Height          =   375
      Left            =   120
      Picture         =   "frmInicial.frx":435C
      TabIndex        =   17
      Top             =   3480
      Width           =   1100
   End
   Begin VB.CommandButton CmdTransferencia 
      Caption         =   "Transferencia"
      Height          =   375
      Left            =   120
      Picture         =   "frmInicial.frx":47A6
      TabIndex        =   16
      Top             =   3000
      Width           =   1100
   End
   Begin VB.CommandButton CmdIntervalo 
      Caption         =   "Intervalo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      Picture         =   "frmInicial.frx":4BE8
      TabIndex        =   15
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton CmdDiscar 
      Caption         =   "Discar"
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4920
      Width           =   975
   End
   Begin VB.CheckBox chkinterno 
      Caption         =   "Interno"
      Height          =   195
      Left            =   6120
      TabIndex        =   13
      Top             =   5280
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.ComboBox Cmbdisca 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmInicial.frx":502A
      Left            =   6120
      List            =   "frmInicial.frx":502C
      TabIndex        =   12
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox txtId 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Text            =   "215"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtRamal 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Text            =   "1243"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtSenha 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      Text            =   "1234"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtIp 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      TabIndex        =   3
      Text            =   "10.0.0.2"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtPorta 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      TabIndex        =   4
      Text            =   "44900"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdIntervaloExterno 
      Caption         =   "Inter Externo"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdIntervaloGrupo 
      Caption         =   "InterGrupo"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdInterDialer 
      Caption         =   "Inter Dialer"
      Height          =   375
      Left            =   120
      Picture         =   "frmInicial.frx":502E
      TabIndex        =   9
      Top             =   2040
      Width           =   1100
   End
   Begin VB.CommandButton cmdInterCustomizado 
      Height          =   375
      Left            =   960
      Picture         =   "frmInicial.frx":5470
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdMessager 
      Caption         =   "Messager"
      Height          =   375
      Left            =   7320
      TabIndex        =   7
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdNdisp 
      Caption         =   "N/Disp"
      Height          =   375
      Left            =   120
      Picture         =   "frmInicial.frx":577A
      TabIndex        =   6
      Top             =   2520
      Width           =   1100
   End
   Begin VB.CommandButton cmdConf 
      Caption         =   "Conferência"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   4440
      Width           =   1095
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   28
      Top             =   6120
      Width           =   13170
      _ExtentX        =   23230
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   10
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "0"
            TextSave        =   "0"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Text            =   "0"
            TextSave        =   "0"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1588
            MinWidth        =   1588
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1588
            MinWidth        =   1588
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel9 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel10 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "17:39"
            Key             =   ""
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
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   29
      Top             =   5865
      Width           =   13170
      _ExtentX        =   23230
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
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1235
            MinWidth        =   1235
            Text            =   "ID"
            TextSave        =   "ID"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Último Status"
            TextSave        =   "Último Status"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "Usuário"
            TextSave        =   "Usuário"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1060
            MinWidth        =   1060
            Text            =   "Fila"
            TextSave        =   "Fila"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Text            =   "N\Atendidas"
            TextSave        =   "N\Atendidas"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Para ver suas ligações atendidas e não atendidas clique aqui"
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1588
            MinWidth        =   1588
            Text            =   "Recados"
            TextSave        =   "Recados"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Para recuperar recados clique aqui"
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1588
            MinWidth        =   1588
            Text            =   "Diálogos"
            TextSave        =   "Diálogos"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Para recuperar diálogos clique aqui"
         EndProperty
         BeginProperty Panel9 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "Siga - me"
            TextSave        =   "Siga - me"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Para transferir chamadas clique aqui"
         EndProperty
         BeginProperty Panel10 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Hora"
            TextSave        =   "Hora"
            Key             =   ""
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
      TabIndex        =   30
      Top             =   480
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      ForeColor       =   255
      BackColorBkg    =   -2147483648
      GridColor       =   16777215
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   3735
      Left            =   6240
      TabIndex        =   31
      Top             =   480
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
      TabIndex        =   32
      Top             =   2400
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Height          =   1815
      Left            =   1440
      TabIndex        =   33
      Top             =   480
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
   Begin VB.Label Label7 
      Caption         =   "Numero 5 seg.:"
      Height          =   255
      Left            =   10080
      TabIndex        =   48
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "ID Ligação:"
      Height          =   255
      Left            =   8520
      TabIndex        =   44
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label LblID_LIGACAO 
      Caption         =   "ID_LIGACAO"
      Height          =   255
      Left            =   8520
      TabIndex        =   39
      Top             =   4560
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "ID :"
      Height          =   255
      Left            =   1320
      TabIndex        =   38
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Ramal :"
      Height          =   255
      Left            =   2160
      TabIndex        =   37
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Senha :"
      Height          =   255
      Left            =   3240
      TabIndex        =   36
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Ip :"
      Height          =   255
      Left            =   4440
      TabIndex        =   35
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "Porta :"
      Height          =   255
      Left            =   5880
      TabIndex        =   34
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmInicial"
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
Dim iDialer As Integer
Dim iNumIntervalos As Integer
Dim iNaoDisponivel As Integer
Dim iConta5 As Integer
Dim iRfc As Integer

Dim iFlgNovoRecptivo As Integer

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Sub Atendimento_Button1_Click()
On Error GoTo err
err:
Select Case err
  Case 0
  
  Case err
    MsgBox "Erro : " & Trim(Str(Error(err))) + " " + Trim(Str(err.Description)), vbCritical, "Anote o erro !!! Atendimento_Button1_Click"
    Exit Sub
End Select


'  x = Atendimento_Button1.met_Ultimo_Nome_VOX
'   MsgBox x
  If Atendimento_Button1.Caption = "&Login" Then
    'Call Atendimento_Button1.met_Chamar_Tela_Logon("1")
    Call Atendimento_Button1.met_Logon_Automatico(Trim(txtRamal.Text), Trim(txtId.Text), Trim(txtSenha.Text), Trim(txtIp.Text), Trim(txtPorta.Text))
    'Call Atendimento_Button1.met_Logon_Automatico_Simple(Trim(txtRamal.Text), Trim(txtSenha.Text))
    'Call Atendimento_Button1.met_Logon_Automatico2(Trim(txtRamal.Text), Trim(txtId.Text), Trim(txtSenha.Text), Trim(txtIp.Text), Trim(txtPorta.Text))
  Else
    Call Atendimento_Button1.met_Logoff
  End If
End Sub


Private Sub Atendimento_Button1_iRetornoAtendeuAtivo(iCanal As String, iNome As String, iFone As Variant, iLigouPara As Variant, iStatus As Variant, iDuracao As Variant, iQtde As Integer)
  If iFlgTransferencia = 0 Then
'    If frmInicial.WindowState = 0 Then
'      frmInicial.Height = 3800
'    End If
    frmInicial.MSFlexGrid4.Visible = True
    frmInicial.MSFlexGrid3.Visible = True
    frmInicial.MSFlexGrid2.Visible = False
    frmInicial.MSFlexGrid1.Visible = False
    iContAtendeuPA = iQtde
    
    Debug.Print iCanal, iNome, iFone, iLigouPara, iStatus, iDuracao, iQtde
    For i = 1 To iContAtendeuPA
      frmInicial.MSFlexGrid3.Rows = iContAtendeuPA + 1
      If frmInicial.MSFlexGrid3.TextMatrix(i, 2) = "" Then
        frmInicial.MSFlexGrid3.TextMatrix(i, 0) = iCanal
        frmInicial.MSFlexGrid3.TextMatrix(i, 1) = iNome
        frmInicial.MSFlexGrid3.TextMatrix(i, 2) = iFone
        frmInicial.MSFlexGrid3.TextMatrix(i, 3) = iLigouPara
        frmInicial.MSFlexGrid3.TextMatrix(i, 4) = iStatus
        frmInicial.MSFlexGrid3.TextMatrix(i, 5) = iDuracao
        Exit For
      Else
        frmInicial.MSFlexGrid3.TextMatrix(frmInicial.MSFlexGrid3.Rows - 1, 3) = iLigouPara
        frmInicial.MSFlexGrid3.TextMatrix(frmInicial.MSFlexGrid3.Rows - 1, 4) = iStatus
        frmInicial.MSFlexGrid3.TextMatrix(frmInicial.MSFlexGrid3.Rows - 1, 5) = iDuracao
        Exit For
      End If
    Next
    'CmdDiscagem.Enabled = False
    CmdDesliga.Enabled = True
    CmdTransferencia.Enabled = True
    
  End If
  If iFlgTransferencia = 1 Then
    frmInicial.MSFlexGrid4.Visible = False
    frmInicial.MSFlexGrid3.Visible = False
    frmInicial.MSFlexGrid2.Visible = True
    frmInicial.MSFlexGrid1.Visible = True
    If MSFlexGrid1.Rows = 2 Then
      Atendimento_Button1.met_Atualizar_Ramais
    End If
  End If
End Sub

Private Sub Atendimento_Button1_iRetornoAtendeuEmTransferencia(iCanal As String, iNome As String, iFone As Variant, iLigouPara As Variant, iStatus As Variant, iDuracao As Variant, iQtde As Integer)
MsgBox iCanal + " - " + iNome + " - " + iFone + " - " + iLigouPara + " - " + iStatus + " - " + iDuracao + " - " + Str(iQtde), vbExclamation, "Atendeu em transferencia"
End Sub

Private Sub Atendimento_Button1_iRetornoAtendeuPA(iCanal As String, iNome As String, iFone As Variant, iLigouPara As Variant, iStatus As Variant, iDuracao As Variant, iQtde As Integer)
    
  If iFlgTransferencia = 0 Then
'    If frmInicial.WindowState = 0 Then
'      frmInicial.Height = 3800
'    End If
    frmInicial.MSFlexGrid4.Visible = True
    frmInicial.MSFlexGrid3.Visible = True
    frmInicial.MSFlexGrid2.Visible = False
    frmInicial.MSFlexGrid1.Visible = False
    iContAtendeuPA = iQtde
    
    Debug.Print iCanal, iNome, iFone, iLigouPara, iStatus, iDuracao, iQtde
    For i = 1 To iContAtendeuPA
      frmInicial.MSFlexGrid3.Rows = iContAtendeuPA + 1
      If frmInicial.MSFlexGrid3.TextMatrix(i, 2) = "" Then
        frmInicial.MSFlexGrid3.TextMatrix(i, 0) = iCanal
        frmInicial.MSFlexGrid3.TextMatrix(i, 1) = iNome
        frmInicial.MSFlexGrid3.TextMatrix(i, 2) = iFone
        frmInicial.MSFlexGrid3.TextMatrix(i, 3) = iLigouPara
        frmInicial.MSFlexGrid3.TextMatrix(i, 4) = iStatus
        frmInicial.MSFlexGrid3.TextMatrix(i, 5) = iDuracao
        Exit For
      Else
        frmInicial.MSFlexGrid3.TextMatrix(frmInicial.MSFlexGrid3.Rows - 1, 3) = iLigouPara
        frmInicial.MSFlexGrid3.TextMatrix(frmInicial.MSFlexGrid3.Rows - 1, 4) = iStatus
        frmInicial.MSFlexGrid3.TextMatrix(frmInicial.MSFlexGrid3.Rows - 1, 5) = iDuracao
        Exit For
      End If
    Next
    'CmdDiscagem.Enabled = False
    CmdDesliga.Enabled = True
    CmdTransferencia.Enabled = True
    
  End If
  If iFlgTransferencia = 1 Then
    frmInicial.MSFlexGrid4.Visible = False
    frmInicial.MSFlexGrid3.Visible = False
    frmInicial.MSFlexGrid2.Visible = True
    frmInicial.MSFlexGrid1.Visible = True
    If MSFlexGrid1.Rows = 2 Then
      Atendimento_Button1.met_Atualizar_Ramais
    End If
  End If
End Sub

Private Sub Atendimento_Button1_iRetornoAtendeuReceptivo(iCanal As String, iNome As String, iFone As Variant, iLigouPara As Variant, iStatus As Variant, iDuracao As Variant, iQtde As Integer)
  
  If (iFlgNovoRecptivo = "0") Then
    iFlgNovoRecptivo = 1
    'Abre a ficha
    MsgBox "Canal " & iCanal & " Nome " & iNome & " Fone " & iFone & " Ligou para " & iLigouPara & " Status " & iStatus & " Duração " & iDuracao & " QTDE " & iQtde, vbExclamation, "Atendeu Receptivo"
  End If
  
  If iFlgTransferencia = 0 Then
'    If frmInicial.WindowState = 0 Then
'      frmInicial.Height = 3800
'    End If
    frmInicial.MSFlexGrid4.Visible = True
    frmInicial.MSFlexGrid3.Visible = True
    frmInicial.MSFlexGrid2.Visible = False
    frmInicial.MSFlexGrid1.Visible = False
    iContAtendeuPA = iQtde
    
    Debug.Print iCanal, iNome, iFone, iLigouPara, iStatus, iDuracao, iQtde
    For i = 1 To iContAtendeuPA
      frmInicial.MSFlexGrid3.Rows = iContAtendeuPA + 1
      If frmInicial.MSFlexGrid3.TextMatrix(i, 2) = "" Then
        frmInicial.MSFlexGrid3.TextMatrix(i, 0) = iCanal
        frmInicial.MSFlexGrid3.TextMatrix(i, 1) = iNome
        frmInicial.MSFlexGrid3.TextMatrix(i, 2) = iFone
        frmInicial.MSFlexGrid3.TextMatrix(i, 3) = iLigouPara
        frmInicial.MSFlexGrid3.TextMatrix(i, 4) = iStatus
        frmInicial.MSFlexGrid3.TextMatrix(i, 5) = iDuracao
        Exit For
      Else
        frmInicial.MSFlexGrid3.TextMatrix(frmInicial.MSFlexGrid3.Rows - 1, 3) = iLigouPara
        frmInicial.MSFlexGrid3.TextMatrix(frmInicial.MSFlexGrid3.Rows - 1, 4) = iStatus
        frmInicial.MSFlexGrid3.TextMatrix(frmInicial.MSFlexGrid3.Rows - 1, 5) = iDuracao
        Exit For
      End If
    Next
    'CmdDiscagem.Enabled = False
    CmdDesliga.Enabled = True
    CmdTransferencia.Enabled = True
    
  End If
  If iFlgTransferencia = 1 Then
    frmInicial.MSFlexGrid4.Visible = False
    frmInicial.MSFlexGrid3.Visible = False
    frmInicial.MSFlexGrid2.Visible = True
    frmInicial.MSFlexGrid1.Visible = True
    If MSFlexGrid1.Rows = 2 Then
      Atendimento_Button1.met_Atualizar_Ramais
    End If
  End If
End Sub

Private Sub Atendimento_Button1_iRetornoAtendeuServico(iCanal As String, iNumeroServico As String)
MsgBox iCanal + " - " + iNumeroServico, vbExclamation, "Atendeu Servico"
End Sub

Private Sub Atendimento_Button1_iRetornoBotoes(iCapLogoff As Variant)
  Atendimento_Button1.Caption = iCapLogoff
End Sub

Private Sub Atendimento_Button1_iRetornoChamadasNaoAtendidas(iNome As String, iInterno As String, iFone As String, iLigouPara As String, iData As String, iHora As String)
  
  Frmnatendidas.MSFlexGrid1.Rows = Frmnatendidas.MSFlexGrid1.Rows + 1
  Frmnatendidas.MSFlexGrid1.TextMatrix(Frmnatendidas.MSFlexGrid1.Rows - 2, 0) = Trim(iNome)
  Frmnatendidas.MSFlexGrid1.TextMatrix(Frmnatendidas.MSFlexGrid1.Rows - 2, 1) = Trim(iInterno)
  Frmnatendidas.MSFlexGrid1.TextMatrix(Frmnatendidas.MSFlexGrid1.Rows - 2, 2) = Trim(iFone)
  Frmnatendidas.MSFlexGrid1.TextMatrix(Frmnatendidas.MSFlexGrid1.Rows - 2, 3) = Trim(iLigouPara)
  Frmnatendidas.MSFlexGrid1.TextMatrix(Frmnatendidas.MSFlexGrid1.Rows - 2, 4) = Trim(iData)
  Frmnatendidas.MSFlexGrid1.TextMatrix(Frmnatendidas.MSFlexGrid1.Rows - 2, 5) = iHora
  j = j + 1
  Frmnatendidas.Command2.Caption = "Não Atendidas " + Str(j)
  
End Sub

Private Sub Atendimento_Button1_iRetornoChamadasAtendidas(iNome As String, iInterno As String, iFone As String, iLigouPara As String, iData As String, iHora As String)
  Frmnatendidas.MSFlexGrid2.Rows = Frmnatendidas.MSFlexGrid2.Rows + 1
  Frmnatendidas.MSFlexGrid2.TextMatrix(Frmnatendidas.MSFlexGrid2.Rows - 2, 0) = Trim(iNome)
  Frmnatendidas.MSFlexGrid2.TextMatrix(Frmnatendidas.MSFlexGrid2.Rows - 2, 1) = Trim(iInterno)
  Frmnatendidas.MSFlexGrid2.TextMatrix(Frmnatendidas.MSFlexGrid2.Rows - 2, 2) = Trim(iFone)
  Frmnatendidas.MSFlexGrid2.TextMatrix(Frmnatendidas.MSFlexGrid2.Rows - 2, 3) = Trim(iLigouPara)
  Frmnatendidas.MSFlexGrid2.TextMatrix(Frmnatendidas.MSFlexGrid2.Rows - 2, 4) = Trim(iData)
  Frmnatendidas.MSFlexGrid2.TextMatrix(Frmnatendidas.MSFlexGrid2.Rows - 2, 5) = iHora
  i = i + 1
  Frmnatendidas.Command1.Caption = "Atendidas " + Str(i)
End Sub

Private Sub Atendimento_Button1_iRetornoChamandoEmTransferencia(iCanal As String, iNome As String, iFone As Variant, iLigouPara As Variant, iStatus As Variant, iDuracao As Variant, iQtde As Integer)
MsgBox (iCanal + " " + iNome + " " + iFone + " " + iLigouPara + " " + iStatus + " " + iDuracao + " " + Str(iQtde))
End Sub

Private Sub Atendimento_Button1_iRetornoCongestionamento(iNumeroNaoAtendeOcupado As String)
MsgBox iNumeroNaoAtendeOcupado, vbExclamation, "Congestionamento"
End Sub

Private Sub Atendimento_Button1_iRetornoConsultaCliente(iTipo As Integer, iMessagem As String)
  MsgBox (iTipo & " " & iMessagem)
  
  If iMessagem = "Ocupado" Then
    CmdConsulta.Caption = "Consulta"
  End If
End Sub

Private Sub Atendimento_Button1_iRetornoCopiaVox(iValue As String)
 MsgBox (iValue)
End Sub

Private Sub Atendimento_Button1_iRetornoDadosAgendamento(Cod_Cli_Nome As Variant, Contato_Endereco As Variant, Cidade As Variant, Estado As Variant, CEP As Variant, Numero_Servico As Variant, Numero_Cliente As Variant, Cod_Campanha As Variant)
MsgBox ("iRetornoDadosAgendamento Cod_Cli_Nome=" + Cod_Cli_Nome + ",Contato_Endereco=" + Contato_Endereco + ",Cidade=" + Cidade + ",Estado=" + Estado + ",CEP=" + CEP + ",Numero_Servico=" + Numero_Servico + ",Numero_Cliente=" + Numero_Cliente + ",Cod_Campanha=" + Cod_Campanha)
End Sub

Private Sub Atendimento_Button1_iRetornoDadosPower(Cod_Cli_Nome As Variant, Contato_Endereco As Variant, Cidade As Variant, Estado As Variant, CEP As Variant, Numero_Servico As Variant, Numero_Cliente As Variant, Cod_Campanha As Variant)
MsgBox ("iRetornoDadosPower Cod_Cli_Nome=" + Cod_Cli_Nome + ",Contato_Endereco=" + Contato_Endereco + ",Cidade=" + Cidade + ",Estado=" + Estado + ",CEP=" + CEP + ",Numero_Servico=" + Numero_Servico + ",Numero_Cliente=" + Numero_Cliente + ",Cod_Campanha=" + Cod_Campanha)
  'Atendimento_Button1.met_Agente_Nao_Disponivel
  Atendimento_Button1.met_Libera_Pausa
  Atendimento_Button1.met_IntervaloGrupo
  Call Atendimento_Button1.met_Solicita_Intervalo_Customizado("6", "x")
  Atendimento_Button1.met_IntervaloGrupo
  
  
  
End Sub

Private Sub Atendimento_Button1_iRetornoDesligado()

MsgBox "Desligado"
  If iFlgTransferencia = 0 Then
    Call Carrega_Flex
  End If
  iFlgNovoRecptivo = 0
  frmInicial.MSFlexGrid4.Visible = True
  frmInicial.MSFlexGrid3.Visible = True
  frmInicial.MSFlexGrid2.Visible = False
  frmInicial.MSFlexGrid1.Visible = False
  'frmInicial.Cmbdisca.Text = ""
  'If frmInicial.WindowState = 0 Then
  '  frmInicial.Height = 2220
  'End If
  iFlgTransferencia = 0
  'CmdTransferencia.Enabled = False
  CmdDesliga.Enabled = False
  CmdDiscagem.Enabled = True
  'CmdVoltar.Enabled = False
  iFlg = 0
  
  If iRfc = 1 Then
    'Call Atendimento_Button1.met_IntervaloGrupo
    'Call Atendimento_Button1.met_Libera_Pausa
    
    Call Atendimento_Button1.met_Solicita_Intervalo_Customizado("0", "")
    Call Atendimento_Button1.met_IntervaloGrupo
    
    iRfc = 0
  End If
  
  
  'Atendimento_Button1.met_Libera_Pausa
End Sub

Private Sub Atendimento_Button1_iRetornoDialSliderMax(iMaximoSlider As Integer)
  Frmescutadial.Slider1.Max = iMaximoSlider
End Sub

Private Sub Atendimento_Button1_iRetornoDialSliderPos(iPosicaoSlider As Integer)
  Frmescutadial.Slider1.Value = iPosicaoSlider
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
  iFlgNovoRecptivo = 0
End Sub

Private Sub Atendimento_Button1_iRetornoInFilaGrupo(iCanal As String, iNome As String, iFone As Variant, iLigouPara As Variant, iStatus As Variant, iDuracao As Variant, iQtde As Integer)
  'If frmInicial.WindowState = 0 Then frmInicial.Height = 5250
  
  frmInicial.MSFlexGrid1.Visible = False
  frmInicial.MSFlexGrid2.Visible = False
  frmInicial.MSFlexGrid3.Visible = True
  frmInicial.MSFlexGrid4.Visible = True
  iContInFilaGrupo = iQtde
  For i = 1 To iContInFilaGrupo
    frmInicial.MSFlexGrid4.Rows = iContInFilaGrupo + 1
    If frmInicial.MSFlexGrid4.TextMatrix(i, 3) = "" Then
      frmInicial.MSFlexGrid4.TextMatrix(i, 0) = iCanal
      frmInicial.MSFlexGrid4.TextMatrix(i, 1) = iNome
      frmInicial.MSFlexGrid4.TextMatrix(i, 2) = iFone
      frmInicial.MSFlexGrid4.TextMatrix(i, 3) = iLigouPara
      frmInicial.MSFlexGrid4.TextMatrix(i, 4) = iStatus
      frmInicial.MSFlexGrid4.TextMatrix(i, 5) = iDuracao
      Exit For
    End If
    If frmInicial.MSFlexGrid4.TextMatrix(i, 3) = iLigouPara Then
      frmInicial.MSFlexGrid4.TextMatrix(i, 4) = iStatus
      frmInicial.MSFlexGrid4.TextMatrix(i, 5) = iDuracao
      Exit For
    End If
  Next
End Sub


Private Sub Atendimento_Button1_iRetornoInfoAdicionalTransferencia(iDadosInfoTransferencia As String)
MsgBox iDadosInfoTransferencia, vbCritical, "iRetornoInfoAdicionalTransferencia"
End Sub

Private Sub Atendimento_Button1_iRetornoInfoDial(iRamal As String, iDataHora As String, iDuracao As String, iFone As String, iLigouPara As String, iMensagem As String, iComentario As String)
'  Frmescutadial.MSFlexGrid1.Rows = Frmescutadial.MSFlexGrid1.Rows + 1
'  Frmescutadial.MSFlexGrid1.TextMatrix(Frmescutadial.MSFlexGrid1.Rows - 2, 0) = Trim(iRamal)
'  Frmescutadial.MSFlexGrid1.TextMatrix(Frmescutadial.MSFlexGrid1.Rows - 2, 1) = Trim(iDataHora)
'  Frmescutadial.MSFlexGrid1.TextMatrix(Frmescutadial.MSFlexGrid1.Rows - 2, 2) = Trim(iDuracao)
'  Frmescutadial.MSFlexGrid1.TextMatrix(Frmescutadial.MSFlexGrid1.Rows - 2, 3) = Trim(iFone)
'  Frmescutadial.MSFlexGrid1.TextMatrix(Frmescutadial.MSFlexGrid1.Rows - 2, 4) = Trim(iLigouPara)
'  Frmescutadial.MSFlexGrid1.TextMatrix(Frmescutadial.MSFlexGrid1.Rows - 2, 5) = iMensagem
'  Frmescutadial.MSFlexGrid1.TextMatrix(Frmescutadial.MSFlexGrid1.Rows - 2, 6) = iComentario
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

Private Sub Atendimento_Button1_iRetornoInfoUra(iValue As String)
  MsgBox "Info: " + iValue, vbExclamation, "Info"
End Sub

Private Sub Atendimento_Button1_iRetornoIntervalosCustomizados(iCod As String, iNome As String)

NI = Form3!mnuIntervalos.Count
If NI = iNumIntervalos + 1 Then
  Exit Sub
End If
'Verifica se o nome ja esta na lista.
For i = 1 To NI - 1
  If Form3!mnuIntervalos(i).Caption = Trim(iNome) Then Exit For
Next i
'Se o nome não estiver na lista, acrescenta o nome
'Cria um novo "menu control".
Load Form3!mnuIntervalos(NI)
'Muda o caption do item do menu
Form3!mnuIntervalos(NI).Caption = Trim(iNome)
Form3!mnuIntervalos(NI).Tag = Trim(iCod)
Form3!mnuIntervalos(NI).Visible = True
If NI = 1 Then
  Form3.mnuIntervalos(NI).Checked = True
  
End If
If Val(Form3.mnuIntervalos(NI).Tag) = 99 Then
  If vIntervaloGrupo = 1 Then
    Form3.mnuIntervalos(NI).Checked = True
  Else
    Form3.mnuIntervalos(NI).Checked = False
  End If
End If


End Sub

Private Sub Atendimento_Button1_iRetornoLogon(iStrRamal As Variant, iStrID As Variant, iStrUltimoStatus As Variant, iStrUsuario As Variant, iStrFila As Variant, iStrNAtendidas As Variant, iStrRecados As Variant, iStrDialogos As Variant, iStrSigaMe As Variant)
  frmInicial.StatusBar1.Panels(1).Text = iStrRamal
  frmInicial.StatusBar1.Panels(2).Text = iStrID
  frmInicial.StatusBar1.Panels(3).Text = iStrUltimoStatus
  frmInicial.StatusBar1.Panels(4).Text = iStrUsuario
  frmInicial.StatusBar1.Panels(5).Text = iStrFila
  frmInicial.StatusBar1.Panels(6).Text = iStrNAtendidas
  frmInicial.StatusBar1.Panels(7).Text = iStrRecados
  frmInicial.StatusBar1.Panels(8).Text = iStrDialogos
  frmInicial.StatusBar1.Panels(9).Text = iStrSigaMe
  CmdGravaOn.Caption = "&Grava/Off"
  If iStrUltimoStatus = "Logon Válido" Then
    CmdDiscagem.Enabled = True
  End If
  CmdIntervalo.Enabled = True
End Sub

Private Sub Atendimento_Button1_iRetornoMessager(iRamal As String, iMessagem As String, iData As String, iHora As String)
  MsgBox (iRamal + " " + iMessagem + " " + iData + " " + iHora)
End Sub

Private Sub Atendimento_Button1_iRetornoNaoAtende(iNumeroNaoAtendeOcupado As String)
MsgBox iNumeroNaoAtendeOcupado, vbCritical, "Numero Nao atende"

End Sub

Private Sub Atendimento_Button1_iRetornoNomeVOX(iNomeVOX As String)
x = iNomeVOX
y = MsgBox("NomeVox=" + iNomeVOX, vbExclamation, "NomeVox")
End Sub

Private Sub Atendimento_Button1_iRetornoNumeroCliente(iTipo As Integer, iID As String, iNome As String, iNumeroA As String, iNumeroB As String)
  If iFlg = 0 Then
    'If iTipo = 1 Then Atendimento_Button1.met_Agente_Nao_Disponivel
    iFlg = 1
  End If
End Sub

Private Sub Atendimento_Button1_iRetornoOcupado(iNumeroNaoAtendeOcupado As String)
MsgBox iNumeroNaoAtendeOcupado, vbExclamation, "Congestionamento"
End Sub

Private Sub Atendimento_Button1_iRetornoOutFilaGrupo(iRamal As String)
  Exit Sub
  'If frmInicial.MSFlexGrid3.Rows = 2 Then
  '  frmInicial.Height = 2220
  'End If
  For i = 1 To MSFlexGrid4.Rows
    If frmInicial.MSFlexGrid4.TextMatrix(i, 0) = iRamal Then
      frmInicial.MSFlexGrid4.TextMatrix(i, 0) = ""
      frmInicial.MSFlexGrid4.TextMatrix(i, 1) = ""
      frmInicial.MSFlexGrid4.TextMatrix(i, 2) = ""
      frmInicial.MSFlexGrid4.TextMatrix(i, 3) = ""
      frmInicial.MSFlexGrid4.TextMatrix(i, 4) = ""
      frmInicial.MSFlexGrid4.TextMatrix(i, 5) = ""
      Exit For
    End If
  Next
End Sub

Private Sub Atendimento_Button1_iRetornoOutFilaRamal(iRamal As String)
  'If frmInicial.MSFlexGrid3.Rows = 2 Then
  '  frmInicial.Height = 2220
  'End If
End Sub

Private Sub Atendimento_Button1_iRetornoRamaisGR(iRamal As String, iNome As String, iNumeroPA As String)
  frmInicial.MSFlexGrid2.Rows = frmInicial.MSFlexGrid2.Rows + 1
  frmInicial.MSFlexGrid2.TextMatrix(frmInicial.MSFlexGrid2.Rows - 2, 0) = iRamal
  frmInicial.MSFlexGrid2.TextMatrix(frmInicial.MSFlexGrid2.Rows - 2, 1) = iNome
  frmInicial.MSFlexGrid2.TextMatrix(frmInicial.MSFlexGrid2.Rows - 2, 2) = iNumeroPA
End Sub

Private Sub Atendimento_Button1_iRetornoRamaisPA(iRamal As String, iOperador As String, iStatus As String)
'  frmInicial.MSFlexGrid1.Rows = frmInicial.MSFlexGrid1.Rows + 1
'  frmInicial.MSFlexGrid1.TextMatrix(frmInicial.MSFlexGrid1.Rows - 2, 0) = iRamal
'  frmInicial.MSFlexGrid1.TextMatrix(frmInicial.MSFlexGrid1.Rows - 2, 1) = iOperador
'  frmInicial.MSFlexGrid1.TextMatrix(frmInicial.MSFlexGrid1.Rows - 2, 2) = iStatus
'
'  CmdVoltar.Enabled = True
 ' CmdDiscagem.Enabled = False
End Sub


Private Sub Atendimento_Button1_iRetornoRamaisPA2(iRamal As String, iOperador As String, iStatus As String, iDepto As String)
  frmInicial.MSFlexGrid1.Rows = frmInicial.MSFlexGrid1.Rows + 1
  frmInicial.MSFlexGrid1.TextMatrix(frmInicial.MSFlexGrid1.Rows - 2, 0) = iRamal
  frmInicial.MSFlexGrid1.TextMatrix(frmInicial.MSFlexGrid1.Rows - 2, 1) = iOperador
  frmInicial.MSFlexGrid1.TextMatrix(frmInicial.MSFlexGrid1.Rows - 2, 2) = iStatus
  frmInicial.MSFlexGrid1.TextMatrix(frmInicial.MSFlexGrid1.Rows - 2, 3) = iDepto
  
  CmdVoltar.Enabled = True
End Sub

Private Sub Atendimento_Button1_iRetornoRamaisPabx(iRamal As String, iNome As String, iTipo As String)
 frmInicial.MSFlexGrid1.Rows = frmInicial.MSFlexGrid1.Rows + 1
  frmInicial.MSFlexGrid1.TextMatrix(frmInicial.MSFlexGrid1.Rows - 2, 0) = iRamal
  frmInicial.MSFlexGrid1.TextMatrix(frmInicial.MSFlexGrid1.Rows - 2, 1) = iNome
  frmInicial.MSFlexGrid1.TextMatrix(frmInicial.MSFlexGrid1.Rows - 2, 2) = iTipo
  
  CmdVoltar.Enabled = True
End Sub

Private Sub Atendimento_Button1_iRetornoRecSliderMax(iMaximoSlider As Integer)
  FrmEscutaRec.Slider1.Max = iMaximoSlider
End Sub

Private Sub Atendimento_Button1_iRetornoRecSliderPos(iPosicaoSlider As Integer)
  FrmEscutaRec.Slider1.Value = iPosicaoSlider
End Sub

Private Sub Atendimento_Button1_iRetornoRetornoConsulta()
  Call Carrega_Flex
  frmInicial.MSFlexGrid4.Visible = False
  frmInicial.MSFlexGrid3.Visible = False
  frmInicial.MSFlexGrid2.Visible = True
  frmInicial.MSFlexGrid1.Visible = True
  frmInicial.Cmbdisca.Text = ""
  'If frmInicial.WindowState = 0 Then
  '  frmInicial.Height = 2220
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

Private Sub Atendimento_Button1_iRetornoRingIp()
frmRingIp.Show
End Sub

Private Sub Atendimento_Button1_iRetornoSilencioDetectado(iMensagem As String)
MsgBox ("Silencio:" + iMensagem)
End Sub

Private Sub Atendimento_Button1_iRetornoSip(iMensagem As String)
MsgBox ("SIP:" + iMensagem)
End Sub

Private Sub Atendimento_Button1_iRetornoStatusConf(iRamalTelefoneArquivo As Variant, iNome As Variant, iStatus As Variant, iCanal As Variant, iID As Variant)
  'Debug.Print iRamalTelefoneArquivo & "," & iNome & "," & iStatus & "," & iCanal & "," & iID
  
  Set LITEM = frmConf.lvwConf.FindItem(iRamalTelefoneArquivo)

  If Not LITEM Is Nothing Then
    LITEM.SubItems(2) = Trim(iStatus)
  Else
  End If
End Sub

Private Sub Atendimento_Button1_iRetornoStatusGeral(iStatus As String)
  
  frmInicial.StatusBar1.Panels(3).Text = iStatus
  If iStatus = "Livre" Then
    CmdVoltar.Enabled = False
    CmdDiscagem.Enabled = True
  End If
End Sub

Private Sub Atendimento_Button1_iRetornoTelefoneAgendado(iMensagem As String)
MsgBox ("iRetornoTelefoneAgendado:" + iMensagem)
End Sub

Private Sub Atendimento_Button1_iRetornoValueAgenteNaoDisponivel(iValue As Integer)
If iValue = 2 Then
  frmInicial.StatusBar1.Panels(3).Text = "Não Disponivel"
  iNaoDisponivel = iValue
End If

If iRfc = 1 Then
  Call Atendimento_Button1.met_IntervaloGrupo
  Call Atendimento_Button1.met_Libera_Pausa
End If

End Sub

Private Sub Atendimento_Button1_iRetornoValueIntervalo(iValue As Integer)
  If iValue = 0 Then frmInicial.StatusBar1.Panels(3).Text = "Livre"
  If iValue = 1 Then frmInicial.StatusBar1.Panels(3).Text = "Intervalo"
End Sub

Private Sub Atendimento_Button2_Click()

End Sub

Private Sub Atendimento_Button1_iRetornoValueIntervaloDialer(iValue As Integer)
iDialer = iValue
End Sub

Private Sub Atendimento_Button1_iRetornoValueIntervaloGrupo(iValue As Integer)
If iValue = 1 Then
  cmdIntervaloGrupo.BackColor = vbRed  'Amarelo Claro
Else
  cmdIntervaloGrupo.BackColor = &H8000000F 'Cinza
End If
End Sub

Private Sub Atendimento_Button1_iRetornoValueIntervaloExterno(iValue As Integer)
If iValue = 1 Then
  cmdIntervaloExterno.BackColor = vbYellow   'Amarelo Claro
Else
  cmdIntervaloExterno.BackColor = &H8000000F 'Cinza
End If
End Sub

Private Sub Atendimento_Button1_iRetornoValueNumIntervalosCustomizados(iValue As Integer)
iNumIntervalos = iValue
End Sub

Private Sub Cmbdisca_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    Call CmdDiscar_Click
  End If
End Sub


Private Sub cmd_Recado_Especific_Click()
  Call Atendimento_Button1.met_Recado_Especific(txtRamal)
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

Private Sub cmdComentario_Click()
Call Atendimento_Button1.met_Inserir_Comentario("xx")
End Sub

Private Sub cmdConf_Click()
Call Atendimento_Button1.met_Conferencia(0)
frmInicial.StatusBar1.Panels(3).Text = "Conferencia"
frmConf.Show
End Sub

Private Sub CmdConsulta_Click()
  If CmdConsulta.Caption = "Consulta" Then
    Call Atendimento_Button1.met_Consulta(Cmbdisca.Text, chkinterno.Value)
    CmdConsulta.Caption = "Retorna"
  Else
    Call Atendimento_Button1.met_Retorna(Cmbdisca.Text, chkinterno.Value)
    CmdConsulta.Caption = "Consulta"
    CmdTransferencia.Enabled = True
  End If
End Sub

Private Sub cmdConsultaComRota_Click()
Call Atendimento_Button1.met_ConsultaRota(Cmbdisca.Text, "1")
End Sub

Private Sub cmdCopiaVox_Click()
Call Atendimento_Button1.met_CopiaVox(txtRamal.Text, txtPathVox.Text, txtArquivo.Text)
End Sub

Private Sub cmdDAg2_Click()
CmdDiscar.Enabled = True
  CmdVoltar.Enabled = True
  Call Atendimento_Button1.met_Discar_Agenda2(Cmbdisca.Text, "-3")
End Sub

Private Sub CmdDesliga_Click()
  Atendimento_Button1.met_Desliga
  'Atendimento_Button1.met_Libera_Pausa
  frmInicial.MSFlexGrid4.Visible = False
  frmInicial.MSFlexGrid3.Visible = False
  frmInicial.MSFlexGrid2.Visible = True
  frmInicial.MSFlexGrid1.Visible = True
  Carrega_Flex
  'frmInicial.Height = 2220
  CmdDiscagem.Enabled = True
  iFlg = 0
End Sub

Private Sub CmdDesligaExpecif_Click()
  Atendimento_Button1.met_Desliga_Especific (iiCanal)
End Sub

Private Sub cmdDesliga5_Click()
iConta5 = 0
 Atendimento_Button1.met_Desliga
  Atendimento_Button1.met_Libera_Pausa
  frmInicial.MSFlexGrid4.Visible = False
  frmInicial.MSFlexGrid3.Visible = False
  frmInicial.MSFlexGrid2.Visible = True
  frmInicial.MSFlexGrid1.Visible = True
  Carrega_Flex
  'frmInicial.Height = 2220
  CmdDiscagem.Enabled = True
  iFlg = 0
End Sub

Private Sub cmdDisca5_Click()
cmdDisca5.Enabled = True
  CmdVoltar.Enabled = True
  Call Atendimento_Button1.met_Discar_Agenda2(txtNum5.Text, "-3")
End Sub

Private Sub CmdDiscagem_Click()

Carrega_Flex

  frmInicial.MSFlexGrid4.Visible = False
  frmInicial.MSFlexGrid3.Visible = False
  frmInicial.MSFlexGrid2.Visible = True
  frmInicial.MSFlexGrid1.Visible = True
  
  


  Atendimento_Button1.met_Discagem
  CmdDiscar.Enabled = True
  'frmInicial.Height = 6465
End Sub

'Private Sub CmdDiscagem2_Click()
'
'Carrega_Flex
'
'  frmInicial.MSFlexGrid4.Visible = False
'  frmInicial.MSFlexGrid3.Visible = False
'  frmInicial.MSFlexGrid2.Visible = True
'  frmInicial.MSFlexGrid1.Visible = True
'
'
'
'
'  Atendimento_Button1.met_Discagem2
'  CmdDiscar.Enabled = True
'  'frmInicial.Height = 6465
'End Sub

Private Sub CmdDiscar_Agenda_Click()
  CmdDiscar.Enabled = True
  CmdVoltar.Enabled = True
  Call Atendimento_Button1.met_Discar_Agenda(Cmbdisca.Text)
End Sub

Private Sub CmdDiscar_Click()
  CmdDiscar.Enabled = False
  frmInicial.MSFlexGrid4.Visible = True
  frmInicial.MSFlexGrid3.Visible = True
  frmInicial.MSFlexGrid2.Visible = False
  frmInicial.MSFlexGrid1.Visible = False
  'Call Atendimento_Button1.met_Setar_CentroCusto(1, "1111", Empty)
  Call Atendimento_Button1.met_Discar(Cmbdisca.Text, chkinterno.Value)
End Sub

Private Sub cmdDiscarAgendaRamal_Click()
Atendimento_Button1.met_Discar_AgendaRamal (Cmbdisca.Text)
End Sub

Private Sub CmdEspera_Click()
  Atendimento_Button1.met_Espera
End Sub


Private Sub cmdInervaloExterno_Click()
Atendimento_Button1.met_IntervaloExterno
End Sub

Private Sub cmdGrupoDialer_Click()
If Atendimento_Button1.get_StatusIntervaloGrupoDialer = 1 Then
  Call Atendimento_Button1.met_IntervaloGrupoDialer(0)
Else
  Call Atendimento_Button1.met_IntervaloGrupoDialer(1)
End If
End Sub

Private Sub cmdInterCustomizado_Click()
'Form3!mnuIntervalos(0).Visible = False
PopupMenu Form3.mnuTiposIntervalos
End Sub

Private Sub cmdInterDialer_Click()
If iDialer = 0 Then
  Call Atendimento_Button1.met_IntervaloDialer(1)
Else
  Call Atendimento_Button1.met_IntervaloDialer(0)
End If
End Sub

Private Sub CmdIntervalo_Click()
  Atendimento_Button1.met_Intervalo
End Sub

Private Sub cmdIntGrupo_Click()
Call Atendimento_Button1.met_IntervaloGrupo
End Sub

Private Sub CmdLiberaPausa_Click()
  Call Atendimento_Button1.met_Libera_Pausa
End Sub

Private Sub cmdKey_Click(Index As Integer)

Select Case Index

    Case 0
        Call Atendimento_Button1.met_EnviaDigitoIp("0")
    Case 1
        Call Atendimento_Button1.met_EnviaDigitoIp("1")
    Case 2
        Call Atendimento_Button1.met_EnviaDigitoIp("2")
    Case 3
        Call Atendimento_Button1.met_EnviaDigitoIp("3")
    Case 4
        Call Atendimento_Button1.met_EnviaDigitoIp("4")
    Case 5
        Call Atendimento_Button1.met_EnviaDigitoIp("5")
    Case 6
        Call Atendimento_Button1.met_EnviaDigitoIp("6")
    Case 7
        Call Atendimento_Button1.met_EnviaDigitoIp("7")
    Case 8
        Call Atendimento_Button1.met_EnviaDigitoIp("8")
    Case 9
        Call Atendimento_Button1.met_EnviaDigitoIp("9")
    Case 10
        Call Atendimento_Button1.met_EnviaDigitoIp("*")
    Case 11
        Call Atendimento_Button1.met_EnviaDigitoIp("#")
        
End Select

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
  iRetorno = Atendimento_Button1.met_Cadastro_Operadora_Param("2", "011", "0", "0", "0", True, False)
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

Private Sub cmdIntervaloExterno_Click()
Atendimento_Button1.met_IntervaloExterno
End Sub

Private Sub cmdIntervaloGrupo_Click()
Atendimento_Button1.met_IntervaloGrupo
Atendimento_Button1.get_StatusIntervaloGrupo

End Sub

Private Sub cmdNdisp_Click()

If iNaoDisponivel <> 0 Then
  Call Atendimento_Button1.met_Libera_Pausa
  iNaoDisponivel = 0
Else
  Call Atendimento_Button1.met_Agente_Nao_Disponivel
End If


End Sub

Private Sub cmdTesteRFC_Click()
'Call Atendimento_Button1.met_IntervaloGrupo
Call Atendimento_Button1.met_Desliga
iRfc = 1

End Sub

Private Sub cmdTransfComInfo_Click()

Call Atendimento_Button1.met_Transfere_Info_Adicional(Cmbdisca.Text, "Teste met_Transfere_Info_Adicional")

End Sub

Private Sub CmdTransfere_Click()
  iFlgTransferencia = 0
  Call Atendimento_Button1.met_Transfere_Ligacao(Cmbdisca.Text, chkinterno.Value)
  CmdTransfere.Enabled = False
  CmdDesliga.Enabled = False
  CmdVoltar.Enabled = False
  CmdVoltar.Enabled = False
  'frmInicial.Height = 2220
End Sub

Private Sub CmdTransferencia_Agenda_Click()
  Atendimento_Button1.met_Transferencia_Agenda
  iFlgTransferencia = 1
  CmdTransferencia.Enabled = False
  CmdTransfere.Enabled = True
  CmdDiscar.Enabled = True
  CmdVoltar.Enabled = True
  
  'frmInicial.Height = 6465

End Sub

Private Sub CmdTransferencia_Click()
    Atendimento_Button1.met_Transferencia
  iFlgTransferencia = 1
  CmdTransferencia.Enabled = False
  CmdTransfere.Enabled = True
  'frmInicial.Height = 6465
End Sub

Private Sub cmdVersao_Click()
MsgBox (Atendimento_Button1.met_Versao_Atendimento_Control)
End Sub

Private Sub CmdVoltar_Click()
  'Call Carrega_Flex
  Atendimento_Button1.met_Voltar_Ligacao
  'frmInicial.Height = 2220
  CmdDiscar.Enabled = False
  CmdTransferencia.Enabled = True
  CmdConsulta.Caption = "Consulta"
  
  
  
  frmInicial.MSFlexGrid4.Visible = True
  frmInicial.MSFlexGrid3.Visible = True
  frmInicial.MSFlexGrid2.Visible = False
  frmInicial.MSFlexGrid1.Visible = False
  
  
End Sub


Private Sub Command2_Click()
'Call Atendimento_Button1.met_Transferencia_Agenda
'Call Atendimento_Button1.met_Consulta("91953907", 0)
'Call Atendimento_Button1.met_Libera_Pausa
Call Atendimento_Button1.met_IntervaloGrupo
Call Atendimento_Button1.met_Libera_Pausa
Atendimento_Button1.met_Discagem
Call Atendimento_Button1.met_Discar(Cmbdisca.Text, 0)
Call Atendimento_Button1.met_Agente_Nao_Disponivel

End Sub

Private Sub Command3_Click()
Call Atendimento_Button1.met_Libera_Pausa
Call Atendimento_Button1.met_Discar_Agenda(Cmbdisca.Text)
Call Atendimento_Button1.met_Agente_Nao_Disponivel

End Sub

Private Sub Command4_Click()

Call Atendimento_Button1.met_AgendaTelefone("9998893", "3", "11", "993083463", "2013-08-29 10:54:00", "1227", "Celso", "", "")
'Atendimento_Button1.met_Voltar_Ligacao

'MsgBox "Status Intervalo Grupo=" + Atendimento_Button1.get_StatusIntervaloGrupo
'  Call Atendimento_Button1.met_Conferencia(0)
'Call Atendimento_Button1.met_Solicita_Intervalo_Customizado("103", "x")
'"dds6a6c5"
'Call Atendimento_Button1.met_ConverteVoxWav("c:\talk\teste.vox", "c:\talk\teste.wav", "6", "4")

End Sub

Private Sub Command1_Click()
's = Atendimento_Button1.met_Versao_Atendimento_Control()
Call Atendimento_Button1.met_Dialogos
Call Atendimento_Button1.met_PlayDial(Trim(txtArquivo.Text), 0)

 ' Call Atendimento_Button1.met_Informacao_Estatistica("11111", "2", "Teste")
End Sub

Private Sub Command5_Click()
'Atendimento_Button1.met_Desliga
Call Atendimento_Button1.met_Agente_Nao_Disponivel
'  Call Atendimento_Button1.met_Adicionar_Conf("236", 0)
'  Call Atendimento_Button1.met_Adicionar_Conf("216", 0)
'  Call Atendimento_Button1.met_Iniciar_Conf(0)
End Sub

Private Sub Command6_Click()
MsgBox Atendimento_Button1.get_StatusIntervaloGrupoDialer
'Call Atendimento_Button1.met_Cadastro_Operadora_Param("5", "11", "4", "5", "6", True, False)
 ' Call Atendimento_Button1.met_Encerrar_Conf(0)
End Sub

Private Sub Form_Load()
 ' Caption = "Atendimento OCX 1.7.00b " & Format(Now, "DD.MM.YYYY HH:MM:SS")
  Carrega_Flex
  CmdGravaOn.BackColor = &H8000000F
  'Atendimento_Button1.Caption = "&Login"
  iDialer = 0
  iNaoDisponivel = 0
  iFlgNovoRecptivo = 0
  iConta5 = 0
  iRfc = 0
  'frmInicial.Height = 2220
  'Call Atendimento_Button1.met_ShowAboutBox
End Sub
Sub Carrega_Flex()
  frmInicial.MSFlexGrid4.Visible = False
  frmInicial.MSFlexGrid3.Visible = False
  frmInicial.MSFlexGrid2.Visible = True
  frmInicial.MSFlexGrid1.Visible = True
  
  frmInicial.MSFlexGrid1.Clear
  frmInicial.MSFlexGrid1.TextMatrix(0, 0) = "Ramal"
  frmInicial.MSFlexGrid1.TextMatrix(0, 1) = "Operador"
  frmInicial.MSFlexGrid1.TextMatrix(0, 2) = "Status"
  frmInicial.MSFlexGrid1.TextMatrix(0, 3) = "Depto"
  frmInicial.MSFlexGrid1.ColWidth(0) = 700
  frmInicial.MSFlexGrid1.ColWidth(1) = 1800
  frmInicial.MSFlexGrid1.ColWidth(2) = 1800
  frmInicial.MSFlexGrid1.Rows = 2
  
  frmInicial.MSFlexGrid2.Clear
  frmInicial.MSFlexGrid2.TextMatrix(0, 0) = "Ramal"
  frmInicial.MSFlexGrid2.TextMatrix(0, 1) = "Nome"
  frmInicial.MSFlexGrid2.TextMatrix(0, 2) = "Nr Pa"
  frmInicial.MSFlexGrid2.ColWidth(0) = 700
  frmInicial.MSFlexGrid2.ColWidth(1) = 1800
  frmInicial.MSFlexGrid2.ColWidth(2) = 800
  frmInicial.MSFlexGrid2.Rows = 2

  frmInicial.MSFlexGrid3.Clear
  frmInicial.MSFlexGrid3.TextMatrix(0, 0) = "Canal"
  frmInicial.MSFlexGrid3.TextMatrix(0, 1) = "Nome"
  frmInicial.MSFlexGrid3.TextMatrix(0, 2) = "Fone"
  frmInicial.MSFlexGrid3.TextMatrix(0, 3) = "LigouPara"
  frmInicial.MSFlexGrid3.TextMatrix(0, 4) = "Status"
  frmInicial.MSFlexGrid3.TextMatrix(0, 5) = "Duracao"
  frmInicial.MSFlexGrid3.ColWidth(0) = 600
  frmInicial.MSFlexGrid3.ColWidth(1) = 1800
  frmInicial.MSFlexGrid3.ColWidth(2) = 1800
  frmInicial.MSFlexGrid3.ColWidth(3) = 1800
  frmInicial.MSFlexGrid3.ColWidth(4) = 1500
  frmInicial.MSFlexGrid3.ColWidth(5) = 1000
  frmInicial.MSFlexGrid3.Rows = 2

  frmInicial.MSFlexGrid4.Clear
  frmInicial.MSFlexGrid4.TextMatrix(0, 0) = "Canal"
  frmInicial.MSFlexGrid4.TextMatrix(0, 1) = "Nome"
  frmInicial.MSFlexGrid4.TextMatrix(0, 2) = "Fone"
  frmInicial.MSFlexGrid4.TextMatrix(0, 3) = "LigouPara"
  frmInicial.MSFlexGrid4.TextMatrix(0, 4) = "Status"
  frmInicial.MSFlexGrid4.TextMatrix(0, 5) = "Duracao"
  frmInicial.MSFlexGrid4.ColWidth(0) = 600
  frmInicial.MSFlexGrid4.ColWidth(1) = 1800
  frmInicial.MSFlexGrid4.ColWidth(2) = 1800
  frmInicial.MSFlexGrid4.ColWidth(3) = 1800
  frmInicial.MSFlexGrid4.ColWidth(4) = 1500
  frmInicial.MSFlexGrid4.ColWidth(5) = 1000
  frmInicial.MSFlexGrid4.Rows = 2
   
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub mnusair_Click()
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
      frmInicial.Carrega_Flex
      Atendimento_Button1.met_ChamadasEfetuadas
      'Frmnatendidas.Show
    Case "Recados"
      frmInicial.Carrega_Flex
      Atendimento_Button1.met_Recados
      'FrmEscutaRec.Show
    Case "Diálogos"
      frmInicial.Carrega_Flex
      Atendimento_Button1.met_Dialogos
      'Frmescutadial.Show
    Case "Siga - me"
      'Atendimento_Button1.met_SigaMe
  End Select
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Timer1_Timer()
If chk5seg.Value = 1 Then
  iConta5 = iConta5 + 1
  If iConta5 = 5 Then
    cmdDisca5.Enabled = True
    CmdVoltar.Enabled = True
    Call Atendimento_Button1.met_Discar_Agenda2(txtNum5.Text, "-3")
  End If
End If
End Sub

