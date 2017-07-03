VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmConf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conferência"
   ClientHeight    =   3465
   ClientLeft      =   3225
   ClientTop       =   5520
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7185
   Begin VB.CommandButton cmdEncerrar 
      Caption         =   "Encerrar"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdIniciar 
      Caption         =   "Iniciar"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtNumero 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Adicionar"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin ComctlLib.ListView lvwConf 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   3836
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Número :"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   855
   End
End
Attribute VB_Name = "frmConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Set LITEM = lvwConf.ListItems.Add(, , Trim(txtNumero.Text))
LITEM.SubItems(1) = ""
LITEM.SubItems(2) = "Parado"
LITEM.SubItems(3) = ""
LITEM.SubItems(4) = ""
txtNumero.Text = ""

End Sub

Private Sub cmdEncerrar_Click()

'If frmInicial.StatusBar1.Panels(3).Text = "Livre" Then
'  Unload Me
'  Exit Sub
'End If
'  Call frmInicial.TalkManager_Button1.met_Cancela_Conf
'Else
'  Call frmInicial.TalkManager_Button1.met_Encerrar_Conf(0)
'  Call frmInicial.TalkManager_Button1.met_Encerrar_Conf(1)
'  Call frmInicial.TalkManager_Button1.met_Cancela_Conf
'End If

If lvwConf.ListItems.Count = 0 Then
  Call frmInicial.TalkManager_Button1.met_Cancela_Conf
Else
  'Call frmInicial.TalkManager_Button1.met_Cancela_Conf
  Call frmInicial.TalkManager_Button1.met_Encerrar_Conf(0)
  Call frmInicial.TalkManager_Button1.met_Encerrar_Conf(1)
End If

Unload Me
End Sub

Private Sub cmdIniciar_Click()
Dim iExterno As Integer

iExterno = 0
For i = 1 To lvwConf.ListItems.Count
  strNum = lvwConf.ListItems(i).Text
  If Len(strNum) = 3 Then
    Call frmInicial.TalkManager_Button1.met_Adicionar_Conf(Trim(strNum), 0)
  Else
    Call frmInicial.TalkManager_Button1.met_Adicionar_Conf(Trim(strNum), 1)
    Call frmInicial.TalkManager_Button1.met_Adicionar_Numero_Conf(Trim(strNum), "x")
    iExterno = 1
  End If
Next

Call frmInicial.TalkManager_Button1.met_Iniciar_Conf(0)

End Sub

Private Sub Form_Load()
'Set lvwConf.Icons = Frmstconf.ImageList1
LWidth7 = lvwConf.Width - 100 * Screen.TwipsPerPixelX
Set LHeader = lvwConf.ColumnHeaders.Add(1, , "R\Tel\Arq", LWidth7 / 4)
Set LHeader = lvwConf.ColumnHeaders.Add(2, , "Nome", LWidth7 / 4)
Set LHeader = lvwConf.ColumnHeaders.Add(3, , "Status", LWidth7 / 2.8)
Set LHeader = lvwConf.ColumnHeaders.Add(4, , "Canal", LWidth7 / 5)
Set LHeader = lvwConf.ColumnHeaders.Add(5, , "Id", LWidth7 / 8)
lvwConf.Sorted = False
lvwConf.View = lvwReport


End Sub

