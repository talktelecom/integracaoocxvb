VERSION 5.00
Begin VB.Form Form3 
   ClientHeight    =   930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3990
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   3990
   Begin VB.TextBox Txtgermens 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Menu mnuTiposIntervalos 
      Caption         =   "TiposIntrevalos"
      Begin VB.Menu mnuIntervalos 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub mnuIntervalos_Click(Index As Integer)

Call frmInicial.Atendimento_Button1.met_Solicita_Intervalo_Customizado(Trim(Str(mnuIntervalos(Index).Tag)), Trim(mnuIntervalos(Index).Caption))

          
End Sub
