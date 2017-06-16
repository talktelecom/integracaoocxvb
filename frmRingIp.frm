VERSION 5.00
Begin VB.Form frmRingIp 
   Caption         =   "RingIp"
   ClientHeight    =   1590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   1950
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   1440
      Top             =   120
   End
   Begin VB.CommandButton cmdAtender 
      Caption         =   "Atender"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmRingIp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const SND_ASYNC = &H1

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal _
    lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

' Play a WAV file.
'
' FileName is a string containing the full path of the file.
' If SyncExec is True, the sound is played synchronously
' Returns True if no errors occurred

Function PlayWAV(FileName As String, Optional SyncExec As Boolean) As Boolean
    If SyncExec Then
        ' play the file synchronously
        PlayWAV = PlaySound(FileName, 0, 0)
    Else
        ' play the file asynchronously
        PlayWAV = PlaySound(FileName, 0, SND_ASYNC)
    End If
End Function

Private Sub cmdAtender_Click()
Call frmInicial.Atendimento_Button1.met_AtenderIp
Unload Me
End Sub

Private Sub Timer1_Timer()
Dim Ok As Boolean
Ok = PlayWAV(App.Path + "\ring.wav", False)

DoEvents
End Sub
