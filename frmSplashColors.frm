VERSION 5.00
Begin VB.Form frmSplashColor 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   Icon            =   "frmSplashColors.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmSplashColors.frx":000C
   ScaleHeight     =   2925
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSplash 
      Interval        =   1000
      Left            =   240
      Top             =   2280
   End
End
Attribute VB_Name = "frmSplashColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DelaySplash As Integer

Private Sub Form_Click()
        Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
        Unload Me
End Sub

Private Sub Form_Load()
    DelaySplash = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmColorMixer.Show
End Sub

Private Sub tmrSplash_Timer()
    If DelaySplash = 2 Then
        Unload Me
    Else
        DelaySplash = DelaySplash + 1
    End If
End Sub
