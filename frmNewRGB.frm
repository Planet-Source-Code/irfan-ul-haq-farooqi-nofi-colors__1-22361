VERSION 5.00
Begin VB.Form frmNewRGB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Values"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   Icon            =   "frmNewRGB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   990
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtBlue 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   2625
      MaxLength       =   3
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtGreen 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtRed 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   960
      MaxLength       =   3
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblBlue 
      AutoSize        =   -1  'True
      Caption         =   "Blue"
      Height          =   195
      Left            =   2775
      TabIndex        =   3
      Top             =   720
      Width           =   315
   End
   Begin VB.Label lblGreen 
      AutoSize        =   -1  'True
      Caption         =   "Green"
      Height          =   195
      Left            =   1890
      TabIndex        =   2
      Top             =   720
      Width           =   435
   End
   Begin VB.Label lblRed 
      AutoSize        =   -1  'True
      Caption         =   "Red"
      Height          =   195
      Left            =   1110
      TabIndex        =   1
      Top             =   720
      Width           =   300
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      Caption         =   "Enter new values for Red, Green and Blue"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3000
   End
End
Attribute VB_Name = "frmNewRGB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RedValue As Integer
Dim GreenValue As Integer
Dim BlueValue As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    RedValue = Val(txtRed.Text)
    GreenValue = Val(txtGreen.Text)
    BlueValue = Val(txtBlue.Text)
    
    If RedValue > 0 And RedValue < 256 Then
        frmColorMixer.hsRed.Value = RedValue
    Else
        MsgBox "Value must be between 0 and 255", vbInformation, "Invalid Vlue"
        Exit Sub
    End If
    
    If GreenValue > 0 And GreenValue < 256 Then
        frmColorMixer.hsGreen.Value = GreenValue
    Else
        MsgBox "Value must be between 0 and 255", vbInformation, "Invalid Vlue"
        Exit Sub
    End If
    
    If BlueValue > 0 And BlueValue < 256 Then
        frmColorMixer.hsBlue.Value = BlueValue
    Else
        MsgBox "Value must be between 0 and 255", vbInformation, "Invalid Vlue"
        Exit Sub
    End If
    
    Unload Me
End Sub
