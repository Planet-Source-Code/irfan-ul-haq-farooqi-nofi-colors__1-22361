VERSION 5.00
Begin VB.Form frmAboutRGB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About!"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   Icon            =   "frmAboutRGB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtAbout 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   3600
   End
   Begin VB.Image Image2 
      Height          =   1215
      Left            =   240
      Picture         =   "frmAboutRGB.frx":0CCA
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblSite 
      AutoSize        =   -1  'True
      Caption         =   "http://www.iamnofi.com"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1800
      TabIndex        =   3
      Top             =   3720
      Width           =   2640
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "nofi@programmer.net"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   765
      TabIndex        =   2
      Top             =   3360
      Width           =   2280
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Irfan ul Haq Farooqi"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   285
      TabIndex        =   1
      Top             =   2880
      Width           =   2760
   End
   Begin VB.Label lblAssignment 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "NOFI Colors"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   915
      TabIndex        =   0
      Top             =   120
      Width           =   2325
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1050
      Left            =   3120
      Picture         =   "frmAboutRGB.frx":2BA98
      Top             =   2520
      Width           =   900
   End
End
Attribute VB_Name = "frmAboutRGB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblSite.Left = 4000
    txtAbout.Text = "Thank you for choosing NOFI Colors, an easiest-to-use small utility that helps you to save the color of your choice. You can Copy the RGB function or QBColor function or the Hexadecimal code of any color to the Clipboard which can be used while coding your application in Microsoft Visual Basic. You can also Copy the HTML code of any color to the Clipboard which can be used while designing your Web Page. Your comments will help me to improve."
    cmdOK.Default = True
End Sub

Private Sub Timer1_Timer()
If lblSite.Left = -3000 Then
    lblSite.Left = 4000
Else
    lblSite.Left = lblSite.Left - 50
End If
cmdOK.SetFocus
End Sub
