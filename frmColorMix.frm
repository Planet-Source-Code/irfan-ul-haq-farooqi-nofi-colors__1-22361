VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{BF51C523-632A-11D4-837D-004005AAE138}#1.0#0"; "JKTRYICN.OCX"
Begin VB.Form frmColorMixer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NOFI Colors"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5535
   Icon            =   "frmColorMix.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   7858
      _Version        =   393216
      TabOrientation  =   2
      TabHeight       =   706
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&RGB Color"
      TabPicture(0)   =   "frmColorMix.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LabelGreen"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LabelRed"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblBlue"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblGreen"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblRed"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblMixRGB"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdCopyRGB"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "hsBlue"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "hsGreen"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "hsRed"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdRGBtoHTML"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdRGBtoHex"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdValues"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "&QB Color"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdQBtoHex"
      Tab(1).Control(1)=   "cmdQBtoHTML"
      Tab(1).Control(2)=   "cmdCopyQB"
      Tab(1).Control(3)=   "Label4"
      Tab(1).Control(4)=   "Label3"
      Tab(1).Control(5)=   "lblQB1(0)"
      Tab(1).Control(6)=   "lblQB1(15)"
      Tab(1).Control(7)=   "lblQB1(14)"
      Tab(1).Control(8)=   "lblQB1(13)"
      Tab(1).Control(9)=   "lblQB1(12)"
      Tab(1).Control(10)=   "lblQB1(11)"
      Tab(1).Control(11)=   "lblQB1(10)"
      Tab(1).Control(12)=   "lblQB1(9)"
      Tab(1).Control(13)=   "lblQB1(8)"
      Tab(1).Control(14)=   "lblQB1(7)"
      Tab(1).Control(15)=   "lblQB1(6)"
      Tab(1).Control(16)=   "lblQB1(5)"
      Tab(1).Control(17)=   "lblQB1(4)"
      Tab(1).Control(18)=   "lblQB1(3)"
      Tab(1).Control(19)=   "lblQB1(2)"
      Tab(1).Control(20)=   "lblMixQB"
      Tab(1).Control(21)=   "lblQB1(1)"
      Tab(1).ControlCount=   22
      TabCaption(2)   =   "C&ustom Color"
      TabPicture(2)   =   "frmColorMix.frx":0CE6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdCopyHex"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdCopyHTML"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "picPallet"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdPallet"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdCopyCode"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "lblCustomColor"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label6"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label5"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "lblMixCode"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      Begin VB.CommandButton cmdValues 
         Caption         =   "Enter Values"
         Height          =   375
         Left            =   960
         TabIndex        =   47
         ToolTipText     =   "Enter New Values of Red, Green and Blue"
         Top             =   3360
         Width           =   1935
      End
      Begin VB.CommandButton cmdQBtoHex 
         Caption         =   "Copy Hex Code"
         Height          =   375
         Left            =   -74040
         TabIndex        =   46
         ToolTipText     =   "Copy the Hexadecimal Code of the selected color to the Clipboard"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton cmdQBtoHTML 
         Caption         =   "Copy HTML Code"
         Height          =   375
         Left            =   -71880
         TabIndex        =   45
         ToolTipText     =   "Copy the HTML Color Code of the selected color to the Clipboard"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton cmdRGBtoHex 
         Caption         =   "Copy Hex Code"
         Height          =   375
         Left            =   960
         TabIndex        =   44
         ToolTipText     =   "Copy the Hexadecimal Code of the selected color to the Clipboard"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton cmdRGBtoHTML 
         Caption         =   "Copy HTML Code"
         Height          =   375
         Left            =   3120
         TabIndex        =   43
         ToolTipText     =   "Copy the HTML Color Code of the selected color to the Clipboard"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton cmdCopyHex 
         Caption         =   "Copy Hex Code"
         Height          =   375
         Left            =   -71880
         TabIndex        =   42
         ToolTipText     =   "Copy the Hexadecimal Code of the selected color to the Clipboard"
         Top             =   3360
         Width           =   1935
      End
      Begin VB.CommandButton cmdCopyHTML 
         Caption         =   "Copy HTML Code"
         Height          =   375
         Left            =   -71880
         TabIndex        =   41
         ToolTipText     =   "Copy the HTML Color Code of the selected color to the Clipboard"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.PictureBox picPallet 
         Height          =   2775
         Left            =   -74400
         MouseIcon       =   "frmColorMix.frx":0D02
         MousePointer    =   99  'Custom
         Picture         =   "frmColorMix.frx":100C
         ScaleHeight     =   2715
         ScaleWidth      =   2595
         TabIndex        =   37
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton cmdCopyQB 
         Caption         =   "Copy QBColor Function"
         Height          =   375
         Left            =   -71880
         TabIndex        =   32
         ToolTipText     =   "Copy the QBColor function (used in VB) of the selected color to the Clipboard"
         Top             =   3360
         Width           =   1935
      End
      Begin VB.CommandButton cmdPallet 
         Caption         =   "Choose Color form Pallet"
         Height          =   375
         Left            =   -74040
         TabIndex        =   14
         ToolTipText     =   "Open Windows Color Pallet to choose custom color"
         Top             =   3360
         Width           =   1935
      End
      Begin VB.CommandButton cmdCopyCode 
         Caption         =   "Copy Code"
         Height          =   375
         Left            =   -74040
         TabIndex        =   13
         ToolTipText     =   "Copy the selected color value to the Clipboard"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.HScrollBar hsRed 
         Height          =   255
         LargeChange     =   10
         Left            =   660
         Max             =   255
         TabIndex        =   5
         Top             =   840
         Width           =   2055
      End
      Begin VB.HScrollBar hsGreen 
         Height          =   255
         LargeChange     =   10
         Left            =   660
         Max             =   255
         TabIndex        =   4
         Top             =   1680
         Width           =   2055
      End
      Begin VB.HScrollBar hsBlue 
         Height          =   255
         LargeChange     =   10
         Left            =   660
         Max             =   255
         TabIndex        =   3
         Top             =   2580
         Width           =   2055
      End
      Begin VB.CommandButton cmdCopyRGB 
         Caption         =   "Copy RGB Function"
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         ToolTipText     =   "Copy the RGB function (used in VB) of the selected color to the Clipboard"
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label lblCustomColor 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   855
         Left            =   -71520
         TabIndex        =   40
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Custom Color"
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
         Left            =   -71460
         TabIndex        =   39
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Selected Color"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -71280
         TabIndex        =   38
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "QB Color Values"
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
         Left            =   -74400
         TabIndex        =   36
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Selected Color"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -70920
         TabIndex        =   35
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Selected Color"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3960
         TabIndex        =   34
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lblMixRGB 
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   2055
         Left            =   3840
         TabIndex        =   33
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblQB1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   0
         Left            =   -74400
         TabIndex        =   31
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblQB1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   15
         Left            =   -71880
         TabIndex        =   30
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label lblQB1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   14
         Left            =   -72720
         TabIndex        =   29
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label lblQB1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   13
         Left            =   -73560
         TabIndex        =   28
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label lblQB1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   12
         Left            =   -74400
         TabIndex        =   27
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label lblQB1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   11
         Left            =   -71880
         TabIndex        =   26
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label lblQB1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   10
         Left            =   -72720
         TabIndex        =   25
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label lblQB1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   9
         Left            =   -73560
         TabIndex        =   24
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label lblQB1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   8
         Left            =   -74400
         TabIndex        =   23
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label lblQB1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   7
         Left            =   -71880
         TabIndex        =   22
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblQB1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   6
         Left            =   -72720
         TabIndex        =   21
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblQB1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   5
         Left            =   -73560
         TabIndex        =   20
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblQB1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   4
         Left            =   -74400
         TabIndex        =   19
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblQB1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   3
         Left            =   -71880
         TabIndex        =   18
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblQB1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   2
         Left            =   -72720
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblMixQB 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   2055
         Left            =   -70920
         TabIndex        =   16
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblMixCode 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   -71520
         TabIndex        =   15
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblQB1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Index           =   1
         Left            =   -73560
         TabIndex        =   12
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblRed 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   2880
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblGreen 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2880
         TabIndex        =   10
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblBlue 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   2880
         TabIndex        =   9
         Top             =   2340
         Width           =   735
      End
      Begin VB.Label LabelRed 
         AutoSize        =   -1  'True
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   660
         TabIndex        =   8
         Top             =   480
         Width           =   450
      End
      Begin VB.Label LabelGreen 
         AutoSize        =   -1  'True
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   660
         TabIndex        =   7
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   660
         TabIndex        =   6
         Top             =   2220
         Width           =   600
      End
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4800
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComDlg.CommonDialog cmnDlg 
      Left            =   4200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      InitDir         =   "App.Path"
   End
   Begin JKTryIcn.TrayIcon TrayIcon 
      Left            =   2400
      Top             =   0
      _ExtentX        =   1508
      _ExtentY        =   661
      Picture         =   "frmColorMix.frx":18816
      ToolTipText     =   "NOFI Color"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu smnNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu smnSave 
         Caption         =   "&Save Color"
         Shortcut        =   ^S
      End
      Begin VB.Menu smnOpen 
         Caption         =   "&Open Color         "
         Shortcut        =   ^O
      End
      Begin VB.Menu smnFD1 
         Caption         =   "-"
      End
      Begin VB.Menu smnExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuCopy 
      Caption         =   "&Copy"
      Begin VB.Menu smnRGB 
         Caption         =   "RGB Function"
         Shortcut        =   ^R
      End
      Begin VB.Menu smnQB 
         Caption         =   "QBColor Function         "
         Enabled         =   0   'False
         Shortcut        =   ^Q
      End
      Begin VB.Menu smnCode 
         Caption         =   "Color Code"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu smnCD1 
         Caption         =   "-"
      End
      Begin VB.Menu smnHex 
         Caption         =   "Hex Value"
         Shortcut        =   ^H
      End
      Begin VB.Menu smnHTML 
         Caption         =   "HTML Value"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu smnAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuTrayMenu 
      Caption         =   "TrayMenu"
      Visible         =   0   'False
      Begin VB.Menu smnRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu smnTD 
         Caption         =   "-"
      End
      Begin VB.Menu smnExitTray 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmColorMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FSO As FileSystemObject
Dim txtStream As TextStream
Dim RedValue As Integer
Dim GreenValue As Integer
Dim BlueValue As Integer
Dim QBValue As Integer
Dim PalletX As Integer
Dim PalletY As Integer
Public CurTab As String

Private Sub cmdCopyCode_Click()
    Clipboard.Clear
    Clipboard.SetText lblMixCode.BackColor
End Sub

Private Sub cmdCopyHex_Click()
    Clipboard.Clear
    Clipboard.SetText HexColor(lblMixCode)
End Sub

Private Sub cmdCopyHTML_Click()
    Clipboard.Clear
    Clipboard.SetText HTMLColor(lblMixCode)
End Sub

Private Sub cmdCopyQB_Click()
    Clipboard.Clear
    Clipboard.SetText "QBColor(" & QBValue & ")" & vbCrLf
End Sub

Private Sub cmdCopyRGB_Click()
    Clipboard.Clear
    Clipboard.SetText "RGB(" & hsRed.Value & ", " & hsGreen.Value & ", " & hsBlue.Value & ")" & vbCrLf
End Sub

Private Sub cmdPallet_Click()
    On Error Resume Next
    cmnDlg.Flags = &H4
    cmnDlg.ShowColor
    lblMixCode.BackColor = cmnDlg.Color
End Sub

Private Sub cmdQBtoHex_Click()
    Clipboard.Clear
    Clipboard.SetText HexColor(lblMixQB)
End Sub

Private Sub cmdQBtoHTML_Click()
    Clipboard.Clear
    Clipboard.SetText HTMLColor(lblMixQB)
End Sub

Private Sub cmdRGBtoHex_Click()
    Clipboard.Clear
    Clipboard.SetText HexColor(lblMixRGB)
End Sub

Private Sub cmdRGBtoHTML_Click()
    Clipboard.Clear
    Clipboard.SetText HTMLColor(lblMixRGB)
End Sub

Private Sub cmdValues_Click()
    frmNewRGB.Show
End Sub

Private Sub Form_Load()
    hsRed.Value = 255
    hsGreen.Value = 255
    hsBlue.Value = 255
    For C = 0 To 15 Step 1
        lblQB1.Item(C).BackColor = QBColor(C)
        lblQB1.Item(C).Caption = C
    Next
    lblMixQB.BackColor = QBColor(15)
    QBValue = 15
    lblMixCode.BackColor = vbWhite
    CurTab = "TabRGB"
    cmnDlg.InitDir = App.Path
    
    TrayIcon.Add
End Sub

Private Sub Form_Paint()
    hsBlue.SetFocus
End Sub

Private Sub hsBlue_Change()
    lblBlue.BackColor = RGB(0, 0, hsBlue.Value)
    lblMixRGB.BackColor = RGB(hsRed.Value, hsGreen.Value, hsBlue.Value)
    lblBlue.Caption = Chr(13) & hsBlue.Value
End Sub

Private Sub hsGreen_Change()
    lblGreen.BackColor = RGB(0, hsGreen.Value, 0)
    lblMixRGB.BackColor = RGB(hsRed.Value, hsGreen.Value, hsBlue.Value)
    lblGreen.Caption = Chr(13) & hsGreen.Value
End Sub

Private Sub hsRed_Change()
    lblRed.BackColor = RGB(hsRed.Value, 0, 0)
    lblMixRGB.BackColor = RGB(hsRed.Value, hsGreen.Value, hsBlue.Value)
    lblRed.Caption = Chr(13) & hsRed.Value
End Sub

Private Sub lblQB1_Click(Index As Integer)
    lblMixQB.BackColor = QBColor(Index)
    QBValue = Index
End Sub

Private Sub picPallet_Click()
    lblMixCode.BackColor = picPallet.Point(PalletX, PalletY)
End Sub

Private Sub picPallet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCustomColor.BackColor = picPallet.Point(X, Y)
    PalletX = X
    PalletY = Y
End Sub

Private Sub smnAbout_Click()
    frmAboutRGB.Show
End Sub

Private Sub smnCopy_Click()
    Clipboard.Clear
    If CurTab = "TabRGB" Then
        Clipboard.SetText "RGB(" & hsRed.Value & ", " & hsGreen.Value & ", " & hsBlue.Value & ")"
    ElseIf CurTab = "TabCode" Then
        Clipboard.SetText lblMixCode.BackColor
    End If
End Sub

Private Sub smnCode_Click()
    Call cmdCopyCode_Click
End Sub

Private Sub smnExit_Click()
    End
End Sub

Private Sub smnExitTray_Click()
    End
End Sub

Private Sub smnHex_Click()
    If CurTab = "TabRGB" Then
        Call cmdRGBtoHex_Click
    ElseIf CurTab = "TabQB" Then
        Call cmdQBtoHex_Click
    ElseIf CurTab = "TabCustom" Then
        Call cmdCopyHex_Click
    End If
End Sub

Private Sub smnHTML_Click()
    If CurTab = "TabRGB" Then
        Call cmdRGBtoHTML_Click
    ElseIf CurTab = "TabQB" Then
        Call cmdQBtoHTML_Click
    ElseIf CurTab = "TabCustom" Then
        Call cmdCopyHTML_Click
    End If
End Sub

Private Sub smnNew_Click()
    If CurTab = "TabRGB" Then
        hsRed.Value = 255
        hsGreen.Value = 255
        hsBlue.Value = 255
    ElseIf CurTab = "TabQB" Then
        lblMixQB.BackColor = QBColor(15)
    ElseIf CurTab = "TabCustom" Then
        lblMixCode.BackColor = vbWhite
    End If
End Sub

Private Sub smnOpen_Click()
    On Error Resume Next
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If CurTab = "TabRGB" Then
        cmnDlg.DialogTitle = "Open RGB Color"
        cmnDlg.Filter = "RGB Color (*.rgc)|*.rgc"
    ElseIf CurTab = "TabQB" Then
        cmnDlg.DialogTitle = "Open QB Color"
        cmnDlg.Filter = "QB Color (*.qbc)|*.qbc"
    ElseIf CurTab = "TabCustom" Then
        cmnDlg.DialogTitle = "Open Custom Color"
        cmnDlg.Filter = "Custom Color (*.ctc)|*.ctc"
    End If
    
    cmnDlg.FileName = ""
    cmnDlg.Flags = &H4 + &H800 + &H1000
    cmnDlg.ShowOpen
    
    If cmnDlg.FileName = "" Then
    Else
        Set txtStream = FSO.OpenTextFile(cmnDlg.FileName, ForReading)
        If Right(cmnDlg.FileName, 3) = "rgc" Then
            Text1.Text = txtStream.ReadLine
            c1 = InStr(1, Text1.Text, ",")
            c2 = InStr(c1 + 1, Text1.Text, ",")
            hsRed.Value = Val(Mid(Text1.Text, 1, c1))
            hsGreen.Value = Val(Mid(Text1.Text, c1 + 1, c2 - c1))
            hsBlue.Value = Val(Mid(Text1.Text, c2 + 1))
        ElseIf Right(cmnDlg.FileName, 3) = "qbc" Then
            lblMixQB.BackColor = QBColor(Val(txtStream.ReadLine))
        ElseIf Right(cmnDlg.FileName, 3) = "ctc" Then
            lblMixCode.BackColor = Val(txtStream.ReadLine)
        Else
            msgResponse = MsgBox("Not a valid color file", vbRetryCancel + vbCritical, "Invalid File")
            If msgResponse = vbRetry Then
                Call smnOpen_Click
            ElseIf msgResponse = vbCancel Then
                'Nothing
            End If
        End If
    End If
End Sub

Private Sub smnPallet_Click()
    cmnDlg.ShowColor
    lblMixCode.BackColor = cmnDlg.Color
End Sub

Private Sub smnQB_Click()
    Call cmdCopyQB_Click
End Sub

Private Sub smnRestore_Click()
    TrayIcon.Restore
End Sub

Private Sub smnRGB_Click()
    Call cmdCopyRGB_Click
End Sub

Private Sub smnSave_Click()
    On Error Resume Next
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If CurTab = "TabRGB" Then
        cmnDlg.DialogTitle = "Save RGB Color"
        cmnDlg.Filter = "RGB Color (*.rgc)|*.rgc"
    ElseIf CurTab = "TabQB" Then
        cmnDlg.DialogTitle = "Save QB Color"
        cmnDlg.Filter = "QB Color (*.qbc)|*.qbc"
    ElseIf CurTab = "TabCustom" Then
        cmnDlg.DialogTitle = "Save Custom Color"
        cmnDlg.Filter = "Custom Color (*.ctc)|*.ctc"
    End If
    
    cmnDlg.FileName = ""
    cmnDlg.Flags = &H4 + &H2 + &H800
    cmnDlg.ShowSave
    
    If cmnDlg.FileName = "" Then
    Else
        Set txtStream = FSO.CreateTextFile(cmnDlg.FileName, True)
        If CurTab = "TabRGB" Then
            txtStream.WriteLine hsRed.Value & "," & hsGreen.Value & "," & hsBlue.Value
        ElseIf CurTab = "TabQB" Then
            txtStream.WriteLine QBValue
        ElseIf CurTab = "TabCustom" Then
            txtStream.WriteLine lblMixCode.BackColor
        End If
        txtStream.Close
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Caption = "&RGB Color" Then
        CurTab = "TabRGB"
        smnRGB.Enabled = True
        smnQB.Enabled = False
        smnCode.Enabled = False
    ElseIf SSTab1.Caption = "&QB Color" Then
        CurTab = "TabQB"
        smnRGB.Enabled = False
        smnQB.Enabled = True
        smnCode.Enabled = False
    ElseIf SSTab1.Caption = "C&ustom Color" Then
        CurTab = "TabCustom"
        smnRGB.Enabled = False
        smnQB.Enabled = False
        smnCode.Enabled = True
    End If
End Sub

Private Function HexColor(Obj As Label) As String
    HexColor = "&H" & Hex(Obj.BackColor)
End Function

Private Function HTMLColor(Obj As Label) As String
    Dim RedHTML As String
    Dim GreenHTML As String
    Dim BlueHTML As String
    
    RedHTML = Hex(Obj.BackColor And 255)
    GreenHTML = Hex(Obj.BackColor \ 256 And 255)
    BlueHTML = Hex(Obj.BackColor \ 65536 And 255)
    RedHTML = IIf(Len(RedHTML) < 2, "0" & RedHTML, RedHTML)
    GreenHTML = IIf(Len(GreenHTML) < 2, "0" & GreenHTML, GreenHTML)
    BlueHTML = IIf(Len(BlueHTML) < 2, "0" & BlueHTML, BlueHTML)
    HTMLColor = "#" & RedHTML & GreenHTML & BlueHTML
End Function

Private Sub TrayIcon_LeftDblClick()
    TrayIcon.Restore
End Sub

Private Sub TrayIcon_MouseDown(Button As Integer, Shift As Integer)
    If Button = vbRightButton Then
        PopupMenu mnuTrayMenu
    End If
End Sub
