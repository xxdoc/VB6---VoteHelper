VERSION 5.00
Begin VB.Form FormMainWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VoteHelper¡¡v1.00¡¡by Sam Toki"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   795
   ClientWidth     =   14985
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "ËÎÌå"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   Icon            =   "FormMainWindow.frx":0000
   LinkTopic       =   "FormMainWindow"
   MaxButton       =   0   'False
   MouseIcon       =   "FormMainWindow.frx":25CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   9225
   ScaleWidth      =   14985
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer TimerMaxQuanBlink 
      Interval        =   750
      Left            =   96
      Top             =   1056
   End
   Begin VB.CommandButton CmdTotalQuan 
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   876
      Left            =   13056
      TabIndex        =   2
      Top             =   192
      Width           =   1740
   End
   Begin VB.TextBox TextItemTitle6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAFAFA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   24.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   684
      Left            =   1056
      MousePointer    =   3  'I-Beam
      TabIndex        =   14
      Text            =   "Candidate Name"
      Top             =   6720
      Width           =   4150
   End
   Begin VB.TextBox TextCommand 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00AA7700&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   590
      Left            =   14145
      TabIndex        =   30
      Top             =   8544
      Width           =   660
   End
   Begin VB.TextBox TextVoteInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAFAFA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   570
      Left            =   190
      MousePointer    =   3  'I-Beam
      TabIndex        =   27
      Text            =   "Enter More Information Here"
      Top             =   7770
      Width           =   14600
   End
   Begin VB.TextBox TextItemTitle5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAFAFA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   24.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   684
      Left            =   1056
      MousePointer    =   3  'I-Beam
      TabIndex        =   13
      Text            =   "Candidate Name"
      Top             =   5664
      Width           =   4150
   End
   Begin VB.TextBox TextItemTitle4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAFAFA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   24.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   684
      Left            =   1056
      MousePointer    =   3  'I-Beam
      TabIndex        =   12
      Text            =   "Candidate Name"
      Top             =   4608
      Width           =   4150
   End
   Begin VB.TextBox TextItemTitle3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAFAFA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   24.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   684
      Left            =   1056
      MousePointer    =   3  'I-Beam
      TabIndex        =   11
      Text            =   "Candidate Name"
      Top             =   3552
      Width           =   4150
   End
   Begin VB.TextBox TextItemTitle2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAFAFA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   24.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   684
      Left            =   1056
      MousePointer    =   3  'I-Beam
      TabIndex        =   10
      Text            =   "Candidate Name"
      Top             =   2496
      Width           =   4150
   End
   Begin VB.TextBox TextItemTitle1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAFAFA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   24.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   684
      Left            =   1056
      MousePointer    =   3  'I-Beam
      TabIndex        =   9
      Text            =   "Candidate Name"
      Top             =   1440
      Width           =   4150
   End
   Begin VB.TextBox TextVoteTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00FAFAFA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   876
      Left            =   190
      MousePointer    =   3  'I-Beam
      TabIndex        =   0
      Text            =   "Enter Vote Topic Here"
      Top             =   190
      Width           =   11052
   End
   Begin VB.Label LabelTotalQuanTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   11430
      TabIndex        =   1
      Top             =   370
      Width           =   1455
   End
   Begin VB.Label LabelItemPerc6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   680
      Left            =   12864
      TabIndex        =   26
      Top             =   6720
      Width           =   1836
   End
   Begin VB.Label LabelItemPerc5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   680
      Left            =   12864
      TabIndex        =   25
      Top             =   5664
      Width           =   1836
   End
   Begin VB.Label LabelItemPerc4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   680
      Left            =   12864
      TabIndex        =   24
      Top             =   4608
      Width           =   1836
   End
   Begin VB.Label LabelItemPerc3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   680
      Left            =   12864
      TabIndex        =   23
      Top             =   3552
      Width           =   1836
   End
   Begin VB.Label LabelItemPerc2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   680
      Left            =   12864
      TabIndex        =   22
      Top             =   2496
      Width           =   1836
   End
   Begin VB.Label LabelItemPerc1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   680
      Left            =   12864
      TabIndex        =   21
      Top             =   1440
      Width           =   1836
   End
   Begin VB.Label LabelItemQuan6 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   680
      Left            =   5472
      TabIndex        =   20
      Top             =   6720
      Width           =   1830
   End
   Begin VB.Label LabelItemQuan5 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   680
      Left            =   5472
      TabIndex        =   19
      Top             =   5664
      Width           =   1830
   End
   Begin VB.Label LabelItemQuan4 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   680
      Left            =   5472
      TabIndex        =   18
      Top             =   4608
      Width           =   1830
   End
   Begin VB.Label LabelItemQuan3 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   680
      Left            =   5472
      TabIndex        =   17
      Top             =   3552
      Width           =   1830
   End
   Begin VB.Label LabelItemQuan2 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   680
      Left            =   5472
      TabIndex        =   16
      Top             =   2496
      Width           =   1830
   End
   Begin VB.Label LabelItemQuan1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   680
      Left            =   5472
      TabIndex        =   15
      Top             =   1440
      Width           =   1830
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      Height          =   684
      Left            =   5376
      Top             =   6720
      Width           =   9420
   End
   Begin VB.Label LabelItemNum6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   972
      Left            =   96
      TabIndex        =   8
      Top             =   6528
      Width           =   876
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      Height          =   684
      Left            =   5376
      Top             =   5664
      Width           =   9420
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      Height          =   684
      Left            =   5376
      Top             =   4608
      Width           =   9420
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      Height          =   684
      Left            =   5376
      Top             =   3552
      Width           =   9420
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      Height          =   684
      Left            =   5376
      Top             =   2496
      Width           =   9420
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H00C0C0C0&
      Height          =   684
      Left            =   5376
      Top             =   1440
      Width           =   9420
   End
   Begin VB.Label LabelInputCommand 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Press key:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   12285
      TabIndex        =   29
      Top             =   8610
      Width           =   1695
   End
   Begin VB.Label LabelStatusBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      Caption         =   "Welcome! Press F5 to start voting, F6 to change quantity."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   190
      TabIndex        =   28
      Top             =   8610
      Width           =   11895
   End
   Begin VB.Label LabelItemNum5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   972
      Left            =   96
      TabIndex        =   7
      Top             =   5472
      Width           =   876
   End
   Begin VB.Label LabelItemNum4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   972
      Left            =   96
      TabIndex        =   6
      Top             =   4416
      Width           =   876
   End
   Begin VB.Label LabelItemNum3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   972
      Left            =   96
      TabIndex        =   5
      Top             =   3360
      Width           =   876
   End
   Begin VB.Label LabelItemNum2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   972
      Left            =   96
      TabIndex        =   4
      Top             =   2304
      Width           =   876
   End
   Begin VB.Label LabelItemNum1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS PGothic"
         Size            =   50.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   972
      Left            =   96
      TabIndex        =   3
      Top             =   1248
      Width           =   876
   End
   Begin VB.Shape ShapeItemBar1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   684
      Left            =   5376
      Top             =   1440
      Width           =   120
   End
   Begin VB.Shape ShapeItemBar2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   684
      Left            =   5376
      Top             =   2496
      Width           =   120
   End
   Begin VB.Shape ShapeItemBar3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   684
      Left            =   5376
      Top             =   3552
      Width           =   120
   End
   Begin VB.Shape ShapeItemBar4 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFC0&
      FillStyle       =   0  'Solid
      Height          =   684
      Left            =   5376
      Top             =   4608
      Width           =   120
   End
   Begin VB.Shape ShapeItemBar5 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   684
      Left            =   5376
      Top             =   5664
      Width           =   120
   End
   Begin VB.Shape ShapeItemBar6 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   684
      Left            =   5376
      Top             =   6720
      Width           =   120
   End
   Begin VB.Menu MenuCtrl 
      Caption         =   "Controls"
      Begin VB.Menu MenuCtrlTotalQuan 
         Caption         =   "¡ù¡¡Quantity: 50"
         Shortcut        =   {F6}
      End
      Begin VB.Menu MenuCtrl1_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuCtrlStart 
         Caption         =   "¡ð¡¡Start"
         Shortcut        =   {F5}
      End
      Begin VB.Menu MenuCtrlClear 
         Caption         =   "£ª¡¡Clear Statistics"
         Shortcut        =   {F7}
      End
      Begin VB.Menu MenuCtrl2_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuCtrlVote1 
         Caption         =   "¢Ù¡¡Vote 1"
         Enabled         =   0   'False
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu MenuCtrlVote2 
         Caption         =   "¢Ú¡¡Vote 2"
         Enabled         =   0   'False
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu MenuCtrlVote3 
         Caption         =   "¢Û¡¡Vote 3"
         Enabled         =   0   'False
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu MenuCtrlVote4 
         Caption         =   "¢Ü¡¡Vote 4"
         Enabled         =   0   'False
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu MenuCtrlVote5 
         Caption         =   "¢Ý¡¡Vote 5"
         Enabled         =   0   'False
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu MenuCtrlVote6 
         Caption         =   "¢Þ¡¡Vote 6"
         Enabled         =   0   'False
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu MenuCtrl3_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuCtrl4_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuCtrlAbout 
         Caption         =   "About..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu MenuCtrl5_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuCtrlEXIT 
         Caption         =   "EXIT"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu Menu1_ 
      Caption         =   "¡¡|¡¡"
      Enabled         =   0   'False
   End
   Begin VB.Menu MenuLanguage 
      Caption         =   "£Á×Ö¤¢"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "FormMainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[] DIM []

Option Explicit

'DIM Menus...
Public setlanguage As String
Public inputnumberdigits As Integer
Public windowanimationswitch As Boolean

'DIM Controls...
Public totalquan As Single
Dim currentquan As Single
Dim itemquan1 As Single
Dim itemquan2 As Single
Dim itemquan3 As Single
Dim itemquan4 As Single
Dim itemquan5 As Single
Dim itemquan6 As Single
Dim itemperc1 As Single
Dim itemperc2 As Single
Dim itemperc3 As Single
Dim itemperc4 As Single
Dim itemperc5 As Single
Dim itemperc6 As Single

Dim status As Single

Dim maxquan As Single
Dim maxquanJudgeLoop As Single 'MAX QUANTITY JUDGE, CODES FROM INTERNET
Dim Arr As Variant 'MAX QUANTITY JUDGE, CODES FROM INTERNET
Dim blinkorder As Single

'DIM Dialogue...
Public answer

'================================================================================

'================================================================================

'[] LOAD []

    Public Sub Form_Load()
        'FIRST GENERAL RESET
        setlanguage = "ENG"
        inputnumberdigits = 4
        windowanimationswitch = True

        totalquan = 50
        currentquan = 1
        itemquan1 = 0
        itemquan2 = 0
        itemquan3 = 0
        itemquan4 = 0
        itemquan5 = 0
        itemquan6 = 0
        itemperc1 = 0
        itemperc2 = 0
        itemperc3 = 0
        itemperc4 = 0
        itemperc5 = 0
        itemperc6 = 0

        LabelItemNum1.BackStyle = 0
        LabelItemNum2.BackStyle = 0
        LabelItemNum3.BackStyle = 0
        LabelItemNum4.BackStyle = 0
        LabelItemNum5.BackStyle = 0
        LabelItemNum6.BackStyle = 0

        status = 0

        maxquan = 0
        blinkorder = 1

        MenuCtrlStart.Enabled = True
        MenuCtrlTotalQuan.Enabled = True
        MenuCtrlVote1.Enabled = False
        MenuCtrlVote2.Enabled = False
        MenuCtrlVote3.Enabled = False
        MenuCtrlVote4.Enabled = False
        MenuCtrlVote5.Enabled = False
        MenuCtrlVote6.Enabled = False
        MenuCtrlClear.Enabled = True
        CmdTotalQuan.Enabled = True
        TextCommand.Enabled = False
        TextCommand.BackColor = &HAA7700
        TimerMaxQuanBlink.Enabled = True

        Call Refresher
        Call TimerMaxQuanBlink_Timer

        MenuCtrlStart.Caption = "¡ð¡¡Start"
        LabelStatusBar.Caption = "Welcome! Press F5 to start voting, F6 to change quantity."
    End Sub

'[] TIMERS []

    'CODES FROM INTERNET
    Public Function MaxQuanJudge(Arr As Variant)
        MaxQuanJudge = Arr(0)
        For maxquanJudgeLoop = 0 To UBound(Arr)
        If Arr(maxquanJudgeLoop) > MaxQuanJudge Then MaxQuanJudge = Arr(maxquanJudgeLoop)
        Next
    End Function

    Public Sub Refresher()
        'REFRESH TOTALQUAN
        MenuCtrlTotalQuan.Caption = "¡ù¡¡Quantity: " & totalquan
        CmdTotalQuan.Caption = totalquan

        LabelItemQuan1.Caption = itemquan1
        LabelItemQuan2.Caption = itemquan2
        LabelItemQuan3.Caption = itemquan3
        LabelItemQuan4.Caption = itemquan4
        LabelItemQuan5.Caption = itemquan5
        LabelItemQuan6.Caption = itemquan6

        Arr = Array(itemquan1, itemquan2, itemquan3, itemquan4, itemquan5, itemquan6)
        maxquan = MaxQuanJudge(Arr)

        If Not ((itemquan1 + itemquan2 + itemquan3 + itemquan4 + itemquan5 + itemquan6) = 0) Then
            itemperc1 = Int(100 * itemquan1 / (itemquan1 + itemquan2 + itemquan3 + itemquan4 + itemquan5 + itemquan6))
            itemperc2 = Int(100 * itemquan2 / (itemquan1 + itemquan2 + itemquan3 + itemquan4 + itemquan5 + itemquan6))
            itemperc3 = Int(100 * itemquan3 / (itemquan1 + itemquan2 + itemquan3 + itemquan4 + itemquan5 + itemquan6))
            itemperc4 = Int(100 * itemquan4 / (itemquan1 + itemquan2 + itemquan3 + itemquan4 + itemquan5 + itemquan6))
            itemperc5 = Int(100 * itemquan5 / (itemquan1 + itemquan2 + itemquan3 + itemquan4 + itemquan5 + itemquan6))
            itemperc6 = Int(100 * itemquan6 / (itemquan1 + itemquan2 + itemquan3 + itemquan4 + itemquan5 + itemquan6))
        End If
            LabelItemPerc1.Caption = itemperc1 & "%"
            LabelItemPerc2.Caption = itemperc2 & "%"
            LabelItemPerc3.Caption = itemperc3 & "%"
            LabelItemPerc4.Caption = itemperc4 & "%"
            LabelItemPerc5.Caption = itemperc5 & "%"
            LabelItemPerc6.Caption = itemperc6 & "%"
        If Not ((itemquan1 + itemquan2 + itemquan3 + itemquan4 + itemquan5 + itemquan6) = 0) Then
            itemperc1 = Int(100 * itemquan1 / maxquan)
            itemperc2 = Int(100 * itemquan2 / maxquan)
            itemperc3 = Int(100 * itemquan3 / maxquan)
            itemperc4 = Int(100 * itemquan4 / maxquan)
            itemperc5 = Int(100 * itemquan5 / maxquan)
            itemperc6 = Int(100 * itemquan6 / maxquan)
        End If
            ShapeItemBar1.Width = 120 + 93 * itemperc1
            ShapeItemBar2.Width = 120 + 93 * itemperc2
            ShapeItemBar3.Width = 120 + 93 * itemperc3
            ShapeItemBar4.Width = 120 + 93 * itemperc4
            ShapeItemBar5.Width = 120 + 93 * itemperc5
            ShapeItemBar6.Width = 120 + 93 * itemperc6

        'CHECK IF VOTE ENDS
        If currentquan > totalquan Then
            currentquan = totalquan
            status = 0

            MenuCtrlTotalQuan.Enabled = False
            MenuCtrlStart.Caption = "¡ð¡¡Start"
            MenuCtrlStart.Enabled = False
            MenuCtrlVote1.Enabled = False
            MenuCtrlVote2.Enabled = False
            MenuCtrlVote3.Enabled = False
            MenuCtrlVote4.Enabled = False
            MenuCtrlVote5.Enabled = False
            MenuCtrlVote6.Enabled = False
            MenuCtrlClear.Enabled = True
            CmdTotalQuan.Enabled = False
            TextCommand.Enabled = False
            TextCommand.BackColor = &HAA7700

            LabelStatusBar.Caption = "Vote finished! Press F7 to clear statistics so as to start a new vote."
        End If
    End Sub

    Public Sub TimerMaxQuanBlink_Timer()
        If maxquan = 0 Then Exit Sub

        If itemquan1 = maxquan Then
            If blinkorder = 1 Then LabelItemNum1.BackStyle = 1 Else LabelItemNum1.BackStyle = 0
            Else: LabelItemNum1.BackStyle = 0
        End If
        If itemquan2 = maxquan Then
            If blinkorder = 1 Then LabelItemNum2.BackStyle = 1 Else LabelItemNum2.BackStyle = 0
            Else: LabelItemNum2.BackStyle = 0
        End If
        If itemquan3 = maxquan Then
            If blinkorder = 1 Then LabelItemNum3.BackStyle = 1 Else LabelItemNum3.BackStyle = 0
            Else: LabelItemNum3.BackStyle = 0
        End If
        If itemquan4 = maxquan Then
            If blinkorder = 1 Then LabelItemNum4.BackStyle = 1 Else LabelItemNum4.BackStyle = 0
            Else: LabelItemNum4.BackStyle = 0
        End If
        If itemquan5 = maxquan Then
            If blinkorder = 1 Then LabelItemNum5.BackStyle = 1 Else LabelItemNum5.BackStyle = 0
            Else: LabelItemNum5.BackStyle = 0
        End If
        If itemquan6 = maxquan Then
            If blinkorder = 1 Then LabelItemNum6.BackStyle = 1 Else LabelItemNum6.BackStyle = 0
            Else: LabelItemNum6.BackStyle = 0
        End If

        If blinkorder = 1 Then blinkorder = 0 Else blinkorder = 1
    End Sub

'[] COMMANDS []

    Public Sub CmdTotalQuan_Click()
        Call MenuCtrlTotalQuan_Click
    End Sub

    Private Sub TextCommand_Change()
        Select Case TextCommand.Text
            Case "1"
                Call MenuCtrlVote1_Click
            Case "2"
                Call MenuCtrlVote2_Click
            Case "3"
                Call MenuCtrlVote3_Click
            Case "4"
                Call MenuCtrlVote4_Click
            Case "5"
                Call MenuCtrlVote5_Click
            Case "6"
                Call MenuCtrlVote6_Click
            Case ""
                Call Refresher
            Case Else
                LabelStatusBar.Caption = "Vote ongoing...¡¡" & currentquan & " / " & totalquan & "¡¡Invalid input!"
        End Select

        TextCommand.Text = ""
        Call Refresher
    End Sub

'[] MENU []

    'CMD General...
    Public Sub MenuCtrlEXIT_Click()
        End
    End Sub
    Public Sub MenuCtrlAbout_Click()
        FormAbout.Show
        FormAbout.Top = (Screen.Height / 2)
        FormAbout.Left = (Screen.Width / 2)
        FormAbout.Width = 0
        FormAbout.Height = 0
        FormAbout.windowanimationtargettop = (Screen.Height / 2) - (7785 / 2)
        FormAbout.windowanimationtargetleft = (Screen.Width / 2) - (12930 / 2)
        FormAbout.windowanimationtargetwidth = 12930
        FormAbout.windowanimationtargetheight = 7785
    End Sub

    'CMD Controls...
    Public Sub MenuCtrlTotalQuan_Click()
        FormInputNumber.currentinputnumber = 1
        FormInputNumber.LabelInputNumber1.Caption = ">"
        FormInputNumber.LabelInputNumber2.Caption = ">"
        FormInputNumber.LabelInputNumber3.Caption = ">"
        FormInputNumber.LabelInputNumber4.Caption = ">"
        FormMainWindow.Enabled = False: FormInputNumber.Show
        FormInputNumber.Top = (Screen.Height / 2)
        FormInputNumber.Left = (Screen.Width / 2)
        FormInputNumber.Width = 0
        FormInputNumber.Height = 0
        FormInputNumber.windowanimationtargettop = (Screen.Height / 2) - (5895 / 2)
        FormInputNumber.windowanimationtargetleft = (Screen.Width / 2) - (6210 / 2)
        FormInputNumber.windowanimationtargetwidth = 6210
        FormInputNumber.windowanimationtargetheight = 5895
    End Sub

    Private Sub MenuCtrlStart_Click()
        Select Case status
            Case 0
                status = 1
                FormInputNumber.Hide

                MenuCtrlTotalQuan.Enabled = False
                MenuCtrlStart.Caption = "£¡¡¡Pause"
                MenuCtrlStart.Enabled = True
                MenuCtrlVote1.Enabled = True
                MenuCtrlVote2.Enabled = True
                MenuCtrlVote3.Enabled = True
                MenuCtrlVote4.Enabled = True
                MenuCtrlVote5.Enabled = True
                MenuCtrlVote6.Enabled = True
                MenuCtrlClear.Enabled = False
                CmdTotalQuan.Enabled = False
                TextCommand.Enabled = True
                TextCommand.BackColor = &HFFCC55
                TextCommand.SetFocus

                LabelStatusBar.Caption = "Vote started!¡¡" & currentquan & " / " & totalquan
            Case 1
                status = 0

                MenuCtrlTotalQuan.Enabled = False
                MenuCtrlStart.Caption = "¡ú¡¡Resume"
                MenuCtrlStart.Enabled = True
                MenuCtrlVote1.Enabled = False
                MenuCtrlVote2.Enabled = False
                MenuCtrlVote3.Enabled = False
                MenuCtrlVote4.Enabled = False
                MenuCtrlVote5.Enabled = False
                MenuCtrlVote6.Enabled = False
                MenuCtrlClear.Enabled = True
                CmdTotalQuan.Enabled = False
                TextCommand.Enabled = False
                TextCommand.BackColor = &HAA7700

                LabelStatusBar.Caption = "Vote paused. Press F5 to resume, F7 to abort and clear statistics."
        End Select

        Call Refresher
    End Sub

    Public Sub MenuCtrlVote1_Click()
        itemquan1 = itemquan1 + 1
        currentquan = currentquan + 1
        LabelStatusBar.Caption = "Vote ongoing...¡¡" & currentquan & " / " & totalquan & "¡¡A new vote to Candidate 1 !"
        Call Refresher
    End Sub
    Public Sub MenuCtrlVote2_Click()
        itemquan2 = itemquan2 + 1
        currentquan = currentquan + 1
        LabelStatusBar.Caption = "Vote ongoing...¡¡" & currentquan & " / " & totalquan & "¡¡A new vote to Candidate 2 !"
        Call Refresher
    End Sub
    Public Sub MenuCtrlVote3_Click()
        itemquan3 = itemquan3 + 1
        currentquan = currentquan + 1
        LabelStatusBar.Caption = "Vote ongoing...¡¡" & currentquan & " / " & totalquan & "¡¡A new vote to Candidate 3 !"
        Call Refresher
    End Sub
    Public Sub MenuCtrlVote4_Click()
        itemquan4 = itemquan4 + 1
        currentquan = currentquan + 1
        LabelStatusBar.Caption = "Vote ongoing...¡¡" & currentquan & " / " & totalquan & "¡¡A new vote to Candidate 4 !"
        Call Refresher
    End Sub
    Public Sub MenuCtrlVote5_Click()
        itemquan5 = itemquan5 + 1
        currentquan = currentquan + 1
        LabelStatusBar.Caption = "Vote ongoing...¡¡" & currentquan & " / " & totalquan & "¡¡A new vote to Candidate 5 !"
        Call Refresher
    End Sub
    Public Sub MenuCtrlVote6_Click()
        itemquan6 = itemquan6 + 1
        currentquan = currentquan + 1
        LabelStatusBar.Caption = "Vote ongoing...¡¡" & currentquan & " / " & totalquan & "¡¡A new vote to Candidate 6 !"
        Call Refresher
    End Sub

    Public Sub MenuCtrlClear_Click()
        currentquan = 1
        itemquan1 = 0
        itemquan2 = 0
        itemquan3 = 0
        itemquan4 = 0
        itemquan5 = 0
        itemquan6 = 0
        itemperc1 = 0
        itemperc2 = 0
        itemperc3 = 0
        itemperc4 = 0
        itemperc5 = 0
        itemperc6 = 0

        LabelItemNum1.BackStyle = 0
        LabelItemNum2.BackStyle = 0
        LabelItemNum3.BackStyle = 0
        LabelItemNum4.BackStyle = 0
        LabelItemNum5.BackStyle = 0
        LabelItemNum6.BackStyle = 0

        status = 0

        maxquan = 0
        blinkorder = 1

        MenuCtrlStart.Enabled = True
        MenuCtrlTotalQuan.Enabled = True
        MenuCtrlVote1.Enabled = False
        MenuCtrlVote2.Enabled = False
        MenuCtrlVote3.Enabled = False
        MenuCtrlVote4.Enabled = False
        MenuCtrlVote5.Enabled = False
        MenuCtrlVote6.Enabled = False
        MenuCtrlClear.Enabled = True
        CmdTotalQuan.Enabled = True
        TextCommand.Enabled = False
        TextCommand.BackColor = &HAA7700
        TimerMaxQuanBlink.Enabled = True

        Call Refresher
        Call TimerMaxQuanBlink_Timer

        MenuCtrlTotalQuan.Caption = "¡ù¡¡Quantity: " & totalquan
        MenuCtrlStart.Caption = "¡ð¡¡Start"
        LabelStatusBar.Caption = "Statistics cleared. Press F5 to start a new vote, F6 to change quantity."
    End Sub

'================================================================================

'================================================================================
