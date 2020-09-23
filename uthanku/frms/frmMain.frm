VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8g.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.Ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "YouTube Downloader v2.0.3"
   ClientHeight    =   9885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12735
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   9885
   ScaleWidth      =   12735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tEffect2 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   420
      Top             =   0
   End
   Begin VB.Timer tEffect 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picOption 
      BorderStyle     =   0  'None
      Height          =   3585
      Left            =   75
      ScaleHeight     =   3585
      ScaleWidth      =   90
      TabIndex        =   80
      Top             =   4755
      Visible         =   0   'False
      Width           =   90
      Begin VB.Frame Frame8 
         Caption         =   " Advanced Option Program "
         Height          =   3495
         Left            =   15
         TabIndex        =   81
         Top             =   -15
         Width           =   12555
         Begin VB.PictureBox backGroundOption 
            BorderStyle     =   0  'None
            Height          =   3240
            Left            =   60
            ScaleHeight     =   3240
            ScaleWidth      =   12450
            TabIndex        =   82
            Top             =   210
            Width           =   12450
            Begin VB.PictureBox picOption2 
               BorderStyle     =   0  'None
               Height          =   3120
               Left            =   690
               ScaleHeight     =   3120
               ScaleWidth      =   90
               TabIndex        =   105
               Top             =   75
               Visible         =   0   'False
               Width           =   90
               Begin VB.TextBox txtDefaultPath 
                  ForeColor       =   &H00404040&
                  Height          =   315
                  Left            =   3120
                  Locked          =   -1  'True
                  TabIndex        =   107
                  ToolTipText     =   "Default path video download"
                  Top             =   165
                  Width           =   7770
               End
               Begin thanku.XPStyleButton cmdChoisePath 
                  Height          =   330
                  Left            =   11055
                  TabIndex        =   108
                  ToolTipText     =   "Browser folder path..."
                  Top             =   165
                  Width           =   525
                  _ExtentX        =   926
                  _ExtentY        =   582
                  BTYPE           =   3
                  TX              =   "..."
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  COLTYPE         =   1
                  FOCUSR          =   0   'False
                  BCOL            =   14215660
                  BCOLO           =   14215660
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   99
                  MICON           =   "frmMain.frx":3452
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label20 
                  Caption         =   "Default path video dawnload:"
                  ForeColor       =   &H00404040&
                  Height          =   240
                  Left            =   135
                  TabIndex        =   106
                  Top             =   180
                  Width           =   3045
               End
            End
            Begin thanku.XPStyleButton cmdOterOption 
               Height          =   330
               Left            =   75
               TabIndex        =   104
               ToolTipText     =   "Show oter Option..."
               Top             =   2805
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   582
               BTYPE           =   3
               TX              =   ">>"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   0   'False
               BCOL            =   14215660
               BCOLO           =   14215660
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   99
               MICON           =   "frmMain.frx":35B4
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Shut down the PC wen all Downloads terminated..."
               ForeColor       =   &H00404040&
               Height          =   435
               Left            =   930
               MouseIcon       =   "frmMain.frx":3716
               MousePointer    =   99  'Custom
               TabIndex        =   89
               ToolTipText     =   "Enable/Disable sound Download...."
               Top             =   1965
               Width           =   3360
            End
            Begin VB.Frame Frame9 
               Caption         =   " Printers Setup "
               Height          =   3090
               Left            =   6345
               TabIndex        =   87
               Top             =   75
               Width           =   6060
               Begin VB.PictureBox backgroundPrinters 
                  BorderStyle     =   0  'None
                  Height          =   2820
                  Left            =   45
                  ScaleHeight     =   2820
                  ScaleWidth      =   5940
                  TabIndex        =   88
                  Top             =   225
                  Width           =   5940
                  Begin VB.ComboBox cmbFontSize 
                     ForeColor       =   &H00404040&
                     Height          =   330
                     ItemData        =   "frmMain.frx":3FE0
                     Left            =   3360
                     List            =   "frmMain.frx":4023
                     MouseIcon       =   "frmMain.frx":4076
                     MousePointer    =   99  'Custom
                     Style           =   2  'Dropdown List
                     TabIndex        =   101
                     Top             =   2235
                     Width           =   1230
                  End
                  Begin VB.ComboBox cmbFonts 
                     ForeColor       =   &H00404040&
                     Height          =   330
                     Left            =   915
                     MouseIcon       =   "frmMain.frx":4940
                     MousePointer    =   99  'Custom
                     Style           =   2  'Dropdown List
                     TabIndex        =   98
                     Top             =   1830
                     Width           =   4830
                  End
                  Begin VB.CheckBox CheckDisplayPrinter 
                     Caption         =   "Display printer Setup dialog before Printing..."
                     ForeColor       =   &H00404040&
                     Height          =   285
                     Left            =   180
                     MouseIcon       =   "frmMain.frx":520A
                     MousePointer    =   99  'Custom
                     TabIndex        =   97
                     ToolTipText     =   "Enable/Disable the Intellisense..."
                     Top             =   1380
                     Width           =   5475
                  End
                  Begin VB.ComboBox cmbPrinters 
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   6.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   300
                     ItemData        =   "frmMain.frx":5AD4
                     Left            =   1605
                     List            =   "frmMain.frx":5AD6
                     MouseIcon       =   "frmMain.frx":5AD8
                     MousePointer    =   99  'Custom
                     Style           =   2  'Dropdown List
                     TabIndex        =   94
                     Top             =   615
                     Width           =   3975
                  End
                  Begin VB.ComboBox cmbTopMargin 
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   6.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   300
                     ItemData        =   "frmMain.frx":63A2
                     Left            =   4350
                     List            =   "frmMain.frx":63AC
                     MouseIcon       =   "frmMain.frx":63BC
                     MousePointer    =   99  'Custom
                     TabIndex        =   93
                     Text            =   "cmbTopMargin"
                     Top             =   150
                     Width           =   1215
                  End
                  Begin VB.ComboBox cmbLeftMargin 
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   6.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00404040&
                     Height          =   300
                     ItemData        =   "frmMain.frx":6C86
                     Left            =   1590
                     List            =   "frmMain.frx":6C90
                     MouseIcon       =   "frmMain.frx":6CA0
                     MousePointer    =   99  'Custom
                     TabIndex        =   91
                     Text            =   "cmbLeftMargin"
                     Top             =   150
                     Width           =   1215
                  End
                  Begin VB.Label Label19 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Printer Font size:"
                     ForeColor       =   &H00404040&
                     Height          =   240
                     Left            =   1140
                     TabIndex        =   102
                     Top             =   2295
                     Width           =   2175
                  End
                  Begin VB.Label lblTotFont 
                     Alignment       =   1  'Right Justify
                     Caption         =   "n/a"
                     Height          =   225
                     Left            =   4815
                     TabIndex        =   100
                     ToolTipText     =   "Installed fonts..."
                     Top             =   2205
                     Width           =   900
                  End
                  Begin VB.Image Image7 
                     Height          =   360
                     Left            =   315
                     Picture         =   "frmMain.frx":756A
                     Top             =   2295
                     Width           =   360
                  End
                  Begin VB.Label Label18 
                     Caption         =   "Fonts:"
                     ForeColor       =   &H00404040&
                     Height          =   240
                     Left            =   180
                     TabIndex        =   99
                     Top             =   1890
                     Width           =   735
                  End
                  Begin VB.Line Line4 
                     BorderColor     =   &H00808080&
                     BorderStyle     =   3  'Dot
                     X1              =   5850
                     X2              =   5850
                     Y1              =   1275
                     Y2              =   2775
                  End
                  Begin VB.Line Line3 
                     BorderColor     =   &H00808080&
                     BorderStyle     =   3  'Dot
                     X1              =   3270
                     X2              =   5835
                     Y1              =   1275
                     Y2              =   1275
                  End
                  Begin VB.Label lblDefaultP 
                     Alignment       =   2  'Center
                     Caption         =   "n.a"
                     ForeColor       =   &H00404040&
                     Height          =   240
                     Left            =   105
                     TabIndex        =   96
                     ToolTipText     =   "Default Printer..."
                     Top             =   1005
                     Width           =   5805
                  End
                  Begin VB.Label Label12 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Printers:"
                     ForeColor       =   &H00404040&
                     Height          =   240
                     Left            =   240
                     TabIndex        =   95
                     Top             =   645
                     Width           =   1365
                  End
                  Begin VB.Label Label17 
                     Caption         =   "Top margin:"
                     ForeColor       =   &H00404040&
                     Height          =   240
                     Left            =   3090
                     TabIndex        =   92
                     Top             =   165
                     Width           =   1365
                  End
                  Begin VB.Label Label16 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Left margin:"
                     ForeColor       =   &H00404040&
                     Height          =   240
                     Left            =   240
                     TabIndex        =   90
                     Top             =   165
                     Width           =   1365
                  End
               End
            End
            Begin VB.CheckBox CheckStartUp 
               Caption         =   "Launch the Program when Windows start..."
               ForeColor       =   &H00404040&
               Height          =   435
               Left            =   930
               MouseIcon       =   "frmMain.frx":7C54
               MousePointer    =   99  'Custom
               TabIndex        =   86
               ToolTipText     =   "Enable/Disable sound Download...."
               Top             =   1440
               Width           =   3360
            End
            Begin VB.CheckBox CheckSnd 
               Caption         =   "Play sound wen downloads finish..."
               ForeColor       =   &H00404040&
               Height          =   435
               Left            =   930
               MouseIcon       =   "frmMain.frx":851E
               MousePointer    =   99  'Custom
               TabIndex        =   85
               ToolTipText     =   "Enable/Disable sound Download...."
               Top             =   930
               Width           =   3360
            End
            Begin VB.CheckBox CheckAutoFind 
               Caption         =   "Enable Autofind on StartUp..."
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   930
               MouseIcon       =   "frmMain.frx":8DE8
               MousePointer    =   99  'Custom
               TabIndex        =   84
               ToolTipText     =   "Enable/Disable the Autofind Video on Application Start...."
               Top             =   570
               Width           =   3360
            End
            Begin VB.CheckBox ChckIntelliSense 
               Caption         =   "Activate IntelliSense..."
               ForeColor       =   &H00404040&
               Height          =   285
               Left            =   930
               MouseIcon       =   "frmMain.frx":96B2
               MousePointer    =   99  'Custom
               TabIndex        =   83
               ToolTipText     =   "Enable/Disable the Intellisense..."
               Top             =   195
               Width           =   3360
            End
            Begin VB.Label lbltestFont 
               Caption         =   "Cantami o diva del pelide Achille"
               Height          =   480
               Left            =   1485
               TabIndex        =   103
               ToolTipText     =   "Preview font text..."
               Top             =   2655
               Width           =   4710
            End
            Begin VB.Image Image8 
               Height          =   360
               Left            =   885
               Picture         =   "frmMain.frx":9F7C
               Top             =   2700
               Width           =   360
            End
            Begin VB.Image Image6 
               Height          =   480
               Left            =   5445
               Picture         =   "frmMain.frx":A666
               Top             =   105
               Width           =   480
            End
            Begin VB.Image Image5 
               Height          =   720
               Left            =   0
               Picture         =   "frmMain.frx":10EB8
               Top             =   45
               Width           =   720
            End
         End
      End
   End
   Begin VB.CheckBox CheckHash 
      Caption         =   "Calculate CRC32 and Hash of the Video file"
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   2085
      MouseIcon       =   "frmMain.frx":12B82
      MousePointer    =   99  'Custom
      TabIndex        =   79
      ToolTipText     =   "Calculate CRC32, MD4, MD5 and Hash of video file wen download finish..."
      Top             =   6855
      Width           =   4860
   End
   Begin VB.ListBox lstFileName 
      Height          =   270
      Left            =   12735
      TabIndex        =   78
      Top             =   -120
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox txtConvert 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   15120
      TabIndex        =   75
      Text            =   "http://www.youtube.com"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Timer TimerDB 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   16650
      Top             =   2205
   End
   Begin VB.TextBox txtPing 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   15135
      TabIndex        =   74
      Top             =   3420
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   0
      Left            =   16350
      Picture         =   "frmMain.frx":1344C
      ScaleHeight     =   315
      ScaleWidth      =   360
      TabIndex        =   73
      Top             =   450
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   1
      Left            =   15990
      Picture         =   "frmMain.frx":139D6
      ScaleHeight     =   315
      ScaleWidth      =   360
      TabIndex        =   72
      Top             =   465
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   2
      Left            =   16725
      Picture         =   "frmMain.frx":13F60
      ScaleHeight     =   315
      ScaleWidth      =   360
      TabIndex        =   71
      Top             =   435
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   3
      Left            =   15615
      Picture         =   "frmMain.frx":144EA
      ScaleHeight     =   315
      ScaleWidth      =   360
      TabIndex        =   70
      Top             =   450
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox Picture8 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   6810
      ScaleHeight     =   375
      ScaleWidth      =   765
      TabIndex        =   64
      Top             =   6795
      Width           =   765
      Begin thanku.XPStyleButton cmdCopyToClipboard 
         Height          =   330
         Left            =   240
         TabIndex        =   65
         ToolTipText     =   "Export the Hashed result to file..."
         Top             =   0
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   582
         BTYPE           =   9
         TX              =   ""
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "frmMain.frx":14A74
         PICN            =   "frmMain.frx":14BD6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin ComctlLib.ListView lstTitles 
      Height          =   2595
      Left            =   4440
      TabIndex        =   21
      Top             =   2100
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   4577
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   4210752
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Video Title"
         Object.Width           =   7832
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Video Url"
         Object.Width           =   7832
      EndProperty
   End
   Begin VB.ComboBox cmbHash 
      Height          =   330
      Left            =   7590
      MouseIcon       =   "frmMain.frx":15AB0
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   60
      Top             =   6810
      Width           =   5070
   End
   Begin VB.PictureBox cmdAbout 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   12150
      MouseIcon       =   "frmMain.frx":1637A
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":164CC
      ScaleHeight     =   480
      ScaleWidth      =   510
      TabIndex        =   58
      ToolTipText     =   "About..."
      Top             =   120
      Width           =   510
   End
   Begin thanku.Downloader Dloader 
      Left            =   16590
      Top             =   855
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Frame Frame7 
      Height          =   1215
      Left            =   7605
      TabIndex        =   52
      Top             =   8295
      Width           =   5040
      Begin VB.PictureBox Picture7 
         BorderStyle     =   0  'None
         Height          =   1050
         Left            =   15
         ScaleHeight     =   1050
         ScaleWidth      =   4980
         TabIndex        =   53
         Top             =   120
         Width           =   4980
         Begin thanku.XPStyleButton cmdClose 
            Height          =   420
            Left            =   1020
            TabIndex        =   54
            ToolTipText     =   "Close this Program..."
            Top             =   570
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   741
            BTYPE           =   3
            TX              =   "&Close Program"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   4210752
            FCOLO           =   4210752
            MCOL            =   16711935
            MPTR            =   99
            MICON           =   "frmMain.frx":17196
            PICN            =   "frmMain.frx":172F8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   1
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin thanku.XPStyleButton cmdAdvanced 
            Height          =   420
            Left            =   3315
            TabIndex        =   55
            ToolTipText     =   "Advanced Option Program..."
            Top             =   570
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   741
            BTYPE           =   3
            TX              =   "&Option"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   4210752
            FCOLO           =   4210752
            MCOL            =   16711935
            MPTR            =   99
            MICON           =   "frmMain.frx":17FD2
            PICN            =   "frmMain.frx":18134
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   1
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin thanku.utcDown utcDown 
            Height          =   720
            Left            =   60
            TabIndex        =   63
            ToolTipText     =   "Download status..."
            Top             =   180
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   1270
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Dot
            X1              =   855
            X2              =   2505
            Y1              =   60
            Y2              =   60
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            BorderStyle     =   3  'Dot
            X1              =   855
            X2              =   855
            Y1              =   90
            Y2              =   1050
         End
      End
   End
   Begin VB.Timer tDownloaded 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   16650
      Top             =   2595
   End
   Begin VB.Timer tlastSearch 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   16635
      Top             =   1800
   End
   Begin VB.Frame Frame6 
      Caption         =   "  Dropped  "
      ForeColor       =   &H00404040&
      Height          =   1440
      Left            =   6090
      TabIndex        =   50
      Top             =   5340
      Width           =   1485
      Begin VB.PictureBox DropArea 
         BorderStyle     =   0  'None
         Height          =   1185
         Left            =   45
         OLEDropMode     =   1  'Manual
         Picture         =   "frmMain.frx":18E0E
         ScaleHeight     =   1185
         ScaleWidth      =   1410
         TabIndex        =   51
         ToolTipText     =   "Drag a valid *.flv Files or Folder video  here..."
         Top             =   210
         Width           =   1410
      End
   End
   Begin VB.TextBox lblDescription 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   990
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   48
      Text            =   "frmMain.frx":1965F
      Top             =   8520
      Width           =   7245
   End
   Begin VB.Frame Frame5 
      Height          =   1215
      Left            =   165
      TabIndex        =   40
      Top             =   7080
      Width           =   12480
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   990
         Left            =   30
         ScaleHeight     =   990
         ScaleWidth      =   12405
         TabIndex        =   41
         Top             =   180
         Width           =   12405
         Begin thanku.XPStyleButton cmdLike 
            Height          =   315
            Left            =   120
            TabIndex        =   61
            ToolTipText     =   "I like this..."
            Top             =   90
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   99
            MICON           =   "frmMain.frx":19663
            PICN            =   "frmMain.frx":197C5
            PICH            =   "frmMain.frx":1A1B8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   4
            NGREY           =   0   'False
            FX              =   0
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.ComboBox cmbExtension 
            ForeColor       =   &H00404040&
            Height          =   330
            Left            =   10635
            MouseIcon       =   "frmMain.frx":1AC92
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   46
            ToolTipText     =   "Possible extension to Download and convert video on the fly..."
            Top             =   615
            Width           =   1680
         End
         Begin VB.TextBox txtVideoTitle 
            Alignment       =   2  'Center
            ForeColor       =   &H00404040&
            Height          =   300
            Left            =   2970
            TabIndex        =   45
            Text            =   "n.a"
            ToolTipText     =   "If you want, enter the video title manually without the extension..."
            Top             =   645
            Width           =   5100
         End
         Begin thanku.XPStyleButton cmdUseLink 
            Height          =   420
            Left            =   9945
            TabIndex        =   42
            ToolTipText     =   "Download the selected video..."
            Top             =   30
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   741
            BTYPE           =   3
            TX              =   "&Download Video"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   4210752
            FCOLO           =   4210752
            MCOL            =   16711935
            MPTR            =   99
            MICON           =   "frmMain.frx":1B55C
            PICN            =   "frmMain.frx":1B6BE
            PICH            =   "frmMain.frx":1C398
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   1
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin thanku.XPStyleButton cmdNavigateURL 
            Height          =   420
            Left            =   6960
            TabIndex        =   56
            ToolTipText     =   "Open the selected video to yours default Browser..."
            Top             =   30
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   741
            BTYPE           =   3
            TX              =   "&Open current Video"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   4210752
            FCOLO           =   4210752
            MCOL            =   16711935
            MPTR            =   99
            MICON           =   "frmMain.frx":1D072
            PICN            =   "frmMain.frx":1D1D4
            PICH            =   "frmMain.frx":1DEAE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   1
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin thanku.XPStyleButton cmdConvert 
            Height          =   420
            Left            =   4065
            TabIndex        =   59
            ToolTipText     =   "Start the conversion of the unloaded video..."
            Top             =   30
            Width           =   2745
            _ExtentX        =   4842
            _ExtentY        =   741
            BTYPE           =   3
            TX              =   "&Start Convertion"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   4210752
            FCOLO           =   4210752
            MCOL            =   16711935
            MPTR            =   99
            MICON           =   "frmMain.frx":1EB88
            PICN            =   "frmMain.frx":1ECEA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   1
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin thanku.XPStyleButton cmdUnlike 
            Height          =   315
            Left            =   870
            TabIndex        =   62
            ToolTipText     =   "I dislike this..."
            Top             =   90
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   99
            MICON           =   "frmMain.frx":1F9C4
            PICN            =   "frmMain.frx":1FB26
            PICH            =   "frmMain.frx":201FF
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   4
            NGREY           =   0   'False
            FX              =   0
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin thanku.XPStyleButton cmdRegistration 
            Height          =   420
            Left            =   1560
            TabIndex        =   66
            ToolTipText     =   "Enter o Send request of Key Activation..."
            Top             =   15
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   741
            BTYPE           =   3
            TX              =   "&Registration"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   4210752
            FCOLO           =   4210752
            MCOL            =   16711935
            MPTR            =   99
            MICON           =   "frmMain.frx":209B4
            PICN            =   "frmMain.frx":20B16
            PICH            =   "frmMain.frx":213F0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   1
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Image imgVideoSize 
            Height          =   240
            Left            =   8160
            MouseIcon       =   "frmMain.frx":21CCA
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":21E1C
            ToolTipText     =   "Get the Size of current selected Video...."
            Top             =   675
            Width           =   240
         End
         Begin VB.Label Label14 
            Caption         =   "Video URL Download:"
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   8490
            TabIndex        =   47
            Top             =   675
            Width           =   2160
         End
         Begin VB.Label Label8 
            Caption         =   "Insert Video Title:"
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   900
            TabIndex        =   44
            Top             =   675
            Width           =   2100
         End
         Begin VB.Image Image3 
            Height          =   240
            Left            =   495
            Picture         =   "frmMain.frx":221A6
            Top             =   675
            Width           =   240
         End
         Begin VB.Line Line14 
            BorderColor     =   &H00808080&
            BorderStyle     =   3  'Dot
            X1              =   120
            X2              =   480
            Y1              =   795
            Y2              =   795
         End
         Begin VB.Line Line13 
            BorderColor     =   &H00808080&
            BorderStyle     =   3  'Dot
            X1              =   120
            X2              =   120
            Y1              =   540
            Y2              =   825
         End
         Begin VB.Line Line12 
            BorderColor     =   &H00808080&
            BorderStyle     =   3  'Dot
            X1              =   120
            X2              =   12300
            Y1              =   540
            Y2              =   540
         End
      End
   End
   Begin VB.Timer tStart 
      Enabled         =   0   'False
      Left            =   16635
      Top             =   1425
   End
   Begin VB.Frame Frame4 
      Caption         =   " Video Conversion Format "
      ForeColor       =   &H00404040&
      Height          =   1440
      Left            =   7590
      TabIndex        =   34
      Top             =   5340
      Width           =   5055
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   15
         ScaleHeight     =   1215
         ScaleWidth      =   5010
         TabIndex        =   35
         Top             =   180
         Width           =   5010
         Begin VB.OptionButton optVideo 
            Caption         =   "Format 4:3"
            Height          =   255
            Index           =   1
            Left            =   2895
            MouseIcon       =   "frmMain.frx":22530
            MousePointer    =   99  'Custom
            TabIndex        =   69
            Top             =   870
            Value           =   -1  'True
            Width           =   2070
         End
         Begin VB.OptionButton optVideo 
            Caption         =   "Format 16:9"
            Height          =   255
            Index           =   0
            Left            =   825
            MouseIcon       =   "frmMain.frx":22682
            MousePointer    =   99  'Custom
            TabIndex        =   68
            Top             =   870
            Width           =   2010
         End
         Begin thanku.utcWait utcWait 
            Height          =   360
            Left            =   150
            TabIndex        =   57
            ToolTipText     =   "Working conversion..."
            Top             =   750
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
         End
         Begin VB.ComboBox cmbFormato 
            ForeColor       =   &H00404040&
            Height          =   330
            ItemData        =   "frmMain.frx":227D4
            Left            =   600
            List            =   "frmMain.frx":227D6
            MouseIcon       =   "frmMain.frx":227D8
            MousePointer    =   99  'Custom
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   135
            Width           =   4350
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Select the video Output format..."
            ForeColor       =   &H00404040&
            Height          =   285
            Left            =   630
            TabIndex        =   67
            Top             =   540
            Width           =   4290
         End
         Begin VB.Image Image4 
            Height          =   480
            Left            =   -30
            Picture         =   "frmMain.frx":230A2
            Top             =   60
            Width           =   480
         End
      End
   End
   Begin VB.TextBox txtVideoPicture 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   15435
      TabIndex        =   33
      Top             =   3705
      Visible         =   0   'False
      Width           =   1620
   End
   Begin thanku.ShowImage SImage 
      Height          =   1260
      Left            =   165
      TabIndex        =   30
      ToolTipText     =   "Preview video picture..."
      Top             =   5475
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   2223
      BorderStyle     =   0
   End
   Begin VB.Frame Frame2 
      Height          =   525
      Left            =   165
      TabIndex        =   24
      Top             =   4710
      Width           =   12465
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   30
         ScaleHeight     =   360
         ScaleWidth      =   12390
         TabIndex        =   25
         Top             =   135
         Width           =   12390
         Begin thanku.XPStyleButton cmdNext 
            Height          =   300
            Left            =   3585
            TabIndex        =   26
            ToolTipText     =   "Next page..."
            Top             =   15
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   529
            BTYPE           =   9
            TX              =   ""
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   99
            MICON           =   "frmMain.frx":23D6C
            PICN            =   "frmMain.frx":23ECE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin thanku.XPStyleButton cmdPrev 
            Height          =   300
            Left            =   15
            TabIndex        =   28
            ToolTipText     =   "Preview page..."
            Top             =   15
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   529
            BTYPE           =   9
            TX              =   ""
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   99
            MICON           =   "frmMain.frx":24DA8
            PICN            =   "frmMain.frx":24F0A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lbltotalxPage 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Titles x page"
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   4245
            TabIndex        =   29
            Top             =   45
            Width           =   8115
         End
         Begin VB.Label lblpages 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Page 0 of 0"
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   525
            TabIndex        =   27
            Top             =   45
            Width           =   3075
         End
      End
   End
   Begin VB.ListBox lstURLs 
      Height          =   270
      Left            =   15150
      TabIndex        =   23
      ToolTipText     =   "Double Click to get Vido Info..."
      Top             =   2835
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.ListBox lstUrlTitles 
      Height          =   270
      Left            =   15150
      TabIndex        =   22
      Top             =   2535
      Visible         =   0   'False
      Width           =   1290
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash SF 
      Height          =   2850
      Left            =   165
      TabIndex        =   19
      ToolTipText     =   "If you don't see the Video preview click on the Button 'Open current Video'..."
      Top             =   1845
      Width           =   4200
      _cx             =   7408
      _cy             =   5027
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   "000000"
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8505
      ScaleHeight     =   255
      ScaleWidth      =   4200
      TabIndex        =   15
      Top             =   9615
      Width           =   4200
      Begin ComctlLib.ProgressBar PB 
         Height          =   210
         Left            =   30
         TabIndex        =   16
         Top             =   15
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   370
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblPercentage 
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   3750
         TabIndex        =   17
         Top             =   0
         Width           =   450
      End
   End
   Begin ComctlLib.StatusBar STB 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   14
      Top             =   9555
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   14887
            MinWidth        =   14887
            Picture         =   "frmMain.frx":25DE4
            Text            =   "Status ready..."
            TextSave        =   "Status ready..."
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Generic msg ..."
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7479
            MinWidth        =   7479
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   705
      Left            =   150
      TabIndex        =   11
      Top             =   810
      Width           =   12510
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   30
         ScaleHeight     =   525
         ScaleWidth      =   12435
         TabIndex        =   12
         Top             =   135
         Width           =   12435
         Begin VB.TextBox txtSearch 
            ForeColor       =   &H00404040&
            Height          =   315
            Left            =   1905
            TabIndex        =   43
            Top             =   105
            Width           =   8160
         End
         Begin thanku.XPStyleButton cmdFind 
            Height          =   465
            Left            =   10275
            TabIndex        =   18
            ToolTipText     =   "Find this titles..."
            Top             =   30
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   820
            BTYPE           =   3
            TX              =   "Find &Title"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   4210752
            FCOLO           =   4210752
            MCOL            =   16711935
            MPTR            =   99
            MICON           =   "frmMain.frx":25FE6
            PICN            =   "frmMain.frx":26148
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   1
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Image imgPasteFromClipboard 
            Height          =   240
            Left            =   165
            MouseIcon       =   "frmMain.frx":26A22
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":26B74
            ToolTipText     =   "Paste from Clipboard a valid youtube URL..."
            Top             =   135
            Width           =   240
         End
         Begin VB.Label Label7 
            Caption         =   "Video Title:"
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   600
            TabIndex        =   13
            Top             =   135
            Width           =   1290
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Download Status"
      ForeColor       =   &H00404040&
      Height          =   1440
      Left            =   2100
      TabIndex        =   0
      Top             =   5340
      Width           =   3960
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   30
         ScaleHeight     =   1215
         ScaleWidth      =   3870
         TabIndex        =   1
         Top             =   180
         Width           =   3870
         Begin VB.Label lblElapced 
            BackStyle       =   0  'Transparent
            Caption         =   "00:00:00"
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   2235
            TabIndex        =   39
            ToolTipText     =   "Work Time..."
            Top             =   915
            Width           =   1620
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Elapsed Time:"
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   300
            TabIndex        =   38
            Top             =   915
            Width           =   1800
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   30
            Picture         =   "frmMain.frx":26EFE
            Top             =   -15
            Width           =   480
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Remaining:"
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   300
            TabIndex        =   7
            Top             =   645
            Width           =   1800
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Download Now:"
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   300
            TabIndex        =   6
            Top             =   375
            Width           =   1800
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "File Size:"
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   300
            TabIndex        =   5
            Top             =   135
            Width           =   1815
         End
         Begin VB.Label lblSaved 
            BackStyle       =   0  'Transparent
            Caption         =   "0 bytes"
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   2235
            TabIndex        =   4
            ToolTipText     =   "Byts to unloaded..."
            Top             =   375
            Width           =   1635
         End
         Begin VB.Label lblSize2 
            BackStyle       =   0  'Transparent
            Caption         =   "0 bytes"
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   2235
            TabIndex        =   3
            ToolTipText     =   "Total file size of the current Video..."
            Top             =   135
            Width           =   1635
         End
         Begin VB.Label lblOf 
            BackStyle       =   0  'Transparent
            Caption         =   "0 bytes"
            ForeColor       =   &H00404040&
            Height          =   270
            Left            =   2235
            TabIndex        =   2
            ToolTipText     =   "Bytes remaining to downloaded..."
            Top             =   645
            Width           =   1620
         End
      End
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   15975
      Top             =   870
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Image imgScanFolders 
      Height          =   480
      Left            =   12180
      MouseIcon       =   "frmMain.frx":27BC8
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":27D1A
      ToolTipText     =   "Scan your Folder and Subfolders..."
      Top             =   1575
      Width           =   480
   End
   Begin VB.Image cmdList 
      Height          =   480
      Left            =   11625
      MouseIcon       =   "frmMain.frx":28B5C
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":28CAE
      ToolTipText     =   "Open the path of Video URL..."
      Top             =   1575
      Width           =   480
   End
   Begin VB.Label lblPing 
      Alignment       =   2  'Center
      Caption         =   "n/a"
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   8025
      TabIndex        =   77
      ToolTipText     =   "Response..."
      Top             =   1575
      Width           =   4050
   End
   Begin VB.Label lblSend 
      Alignment       =   2  'Center
      Caption         =   "n/a"
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   4800
      TabIndex        =   76
      ToolTipText     =   "Attempt..."
      Top             =   1575
      Width           =   3360
   End
   Begin VB.Image img 
      Height          =   240
      Left            =   4425
      MouseIcon       =   "frmMain.frx":29578
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":296CA
      ToolTipText     =   "Server status... click to ping youtube.com!"
      Top             =   1590
      Width           =   240
   End
   Begin VB.Image imgs 
      Height          =   270
      Index           =   4
      Left            =   16665
      Picture         =   "frmMain.frx":29C54
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgs 
      Height          =   270
      Index           =   3
      Left            =   16335
      Picture         =   "frmMain.frx":2A24E
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgs 
      Height          =   270
      Index           =   2
      Left            =   16035
      Picture         =   "frmMain.frx":2A848
      Top             =   105
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      X1              =   6480
      X2              =   7470
      Y1              =   8415
      Y2              =   8415
   End
   Begin VB.Label Label15 
      Caption         =   "User note of this Video"
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   3975
      TabIndex        =   49
      Top             =   8310
      Width           =   2610
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      X1              =   165
      X2              =   3885
      Y1              =   8430
      Y2              =   8430
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      X1              =   165
      X2              =   165
      Y1              =   8445
      Y2              =   9570
   End
   Begin VB.Label lblvideoTitle 
      Alignment       =   2  'Center
      Caption         =   "n.a"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   4440
      TabIndex        =   37
      ToolTipText     =   "Video Title..."
      Top             =   1845
      Width           =   7065
   End
   Begin VB.Label lblSize 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Size: 0 kb"
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   180
      TabIndex        =   32
      Top             =   6840
      Width           =   1845
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   1335
      Left            =   150
      Top             =   5445
      Width           =   1860
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Preview Image"
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   285
      TabIndex        =   31
      Top             =   5235
      Width           =   1695
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Video Preview"
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   1410
      TabIndex        =   20
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "This is my 2nd edition of this tool to download video from Youtube and converter! I hope you like it..."
      ForeColor       =   &H00404040&
      Height          =   660
      Left            =   7965
      TabIndex        =   10
      Top             =   75
      Width           =   4065
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "YouTube Downloader v2.0.3"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   2790
      TabIndex        =   9
      Top             =   150
      Width           =   5235
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "YouTube Downloader v2.0.3"
      BeginProperty Font 
         Name            =   "Bodoni MT Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   2730
      TabIndex        =   8
      Top             =   165
      Width           =   5235
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   105
      Picture         =   "frmMain.frx":2AE42
      Top             =   0
      Width           =   2400
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000003&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   825
      Left            =   -450
      Top             =   -45
      Width           =   13245
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuContextMenu 
         Caption         =   "#1"
         Index           =   0
      End
      Begin VB.Menu mnuContextMenu 
         Caption         =   "#2"
         Index           =   1
      End
      Begin VB.Menu mnuContextMenu 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuContextMenu 
         Caption         =   "Option list"
         Index           =   3
         Begin VB.Menu mnuContextMenuOption 
            Caption         =   "Print this list"
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim isDays As String: Dim istrMessage As String

Private strPing As Integer

Private CaptionForm As String

' .... Init the Class ToolTip for the lstTracks
Private lItemIndex As Long

' .... Disabled Closed
Public readyToClose As Boolean

' .... File in download?
Private Downloading As Boolean

' .... Send command
Private Declare Function ShellExecute Lib "SHELL32.DLL" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Private Enum SWSHOW
    SW_SHOW = 5
    SW_SHOWDEFAULT = 10
    SW_SHOWMAXIMIZED = 3
    SW_SHOWMINIMIZED = 2
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_SHOWNOACTIVATE = 4
    SW_SHOWNORMAL = 1
End Enum

' .... Inizializzo la Classe per il Ping del Server
Private WithEvents PingServer As clsPingServer
Attribute PingServer.VB_VarHelpID = -1

' .... Init my Class dll
Private WithEvents YouTubeThanks As YouTubeThanks
Attribute YouTubeThanks.VB_VarHelpID = -1

' .... Move the file part of video file to the Recycle Bin
Private Type SHFILEOPTSTRUCT
  hWnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAnyOperationsAborted As Long
  hNameMappings As Long
  lpszProgressTitle As Long
End Type

Private Declare Function SHFileOperation Lib "SHELL32.DLL" _
  Alias "SHFileOperationA" (lpFileOp As SHFILEOPTSTRUCT) As Long

Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40

Private STOP_PRESSED As Boolean

' .... Time video download count
Private Hours As Integer
Private Minutes As Integer
Private Seconds As Integer
Private Days As Integer
Private AddMinutes As Boolean
Private AddHours As Boolean
Private AddDays As Boolean

Private fln As String

Private Enum DISP_BYTES_FORMAT
    DISP_BYTES_LONG
    DISP_BYTES_SHORT
    DISP_BYTES_ALL
End Enum

Private Const K_B = 1024#
Private Const M_B = (K_B * 1024#) ' MegaBytes
Private Const G_B = (M_B * 1024#) ' GigaBytes
Private Const T_B = (G_B * 1024#) ' TeraBytes
Private Const P_B = (T_B * 1024#) ' PetaBytes
Private Const E_B = (P_B * 1024#) ' ExaBytes
Private Const Z_B = (E_B * 1024#) ' ZettaBytes
Private Const Y_B = (Z_B * 1024#) ' YottaBytes

Private Video_Folder As String
Private itmSelected As Integer
Private pPage As Integer
Private lastPage As Integer

' .... For the ListView
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private itmX As ListItem

Private Const LVM_FIRST = &H1000
Private Const LVIF_STATE = &H8

Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55

Private Const LVS_EX_FULLROWSELECT = &H20
Private Const LVS_EX_GRIDLINES = &H1
Private Const LVS_EX_CHECKBOXES As Long = &H4
Private Const LVS_EX_TRACKSELECT = &H8
Private Const LVS_EX_ONECLICKACTIVATE = &H40
Private Const LVS_EX_TWOCLICKACTIVATE = &H80
Private Const LVS_EX_SUBITEMIMAGES = &H2
 
Private Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Private Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Private Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Private Const LVIS_STATEIMAGEMASK As Long = &HF000

Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Type LVITEM
    mask         As Long
    iItem        As Long
    iSubItem     As Long
    State        As Long
    stateMask    As Long
    pszText      As String
    cchTextMax   As Long
    iImage       As Long
    lParam       As Long
    iIndent      As Long
End Type

' .... Servis Constants
Private Const sQuote As String = """"

' .... Intellisence
Private WasDelete As Boolean

Private Type iSense
    sOut As String * 50
End Type

' .... Send Text file to textBox
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

Private Const WM_SETTEXT = &HC
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE

' .... Load *.CUR from RES file
Private Declare Function SetSystemCursor Lib "user32" (ByVal hcur As Long, ByVal ID As Long) As Long
Private Declare Function GetCursor Lib "user32.dll" () As Long
Private Declare Function CopyCursor Lib "user32" Alias "CopyIcon" (ByVal hcur As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
 
Private Const OCR_NORMAL As Long = 32512
Private hOldCursor As Long
Private cur As Long: Private hCursor As Long

Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const SM_CXDRAG = 68
Private Const SM_CYDRAG = 69
Private Const IDC_ARROW = 32512&
Private Const IDC_HAND = 32649&

' .... Disable the closeBox Window Form
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&

' .... ToolTip Class
Const LVM_HITTEST = LVM_FIRST + 18

Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private Type LVHITTESTINFO
   pt As POINTAPI
   flags As Long
   iItem As Long
   iSubItem As Long
End Type

Dim TT As CTooltip
Dim m_lCurItemIndex As Long
Private Function GetSizeBytes(Dec As Variant, Optional DispBytesFormat As DISP_BYTES_FORMAT = DISP_BYTES_ALL) As String
    Dim DispLong As String: Dim DispShort As String: Dim s As String
    If DispBytesFormat <> DISP_BYTES_SHORT Then DispLong = FormatNumber(Dec, 0) & " bytes" Else DispLong = ""
    If DispBytesFormat <> DISP_BYTES_LONG Then
        If Dec > Y_B Then
            DispShort = FormatNumber(Dec / Y_B, 2) & " Yb"
        ElseIf Dec > Z_B Then
            DispShort = FormatNumber(Dec / Z_B, 2) & " Zb"
        ElseIf Dec > E_B Then
            DispShort = FormatNumber(Dec / E_B, 2) & " Eb"
        ElseIf Dec > P_B Then
            DispShort = FormatNumber(Dec / P_B, 2) & " Pb"
        ElseIf Dec > T_B Then
            DispShort = FormatNumber(Dec / T_B, 2) & " Tb"
        ElseIf Dec > G_B Then
            DispShort = FormatNumber(Dec / G_B, 2) & " Gb"
        ElseIf Dec > M_B Then
            DispShort = FormatNumber(Dec / M_B, 2) & " Mb"
        ElseIf Dec > K_B Then
            DispShort = FormatNumber(Dec / K_B, 2) & " Kb"
        Else
            DispShort = FormatNumber(Dec, 0) & " bytes"
        End If
    Else
        DispShort = ""
    End If
    Select Case DispBytesFormat
        Case DISP_BYTES_SHORT:
            GetSizeBytes = DispShort
        Case DISP_BYTES_LONG:
            GetSizeBytes = DispLong
        Case Else:
            GetSizeBytes = DispLong & " (" & DispShort & ")"
    End Select
End Function

Private Sub CheckStartUp_Click()
    On Local Error Resume Next
    If CheckStartUp.value = 1 Then
        Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "YouTube Downloader v2.0.3", App.Path + "\" + App.EXEName + ".exe")
    ElseIf CheckStartUp.value = 0 Then
        Call DeleteValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "YouTube Downloader v2.0.3")
    End If
End Sub


Private Sub cmbExtension_Click()
    On Local Error GoTo ErrorHandler
    lblSize2.Caption = GetSizeBytes(YouTubeThanks.GetVideoFileSize(cmbExtension.ListIndex), DISP_BYTES_SHORT)
Exit Sub
ErrorHandler:
    Err.Clear
End Sub


Private Sub cmbFonts_Click()
    On Local Error Resume Next
    lbltestFont.FontName = cmbFonts.List(cmbFonts.ListIndex)
    lbltestFont.FontBold = False: lbltestFont.FontItalic = False
    lbltestFont.FontStrikethru = False: lbltestFont.FontUnderline = False
End Sub


Private Sub cmbFontSize_Click()
    On Local Error Resume Next
    lbltestFont.FontSize = cmbFontSize.List(cmbFontSize.ListIndex)
    lbltestFont.FontBold = False: lbltestFont.FontItalic = False
    lbltestFont.FontStrikethru = False: lbltestFont.FontUnderline = False
End Sub


Private Sub cmbFormato_Click()
    On Local Error Resume Next
    If Mid(cmbFormato.List(cmbFormato.ListIndex), 1, 1) = "-" Then Exit Sub
    
    estensione = StripLeft(cmbFormato.List(cmbFormato.ListIndex), "(", False)
    estensione = Mid$(estensione, 1, Len(estensione) - 1)
    
Exit Sub
End Sub

Private Sub cmbPrinters_Click()
    Dim x As Printer
    On Error Resume Next
        For Each x In Printers
            If Len(x.DeviceName) > 0 And x.DeviceName = cmbPrinters.Text Then
                   If SetDefaultPrinter(x) = True Then
                   End If
                Exit For
            End If
        Next
    lblDefaultP.Caption = GetDefaultPrinter()
End Sub


Private Sub cmdAbout_Click()
    YouTubeThanks.About
End Sub

Private Sub cmdAdvanced_Click()
    tEffect.Enabled = True
End Sub

Private Sub cmdChoisePath_Click()
    Dim strFolder As String
    strFolder = FolderBrowse(Me.hWnd, "Select the default folder:")
    If Len(strFolder) = 0 Then Exit Sub
    
    If Mid$(strFolder, 3, 3) <> "\" Then
        txtDefaultPath.Text = strFolder & "\"
    Else
        txtDefaultPath.Text = strFolder
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdConvert_Click()
    Dim strFile As String
    
    strFile = App.Path + "\Downloads\" + Mid$(URLVideoTitle, 1, Len(URLVideoTitle) - 4) + "\" + URLVideoTitle
    
    If FileExists(strFile) = False Then
            MsgBox "The Video file " & Mid$(URLVideoTitle, 1, Len(URLVideoTitle) - 4) & ", doesen't exist!", vbExclamation, App.Title
        Exit Sub
    End If
    
    If Mid(cmbFormato.List(cmbFormato.ListIndex), 1, 1) = "-" Then
            MsgBox "Select a valid format conversion...", vbExclamation, App.Title
        Exit Sub
    End If
    
    On Local Error GoTo ErrorHandler
    
    ' .... Request conversion video YES or NOT
    If MsgBox("Convert the Video: " & Mid$(URLVideoTitle, 1, Len(URLVideoTitle) - 4) & ", to " & Mid$(URLVideoTitle, 1, Len(URLVideoTitle) - 4) & "." & estensione & "?", _
    vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
    
    utcWait.Start = True
    
    If estensione = "mpg" Then
        If optVideo.Item(1).value Then
                If YouTubeThanks.StartConversion(strFile, XVID, DvD43, estensione, SW_SHOWNORMAL) Then:
            'Exit Sub
        ElseIf optVideo.Item(0).value Then
                If YouTubeThanks.StartConversion(strFile, XVID, DvD169, estensione, SW_SHOWNORMAL) Then:
            'Exit Sub
        End If
    ElseIf estensione = "avi" Then
            If YouTubeThanks.StartConversion(strFile, XVID, , estensione, SW_SHOWNORMAL) Then:
        'Exit Sub
    Else
        If YouTubeThanks.StartConversion(strFile, , , estensione, SW_SHOWNORMAL) Then:
    End If

Exit Sub
ErrorHandler:
    utcWait.Start = False
        MsgBox "Unaspected Error #" & Err.Number & "." & vbCrLf & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub
Private Sub cmdCopyToClipboard_Click()
    Dim i As Integer: Dim FX As Variant
    On Local Error GoTo ErrorHeadler
    FX = FreeFile
    If FileExists(App.Path + "\Downloads\" + Mid$(URLVideoTitle, 1, Len(URLVideoTitle) - 4) + "\" + URLVideoTitle) Then
        Open Mid$(App.Path + "\Downloads\" + Mid$(URLVideoTitle, 1, Len(URLVideoTitle) - 4) + "\" + URLVideoTitle, 1, Len(App.Path + "\Downloads\" + Mid$(URLVideoTitle, 1, Len(URLVideoTitle) - 4) + "\" + URLVideoTitle) - 4) + ".txt" For Output As #FX
        Print #FX, Tab(5); "Hashed table of: " & URLVideoTitle
        Print #FX, Tab(5); "Generated by YouTube Downloader v2.0.3 on "; Format(Now, "Long Date") & " " & Format(Time, "Long Time")
        Print #FX, Tab(5); "------------------------------------------"
        For i = 0 To cmbHash.ListCount
            If Len(cmbHash.List(i)) > 0 Then
                Print #FX, Tab(5); cmbHash.List(i)
            Else
                Exit For
            End If
        Next
        Close #FX
        MsgBox "Hash/CRC32 and MD5 exported success to a file " & Mid$(URLVideoTitle, 1, Len(URLVideoTitle) - 4) & ".txt!", vbInformation, App.Title
    Else
        MsgBox "The file/folder " & Mid$(URLVideoTitle, 1, Len(URLVideoTitle) - 4) & ", does not exist!", vbExclamation, App.Title
    End If
Exit Sub
ErrorHeadler:
        MsgBox "Error #" & Err.Number & "." & vbCrLf & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub

Private Sub cmdFind_Click()
    Dim ObjLink As HTMLLinkElement
    Dim objMSHTML As New MSHTML.HTMLDocument
    Dim objDocument As MSHTML.HTMLDocument
    Dim strString As String: Dim strString2 As String
    Dim strVideoTitle As String: Dim TempArray(20) As String
    Dim pos1 As Long: Dim pos2 As Long
    Dim i As Integer: Dim FF As Long
    Dim x As Integer: Dim j As Integer
    Dim stringURL As String
    
    On Local Error GoTo ErrorHandler
    
    If Len(txtSearch.Text) = 0 Then
        Exit Sub
    ElseIf Len(txtSearch.Text) < 2 Then
            MsgBox "The Word you find is too short!", vbExclamation, App.Title
            STB.Panels(1).Text = "The Word you find is too short..."
            STB.Panels(1).Picture = imgs(4).Picture
        Exit Sub
    End If
    
    txtSearch.Enabled = False
    
    strString = txtSearch.Text
    strString = Replace(strString, " ", "+")
    strString = Replace(strString, "-", "+")
    strString = Replace(strString, "_", "+")
    
    pPage = 0
    
    ' .... Get the direct URL of Video?
    If InStr(strString, "watch?v=") > 0 Then
        stringURL = strString
    Else
        stringURL = "http://www.youtube.com/results?aq=f&search_query=" & strString & "&hl=en"
    End If
    
    Screen.MousePointer = vbHourglass
    
    URLVideoTitle = Empty
    
    lstTitles.ListItems.Clear: frmMain.Tag = Empty: lstURLs.Clear: lstUrlTitles.Clear
    
    If FileExists(App.Path + "\no-foto.jpg") Then SImage.loadimg App.Path + "\no-foto.jpg"
    
    x = 0: i = 0: j = 0: lblPercentage.Caption = "0%"
    
    STB.Panels(1).Text = "Gettting document via HTTP..."
    STB.Panels(1).Picture = imgs(2).Picture
        
    Set objDocument = objMSHTML.createDocumentFromUrl(stringURL, vbNullString)
    
    While objDocument.ReadyState <> "complete"
        DoEvents
    Wend
    
    STB.Panels(1).Text = "Getting and parsing HTML document..."
    STB.Panels(1).Picture = imgs(3).Picture
    
    strString2 = Inet.OpenURL(stringURL)
    
    While Inet.StillExecuting
        DoEvents
    Wend

    ' .... Italian
    If InStr(strString2, "Il video che hai richiesto non Ã¨ disponibile.") > 0 Then
        STB.Panels(1).Text = "Il video che hai richiesto non è disponibile..."
        STB.Panels(1).Picture = imgs(4).Picture
            MsgBox "Il video che hai richiesto non è disponibile.", vbExclamation, App.Title
                txtSearch.Enabled = True
        Screen.MousePointer = vbDefault
    Exit Sub
        
    ' .... English
    ElseIf InStr(strString2, "The video you have requested is not available.") > 0 Then
        STB.Panels(1).Text = "The video you have requested is not available..."
        STB.Panels(1).Picture = imgs(4).Picture
                MsgBox "The video you have requested is not available.", vbExclamation, App.Title
                    txtSearch.Enabled = True
            Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    DoEvents
    STB.Panels(1).Text = "Please wait... extracting links..."
    STB.Panels(1).Picture = imgs(2).Picture
    
    
    ' =========================================================== Adding in the new version don't write oter code ;)
    
    If InStr(strString, "watch?v=") > 0 Then
            
    ' .... Extract Video Title
    If InStr(strString2, "<meta property=""" & "og:title""" & " content=""") > 0 Then
        pos1 = InStr(pos1 + 1, strString2, "<meta property=""" & "og:title""" & " content=""", vbTextCompare)
        pos2 = InStr(pos1 + 1, strString2, """>", vbTextCompare)
        TempArray(0) = Mid$(strString2, pos1 + 35, pos2 - pos1 - 35)
        txtVideoPicture.Text = TempArray(0)
    Else
        TempArray(0) = "n.a"
    End If
    
    ' --------------------------------------------------------------------------------------------------------------
    If TempArray(0) <> Empty Or Mid$(TempArray(0), 1, 3) <> "ltr" And TempArray(0) <> "n.a" Then
            TempArray(0) = Replace(TempArray(0), "&amp;quot;", sQuote)
            TempArray(0) = Replace(TempArray(0), "&amp;#39;", "'")
            TempArray(0) = Replace(TempArray(0), "&#231;", "ç")
            TempArray(0) = Replace(TempArray(0), "&#232;", "è")
            TempArray(0) = Replace(TempArray(0), "&#233;", "é")
            TempArray(0) = Replace(TempArray(0), "&#224;", "à")
            TempArray(0) = Replace(TempArray(0), "&#242;", "ò")
            TempArray(0) = Replace(TempArray(0), "&#249;", "ù")
            TempArray(0) = Replace(TempArray(0), "Â®", "'s")
            TempArray(0) = Replace(TempArray(0), ":", "-")
            TempArray(0) = Replace(TempArray(0), ";", "-")
            TempArray(0) = Replace(TempArray(0), """", "'")
            TempArray(0) = Replace(TempArray(0), "+", "-")
            TempArray(0) = Replace(TempArray(0), ".", "_")
            TempArray(0) = Replace(TempArray(0), "|", "-")
            TempArray(0) = Replace(TempArray(0), "%", "x100")
            TempArray(0) = Replace(TempArray(0), "$", "(dollar)")
            TempArray(0) = Replace(TempArray(0), "£", "(lit)")
            TempArray(0) = Replace(TempArray(0), "!", "")
            TempArray(0) = Replace(TempArray(0), "&quot;", "'")
            TempArray(0) = Replace(TempArray(0), "/", "-")
            TempArray(0) = Replace(TempArray(0), "\", "-")
                    
            lstUrlTitles.AddItem TempArray(0): lstURLs.AddItem txtSearch.Text
            
            With lstTitles.ListItems.Add(, , TempArray(0))
                .SubItems(1) = txtSearch.Text
            End With
            
            txtSearch.Text = TempArray(0)
            
            If lstURLs.ListCount > 0 Then
                lstUrlTitles.Selected(0) = True
                URLVideoTitle = LCase(lstUrlTitles.List(lstUrlTitles.ListIndex))
        
            ' .... The first Item of the list must be visible ?
                Set itmX = lstTitles.ListItems(1): itmX.EnsureVisible: itmX.Selected = True: lstTitles.SetFocus
            Else
                cmdUseLink.Enabled = False
                cmdConvert.Enabled = False
            End If
            
            lbltotalxPage.Caption = "Total titles found [1] per page [" & lstUrlTitles.ListCount & "]"
            
            STB.Panels(1).Text = "Total titles found [1] per page [" & lstUrlTitles.ListCount & "]"
            STB.Panels(1).Picture = imgs(2).Picture
            
            ' .... Show the Picture
            If FileExists(App.Path + "\prevTmp.jpg") Then
                SImage.loadimg App.Path + "\prevTmp.jpg"
            If FileExists(App.Path + "\prevTmp.jpg") Then
                    Call Kill(App.Path + "\prevTmp.jpg")
                End If
            End If
            
            lblPercentage.Caption = "0%"
    
            If Inet.StillExecuting Then Inet.Cancel
    
            Set ObjLink = Nothing: Set objMSHTML = Nothing: Set objDocument = Nothing
    
            txtSearch.Enabled = True
            lblpages.Caption = "Page 1 of 1"
            cmdNext.Enabled = False: cmdPrev.Enabled = False
            
        End If
            Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    ' =========================================================== Adding in the new version don't write oter code ENDED ... ;)
    
    
    For Each ObjLink In objDocument.links
    If InStr(ObjLink, "watch?v=") Then
        If InStr(ObjLink, "&") = 0 Then

        strVideoTitle = "<a href=" & sQuote & "/watch?v=" & Mid$(ObjLink, 32, Len(ObjLink)) & sQuote & " title=" & sQuote

                ' .... Thanks to Todd for the suggestion :)
                ' .... ********************************************************
                If InStr(ObjLink.innerText, "aAggiunto alla coda") = 0 _
                    And InStr(ObjLink.innerText, "Add to queue") = 0 _
                    And InStr(ObjLink.innerText, "to add this to a playlist") = 0 _
                    And InStr(ObjLink.innerText, "Aggiungi a") = 0 Then
                    TempArray(0) = ObjLink.innerText
                Else
                    TempArray(0) = "n.a"
                End If
                ' .... *******************************************************
                
                ' .... Now convert the special chars
                If TempArray(0) <> Empty And Mid$(TempArray(0), 1, 3) <> "ltr" Then
                    
                    TempArray(0) = Replace(TempArray(0), "&amp;quot;", sQuote)
                    TempArray(0) = Replace(TempArray(0), "&amp;#39;", "'")
                    TempArray(0) = Replace(TempArray(0), "&#231;", "ç")
                    TempArray(0) = Replace(TempArray(0), "&#232;", "è")
                    TempArray(0) = Replace(TempArray(0), "&#233;", "é")
                    TempArray(0) = Replace(TempArray(0), "&#224;", "à")
                    TempArray(0) = Replace(TempArray(0), "&#242;", "ò")
                    TempArray(0) = Replace(TempArray(0), "&#249;", "ù")
                    TempArray(0) = Replace(TempArray(0), "Â®", "'s")
                    TempArray(0) = Replace(TempArray(0), ":", "-")
                    TempArray(0) = Replace(TempArray(0), ";", "-")
                    TempArray(0) = Replace(TempArray(0), """", "'")
                    TempArray(0) = Replace(TempArray(0), "+", "-")
                    TempArray(0) = Replace(TempArray(0), ".", "_")
                    TempArray(0) = Replace(TempArray(0), "|", "-")
                    TempArray(0) = Replace(TempArray(0), "%", "x100")
                    TempArray(0) = Replace(TempArray(0), "$", "(dollar)")
                    TempArray(0) = Replace(TempArray(0), "£", "(lit)")
                    TempArray(0) = Replace(TempArray(0), "!", "")
                    TempArray(0) = Replace(TempArray(0), "&quot;", "'")
                    TempArray(0) = Replace(TempArray(0), "/", "-")
                    TempArray(0) = Replace(TempArray(0), "\", "-")
                    
                    If TempArray(0) <> "n_a" Then
                        lstUrlTitles.AddItem TempArray(0)
                        lstURLs.AddItem ObjLink
                        With lstTitles.ListItems.Add(, , TempArray(0))
                            .SubItems(1) = ObjLink
                        End With
                        STB.Panels(1).Text = "Extracted links: " & ObjLink & "...."
                        STB.Panels(1).Picture = imgs(2).Picture
                    End If
                End If
                TempArray(0) = Empty
            End If
        End If
        DoEvents
        x = x + 1
    Next
    
    ' .... More page in this 'search' video title?
    If InStr(strString2, "Go to page 2") > 0 Then cmdNext.Visible = True
    
    ' .... Recursive search for several pages, I think that seven page is the standard imposed from youtube
    i = 1
    For i = 1 To 50
        If InStr(strString2, "Go to page " & i) > 0 Then
                lblpages.Caption = "Pages 1 of " & i
            lastPage = i
        End If
    Next i
    
    pPage = 1
    
    If lstURLs.ListCount > 0 Then
        lstUrlTitles.Selected(0) = True
        URLVideoTitle = LCase(lstUrlTitles.List(lstUrlTitles.ListIndex))
        
        ' .... Simulate click on the first Item of the list!!
    'If lstTitles.ListItems.Count > 0 Then lstTitles_Click ' Mmmm... this is an hold code but i thik this is not a perfect solution ;)
        
        ' .... The first Item of the list must be visible ?
        Set itmX = lstTitles.ListItems(1): itmX.EnsureVisible: itmX.Selected = True: lstTitles.SetFocus
    Else
        cmdUseLink.Enabled = False
        cmdConvert.Enabled = False
    End If
    
    lbltotalxPage.Caption = "Total titles found [" & x & "] per page [" & lstUrlTitles.ListCount & "]"
    
    STB.Panels(1).Text = "Total titles found [" & x & "] per page [" & lstUrlTitles.ListCount & "]"
    STB.Panels(1).Picture = imgs(2).Picture
    
    ' .... Show the Picture
    If FileExists(App.Path + "\prevTmp.jpg") Then
        SImage.loadimg App.Path + "\prevTmp.jpg"
    If FileExists(App.Path + "\prevTmp.jpg") Then
            Call Kill(App.Path + "\prevTmp.jpg")
        End If
    End If
    
    lblPercentage.Caption = "0%"
    
    If Inet.StillExecuting Then Inet.Cancel
    
    Set ObjLink = Nothing: Set objMSHTML = Nothing: Set objDocument = Nothing
    
    txtSearch.Enabled = True
    
    Screen.MousePointer = vbDefault
Exit Sub
ErrorHandler:
    Screen.MousePointer = vbDefault
        txtSearch.Enabled = True
        If Inet.StillExecuting Then Inet.Cancel
        lblPercentage.Caption = "0%"
        If FileExists(App.Path + "\no-foto.jpg") Then SImage.loadimg App.Path + "\no-foto.jpg"
        MsgBox "Error: " & Err.Number & "." & vbCrLf & Err.Description & vbCrLf & Err.LastDllError & vbCrLf & Err.Source, vbCritical, App.Title
    Err.Clear
End Sub

Private Sub cmdLike_Click()
    On Local Error GoTo ErrorHandler
    Dim IE As Object
    Set IE = CreateObject("internetexplorer.application")
    IE.Visible = True
    IE.Height = 300
    IE.Width = 550
    IE.menubar = False
    IE.Toolbar = False
    IE.StatusBar = False
    IE.resizable = True
    IE.navigate URLFlashVideo
    While IE.busy: DoEvents: Wend
    IE.document.All("watch-like").Click
    Set IE = Nothing
Exit Sub
ErrorHandler:
    Err.Clear
End Sub

Private Sub cmdList_Click()
    Dim FSO As New Scripting.FileSystemObject
    If FSO.FolderExists(App.Path + "\Downloads\") Then
        ShellExecute 0&, vbNullString, App.Path + "\Downloads\", vbNullString, "C:\", SWSHOW.SW_SHOWNORMAL
    Else
        MsgBox "The Folder [Downloads] does not exist!", vbExclamation, App.Title
    End If
    Set FSO = Nothing
End Sub

Private Sub cmdNavigateURL_Click()
    
    Select Case MsgBox("Open the video to current default Browser or Internet Explorer?" & vbCrLf & "1) Default Browser..." _
        & vbCrLf & "2) Internet Explorer Browser..." & vbCrLf & "3) Nothing...", vbYesNoCancel + vbInformation + _
        vbDefaultButton1, App.Title)
    
    Case vbYes
        If YouTubeThanks.OpenToDefaultBrowser(cmbExtension.ListIndex, txtVideoTitle.Text, False, SW_SHOWNORMAL) Then
    Else
        MsgBox "Unaspected error to Open video '" & txtVideoTitle.Text & "' on the default Browser!", vbExclamation, App.Title
    End If
    
    Case vbNo
        If YouTubeThanks.OpenToDefaultBrowser(cmbExtension.ListIndex, txtVideoTitle.Text, True, SW_SHOWNORMAL) Then
    Else
        MsgBox "Unaspected error to Open video '" & txtVideoTitle.Text & "' on the default Browser!", vbExclamation, App.Title
    End If
    Case vbCancel
        Exit Sub
    End Select
End Sub

Private Sub cmdNext_Click()

    If pPage = lastPage Then
        Exit Sub
    Else
        pPage = pPage + 1
    End If
    
    Call NextPrev
    
    On Local Error Resume Next
    
    ' .... Simulate click on the first Item of the list
    If lstTitles.ListItems.Count > 0 Then lstTitles_Click
    
    If lstTitles.ListItems.Count > 0 Then
        Set itmX = lstTitles.ListItems(1)
        itmX.EnsureVisible
        itmX.Selected = True
        lstTitles.SetFocus
    End If
End Sub

Private Sub cmdOterOption_Click()
    tEffect2.Enabled = True
End Sub

Private Sub cmdPrev_Click()
    
    If pPage = 1 Then
        Exit Sub
    Else
        pPage = pPage - 1
    End If
    
    Call NextPrev
    
    On Local Error Resume Next
    
    ' .... Simulate click on the first Item of the list
    If lstTitles.ListItems.Count > 0 Then lstTitles_Click
    
    If lstTitles.ListItems.Count > 0 Then
        Set itmX = lstTitles.ListItems(1)
        itmX.EnsureVisible
        itmX.Selected = True
        lstTitles.SetFocus
    End If
End Sub

Private Sub cmdRegistration_Click()
    If MsgBox(YouTubeThanks.Message & vbCrLf & "You want to activate or send request Key of the program now?", vbYesNo + vbInformation + _
        vbDefaultButton1, "Close Application") = vbYes Then
            YouTubeThanks.EnterRegistationCode
    End If
End Sub

Private Sub cmdUnlike_Click()
    On Local Error GoTo ErrorHandler
    Dim IE As Object
    Set IE = CreateObject("internetexplorer.application")
    IE.Visible = True
    IE.Height = 300
    IE.Width = 550
    IE.menubar = False
    IE.Toolbar = False
    IE.StatusBar = False
    IE.resizable = True
    IE.navigate URLFlashVideo
    While IE.busy: DoEvents: Wend
    IE.document.All("watch-unlike").Click
    Set IE = Nothing
Exit Sub
ErrorHandler:
    Err.Clear
End Sub

Private Sub cmdUseLink_Click()
    Dim FSO As New FileSystemObject

    On Local Error GoTo ErrorHandler
    
    ' .... Video title is necessary
    If txtVideoTitle.Text = "n.a" Or Len(txtVideoTitle.Text) = 0 Then
            MsgBox "The title of video is necessary!", vbExclamation, App.Title
        Exit Sub
    End If
    
    ' .... Construction of the FileName to pass a Class
    URLVideoTitle = txtVideoTitle.Text + "." + LCase(StripLeft(cmbExtension.List(cmbExtension.ListIndex), "-", False)) ' .... OK to use
    
    ' .... Starting download?
    If cmdUseLink.Caption = "&Download Video" Then
        
        ' .... Request video Download YES or NOT
        If MsgBox("Download the Video: " & Mid$(URLVideoTitle, 1, Len(URLVideoTitle) - 4) & "?", _
        vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
        
        ' .... Create the Folder 'Download' and the Subfolder
        If Not FSO.FolderExists(App.Path + "\Downloads\" + Mid$(URLVideoTitle, 1, Len(URLVideoTitle) - 4)) Then
            Call MakeDirectory(App.Path + "\Downloads\" + Mid$(URLVideoTitle, 1, Len(URLVideoTitle) - 4))
        End If
        
        ' .... Video already exists overwrite YES or NOT
        If FileExists(App.Path + "\Downloads\" + Mid$(URLVideoTitle, 1, Len(URLVideoTitle) - 4) + "\" + URLVideoTitle) Then
                If MsgBox("The video file you want to download already exists! " _
            & "Do you want to overwrite it?", vbYesNo + vbExclamation + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
        End If
        
        utcDown.Start = True
        
        ' .... Variable assume TRUE if the Download as started
        Downloading = True
        
        ' .... Disable all buttons
        cmdFind.Enabled = False: cmdAdvanced.Enabled = False: cmdConvert.Enabled = False: _
        cmdClose.Enabled = False: lstTitles.Enabled = False: cmdNext.Enabled = False: cmdPrev.Enabled = False
        cmdNavigateURL.Enabled = False: cmdCopyToClipboard.Enabled = False
        
        ' .... Dispaly info and change icon of the StatusBar
        If Len(Mid$(URLVideoTitle, 1, Len(URLVideoTitle) - 4)) > 20 Then STB.Panels(1).Text = _
        "Starting Download of " & Mid$(Mid$(URLVideoTitle, 1, Len(URLVideoTitle) - 4), 1, 20)
        ' .... Truncate the Video title at 20 chars
        STB.Panels(1).Picture = imgs(3).Picture
        
        ' .... Activate the Time
        Call StartCount
        
        ' .... Change the caption of button, now assume "Abort &Download"
        cmdUseLink.Caption = "Abort &Download"
        
        ' .... Download Started = YES
        YouTubeThanks.StartedDownload = True
        
        ' .... START THE DOWNLOAD OF VIDEO FILE
        YouTubeThanks.Download cmbExtension.ListIndex, App.Path + "\Downloads\" + Mid$(URLVideoTitle, 1, Len(URLVideoTitle) - 4) + "\" _
                                                                    + URLVideoTitle, Mid$(URLVideoTitle, 1, Len(URLVideoTitle) - 4)
        
        
        ' .... Download finish success?
        If FileExists(App.Path + "\Downloads\" + Mid$(URLVideoTitle, 1, Len(URLVideoTitle) - 4) + "\" + URLVideoTitle) Then
        
        ' ....If the video output format <> FLV ... disabled the conversion function
        If ValidFormat(App.Path + "\Downloads\" + Mid$(URLVideoTitle, 1, Len(URLVideoTitle) - 4) + "\" + URLVideoTitle) Then
            cmbFormato.Enabled = True: cmdConvert.Enabled = True
        Else
            cmbFormato.Enabled = False: cmdConvert.Enabled = False
        End If
        
        ' .... Ok restore the Status Quo
        STB.Panels(1).Text = "Download of video file success..."
        STB.Panels(1).Picture = imgs(2).Picture
        
        ' .... and Enabled/Disabled restore the:
        tStart.Enabled = False: lblElapced.Caption = "00:00:00"
        lblSize2.Caption = "0 bytes": lblSaved.Caption = "0 bytes": lblOf.Caption = "0 bytes"
        tDownloaded.Enabled = False: cmdNavigateURL.Enabled = True
        Downloading = False: SoftwareRegistered
        cmdFind.Enabled = True: cmdAdvanced.Enabled = True: cmdClose.Enabled = True: lstTitles.Enabled = True
        
        cmdUseLink.Caption = "&Download Video"
        
        utcDown.Start = False
        
        ' .... Creo la lista dei link's scaricati
        Call CreateLinksList
        
        Screen.MousePointer = vbDefault
            
        Else
            cmdUseLink.Caption = "&Download Video"
            Screen.MousePointer = vbDefault
            tStart.Enabled = False: tStart.Interval = 0: Call Restart(lblElapced)
            cmdConvert.Enabled = False
            tStart.Enabled = False: lblElapced.Caption = "00:00:00"
            tDownloaded.Enabled = False: cmdNavigateURL.Enabled = True
            Downloading = False: SoftwareRegistered
            cmdFind.Enabled = True: cmdAdvanced.Enabled = True: cmdClose.Enabled = True: lstTitles.Enabled = True
            utcDown.Start = False
            Screen.MousePointer = vbDefault
        End If
        
        ' .... Abort download?
        ElseIf cmdUseLink.Caption = "Abort &Download" Then
            
        ' .... Abort the Download of Video?
        If MsgBox("Abort download?" & vbLf & vbLf & "Are you sure?", _
        vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
        
        ' .... Change the caption of button in "&Download Video"
        cmdUseLink.Caption = "&Download Video"
        
        ' .... Send the command Cancel to a Class
        YouTubeThanks.AbortDownload = True
        
        Downloading = False
        YouTubeThanks.StartedDownload = False
        
        tStart.Enabled = False: tStart.Interval = 0: Call Restart(lblElapced)
        cmdConvert.Enabled = False
        tStart.Enabled = False: lblElapced.Caption = "00:00:00"
        tDownloaded.Enabled = False: cmdNavigateURL.Enabled = True
        Downloading = False: SoftwareRegistered
        cmdFind.Enabled = True: cmdAdvanced.Enabled = True: cmdClose.Enabled = True: lstTitles.Enabled = True
        
        cmdUseLink.Caption = "&Download Video"
        utcDown.Start = False
        Screen.MousePointer = vbDefault
        
        End If
Exit Sub
ErrorHandler:
    Screen.MousePointer = vbDefault
    cmdCopyToClipboard.Enabled = False
    utcDown.Start = False
    STB.Panels(1).Text = "Done..."
    STB.Panels(1).Picture = imgs(4).Picture
    Downloading = False: SoftwareRegistered
    cmdNext.Enabled = True: cmdPrev.Enabled = True
    cmdFind.Enabled = True: cmdAdvanced.Enabled = True: cmdClose.Enabled = True: lstTitles.Enabled = True
    tDownloaded.Enabled = False: cmdNavigateURL.Enabled = True
    lblSize2.Caption = "0 bytes": lblSaved.Caption = "0 bytes": lblOf.Caption = "0 bytes"
    tStart.Enabled = False: lblElapced.Caption = "00:00:00"
    MsgBox "Error: " & Err.Number & "." & vbCrLf & Err.Description, vbCritical, App.Title
    Err.Clear
End Sub
Private Function MakeDirectory(szDirectory As String) As Boolean
    Dim strFolder As String
    Dim szRslt As String
    On Error GoTo IllegalFolderName
If Right$(szDirectory, 1) <> "\" Then szDirectory = szDirectory & "\"
strFolder = szDirectory
szRslt = Dir(strFolder, 63)
While szRslt = ""
    DoEvents
    szRslt = Dir(strFolder, 63)
    strFolder = Left$(strFolder, Len(strFolder) - 1)
    If strFolder = "" Then GoTo IllegalFolderName
Wend
If Right$(strFolder, 1) <> "\" Then strFolder = strFolder & "\"
While strFolder <> szDirectory
    strFolder = Left$(szDirectory, Len(strFolder) + 1)
    If Right$(strFolder, 1) = "\" Then MkDir strFolder
Wend
MakeDirectory = True
Exit Function
IllegalFolderName:
        MakeDirectory = False
    Err.Clear
End Function

Private Sub Dloader_DownloadComplete(MaxBytes As Long, SaveFile As String)
    PB.value = 0: STB.Panels(1).Text = "Picture downloaded success..."
    STB.Panels(1).Picture = imgs(3).Picture
End Sub

Private Sub Dloader_DownloadError(SaveFile As String)
    STB.Panels(1).Text = "Error to download the Preview picture..."
    STB.Panels(1).Picture = imgs(4).Picture
End Sub


Private Sub Dloader_DownloadProgress(CurBytes As Long, MaxBytes As Long, SaveFile As String)
    On Local Error Resume Next
    With PB
        .Max = MaxBytes: .value = CurBytes
    End With
    STB.Panels(1).Text = GetSizeBytes(CurBytes, DISP_BYTES_SHORT) & " of " & GetSizeBytes(MaxBytes, DISP_BYTES_SHORT)
    STB.Panels(1).Picture = imgs(2).Picture
End Sub

Private Sub DropArea_Click()
    If lstFileName.ListCount > 0 Then
        Call StartEncodingListFiles(True)
    Else
        MsgBox "The list appears to be empty. Drag the files (*. flv) to be converted over the icon (?)...", vbExclamation, App.Title
    End If
End Sub

Private Sub DropArea_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Local Error Resume Next
    hOldCursor = CopyCursor(GetCursor())
    hCursor = LoadResPicture(102, vbResIcon).handle
    Call SetSystemCursor(hCursor, OCR_NORMAL)
End Sub

Private Sub DropArea_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Local Error Resume Next
    If x >= 0 And Y >= 0 And x < DropArea.Width And Y < DropArea.Height Then
        'SetCursor LoadCursor(0, IDC_HAND)
    Else
        'SetCursor LoadCursor(0, IDC_ARROW)
    End If
End Sub


Private Sub DropArea_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    SetSystemCursor hOldCursor, OCR_NORMAL
End Sub

Private Sub DropArea_OLECompleteDrag(Effect As Long)
    SetSystemCursor hOldCursor, OCR_NORMAL
End Sub


Private Sub DropArea_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
        Dim vFN: Dim strFile As String: Dim i As Integer: Dim smax As Integer
        Dim xi As Integer: Dim ix As Integer
        Dim cCol As FilesCollection
        On Local Error GoTo ErrorHandler
        smax = Data.Files.Count: PB.Max = smax: ix = 0
        For Each vFN In Data.Files
            If (GetAttr(vFN) And vbDirectory) = vbDirectory Then
                    GetSubFiles vFN + "\", vbDirectory, vbArchive, cCol, "*.flv"
                    smax = cCol.Count: PB.Max = smax
                    For ix = 1 To cCol.Count
                            lstFileName.AddItem cCol.Path(ix)
                        DoEvents: PB.value = ix
                    Next ix
                    Call ChkLst(lstFileName)
                    On Local Error Resume Next: lblPercentage.Caption = "%0": PB.value = 0
                    If lstFileName.ListCount > 0 Then
                        If MsgBox("There are " & lstFileName.ListCount & " files in the list ready to be converted. Convert now?" & vbCrLf _
                        & "If you do not want to convert them now, click (?) To start the conversion process", _
                        vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
                        Call StartEncodingListFiles(False)
                    End If
                Exit Sub
            End If
            DoEvents
        Next vFN
        If Data.GetFormat(vbCFFiles) Then
            For Each vFN In Data.Files
                If StrComp((Right$(vFN, 3)), "flv", vbTextCompare) = 0 Then
                    strFile = vFN: xi = xi + 1
                    lstFileName.AddItem strFile
                End If
                DoEvents
                On Local Error Resume Next
                i = i + 1: PB.value = i: lblPercentage.Caption = i & "%"
            Next vFN
        End If
        lblPercentage.Caption = "%0": Call ChkLst(lstFileName)
        If xi = 0 Then
            MsgBox "No a valid file dragged here!", vbExclamation, App.Title
        Else
            If MsgBox("There are " & lstFileName.ListCount & " files in the list ready to be converted. Convert now?" & vbCrLf _
            & "If you do not want to convert them now, click (?) To start the conversion process", _
            vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
            i = 0
            PB.Max = lstFileName.ListCount
            For i = 0 To lstFileName.ListCount
                ' .............................. Start Conversion
                If estensione = "mpg" Then
                    If optVideo.Item(1).value Then
                        If YouTubeThanks.StartConversion(lstFileName.List(i), XVID, DvD43, estensione, SW_SHOWNORMAL) Then:
                    ElseIf optVideo.Item(0).value Then
                        If YouTubeThanks.StartConversion(lstFileName.List(i), XVID, DvD169, estensione, SW_SHOWNORMAL) Then:
                    End If
                ElseIf estensione = "avi" Then
                        If YouTubeThanks.StartConversion(lstFileName.List(i), XVID, , estensione, SW_SHOWNORMAL) Then:
                Else
                        If YouTubeThanks.StartConversion(lstFileName.List(i), , , estensione, SW_SHOWNORMAL) Then:
                End If
                ' .............................. End Conversion
                PB.value = i: lblPercentage.Caption = i & "%"
                DoEvents
            Next i
        End If
        lblPercentage.Caption = "0%": lstFileName.Clear
    On Local Error Resume Next
    PB.value = 0
Exit Sub
ErrorHandler:
    If Err.Number = 380 Then
        MsgBox "No a valid file dragged here!", vbExclamation, App.Title: lstFileName.Clear
    Else
        MsgBox "Error #" & Err.Number & "." & vbCrLf & Err.Description, vbExclamation, App.Title: lstFileName.Clear
    End If
    On Local Error Resume Next
        PB.value = 0: lblPercentage.Caption = "0%"
    Err.Clear
End Sub


Private Sub DropArea_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single, State As Integer)
 On Error Resume Next
    If Data.GetFormat(vbCFFiles) Then
        Effect = vbDropEffectCopy And Effect
    Else
        Effect = vbDropEffectNone
    End If
Exit Sub
End Sub

Private Sub DropArea_OLESetData(Data As DataObject, DataFormat As Integer)
    If DataFormat = vbCFText Then
        'Data.SetData txtSource.SelText, vbCFText
    End If
End Sub

Private Sub DropArea_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    On Error Resume Next
    AllowedEffects = vbDropEffectMove Or vbDropEffectCopy
End Sub

Private Sub Form_Load()
    
    ' .... Verify the Instance
    If App.PrevInstance Then
    If MsgBox("An hoter instance is runnig." & vbLf & vbLf & "You want to close the First instance before running a new instance?", _
        vbYesNo + vbQuestion, App.Title) = vbNo Then End
        If ForceClose = False Then
            MsgBox "Unable to Kill the current Process of {" & App.EXEName & "}!" _
                & vbCrLf & "Ok to Continue in a New Instance of the Program!!!", vbExclamation, App.Title
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    
    ' .... Init ListView
    If InitListView(lstTitles, True, True, True, True, True, True) Then:
    
    With lstTitles
        .View = lvwReport
    End With
    
    ' .... Init my Library
    Set YouTubeThanks = New YouTubeThanks
    
    ' .... Inizializzo la Classe per il Ping del Server
    Set PingServer = New clsPingServer
    
    ' .... Reset INI File Path
    INI.ResetINIFilePath
    
    ' .... Disable the Window CloseBox
    DisableCloseWindowButton Me
    
    ' .... Software already registered ?
    SoftwareRegistered
    
    ' .... Init the Class ToolTip
    Set TT = New CTooltip: TT.Style = TTBalloon: TT.Icon = TTIconInfo
    
    ' .... Get default image
    If FileExists(App.Path + "\no-foto.jpg") Then _
    SImage.loadimg App.Path + "\no-foto.jpg"
    
    ' .... Populate combo conversion format
    Call PopulateCombo
    
    ' .... Display the installed Printers on this SO
    Call GetPrinters
    
    ' .... Retrive installed Font's on this SO
    Call PrintFontName
    
    ' .... Read settings INI
    Call ReadSettings
    
    ' .... Extend the Width of ComboBox
    SendMessage cmbPrinters.hWnd, CB_SETDROPPEDWIDTH, 350, 0
    
    ' .... Disable the Search button if the field search is empty
    If Len(txtSearch.Text) > 0 Then cmdFind.Enabled = True
    
    ' .... Enabled the Autofind of the Last title
    If CheckAutoFind.value = 1 Then _
    tlastSearch.Enabled = True
    
    ' .... Retrive the Prodess ID of Application in case of Crash {=PID}
    PID = WindowToProcessId(Me.hWnd)
    INI.DeleteKey "PROCESSID", "PID": INI.CreateKeyValue "PROCESSID", "PID", PID
    
    ' .... Activate Timer network
    TimerDB.Enabled = True
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' .... STop all sound
    EndPlaySound
    ' .... Confirm EXIT?
    If MsgBox("Are you sure to Close this Application?", vbYesNo + vbInformation + _
        vbDefaultButton1, "Close Application") = vbYes Then
    readyToClose = True
    ' .... Cleanup my Class
    Set YouTubeThanks = Nothing
    ' .... Cleanup Pig Class
    Set PingServer = Nothing
    ' .... Save settings to INI file
    Call SaveSetting
    ' .... Cleanup Class INI
    Set INI = Nothing
    ' .... Form on Nothing
    Set frmMain = Nothing
    Else
        readyToClose = False
    End If
    Cancel = Not readyToClose
End Sub

Private Sub Form_Resize()
    On Local Error GoTo ErrorHandler
    If Me.WindowState = vbMinimized And Downloading = True Then
        tDownloaded.Enabled = True
    ElseIf Me.WindowState = vbNormal And Downloading = True Then
        tDownloaded.Enabled = False
        SoftwareRegistered
    End If
Exit Sub
ErrorHandler:
    Err.Clear
End Sub


Private Sub Form_Terminate()
    If (Forms.Count = 0) Then UnloadApp
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub



Private Function InitListView(sListView As ListView, Optional GRIDLINES As Boolean = True, Optional ONECLICKACTIVATE _
As Boolean = True, Optional FULLROWSELECT As Boolean = True, Optional TRACKSELECT As Boolean = True, _
Optional CHECKBOXES As Boolean = True, Optional SUBITEMIMAGES As Boolean = True) As Boolean
    Dim rStyle As Long: Dim r As Long: On Error GoTo ErrorHeadler
    ' .... Show Greed = True
    If GRIDLINES Then
        rStyle = SendMessageLong(sListView.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
        rStyle = rStyle Or LVS_EX_GRIDLINES
        r = SendMessageLong(sListView.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
    End If
    ' .... One click = True
    If ONECLICKACTIVATE Then
        rStyle = SendMessageLong(sListView.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
        rStyle = rStyle Or LVS_EX_ONECLICKACTIVATE
        r = SendMessageLong(sListView.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
    End If
    ' .... Select all Items = True
    If FULLROWSELECT Then
        rStyle = SendMessageLong(sListView.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
        rStyle = rStyle Or LVS_EX_FULLROWSELECT
        r = SendMessageLong(sListView.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
    End If
    ' .... Track Select = True
    If TRACKSELECT Then
        rStyle = SendMessageLong(sListView.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
        rStyle = rStyle Or LVS_EX_TRACKSELECT
        r = SendMessageLong(sListView.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
    End If
    ' .... CheckBox = True
    If CHECKBOXES Then
        rStyle = SendMessageLong(sListView.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
        rStyle = rStyle Or LVS_EX_CHECKBOXES
        r = SendMessageLong(sListView.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
    End If
    ' .... SubItem Image = True
    If SUBITEMIMAGES Then
        rStyle = SendMessageLong(sListView.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
        rStyle = rStyle Or LVS_EX_SUBITEMIMAGES
        r = SendMessageLong(sListView.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
    End If
    InitListView = True
Exit Function
ErrorHeadler:
    InitListView = False
        MsgBox "Error to Init the sListView!", vbExclamation, App.Title
    Err.Clear
End Function

Private Sub img_Click()
    Call PingS: DelayTime 1, True
End Sub

Private Sub imgPasteFromClipboard_Click()
    Dim ClpFmt: Dim tmpText As String
    On Error Resume Next
    If Clipboard.GetFormat(vbCFText) Then ClpFmt = ClpFmt + 1
    If Clipboard.GetFormat(vbCFBitmap) Then ClpFmt = ClpFmt + 2
    If Clipboard.GetFormat(vbCFDIB) Then ClpFmt = ClpFmt + 4
    If Clipboard.GetFormat(vbCFRTF) Then ClpFmt = ClpFmt + 8
   Select Case ClpFmt
      Case 1 '/// Only TXT
            SetFocusField txtSearch, True
            tmpText = Clipboard.GetText()
            If Mid$(tmpText, 1, 31) <> "http://www.youtube.com/watch?v=" Then Exit Sub
            txtSearch.Text = Clipboard.GetText(): cmdFind_Click
      Case 2, 4, 6 '/// Only PICTURE
      Case 3, 5, 7 '/// TXT and PICTURE
      Case 8, 9 '/// TXT RTF
            SetFocusField txtSearch, True
            tmpText = Clipboard.GetText()
            If Mid$(tmpText, 1, 31) <> "http://www.youtube.com/watch?v=" Then Exit Sub
            txtSearch.Text = Clipboard.GetText(): cmdFind_Click
      Case Else '/// CLIPBOARD EMPTY
   End Select
End Sub

Private Sub imgScanFolders_Click()
    Dim strFolder As String
    
        If imgScanFolders.Tag = "ABORT" Then
            YouTubeThanks.STOPSCAN = True: imgScanFolders.Tag = Empty
        Else
            strFolder = FolderBrowse(Me.hWnd, "Select the folder:")
            If Len(strFolder) = 0 Then Exit Sub
            PB.Max = 100
            imgScanFolders.Tag = "ABORT": YouTubeThanks.STOPSCAN = False
            YouTubeThanks.ScanFolderandCreatehtmlpdfFile strFolder, True, App.Title, App.Title, "*.*", True, 60, SW_SHOWNORMAL
        End If
End Sub

Private Sub imgVideoSize_Click()
    lblSize2.Caption = GetSizeBytes(YouTubeThanks.GetVideoFileSize(cmbExtension.ListIndex), DISP_BYTES_SHORT)
End Sub

Private Sub lstTitles_Click()
    On Local Error Resume Next: lstUrlTitles.Selected(itmSelected) = True
    URLVideoTitle = LCase(lstUrlTitles.List(lstUrlTitles.ListIndex))
End Sub


Private Sub lstTitles_ItemClick(ByVal Item As ComctlLib.ListItem)
    On Local Error Resume Next: itmSelected = Item.Index - 1
End Sub


Private Sub lstTitles_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lItemIndex As Long: Dim lvhti As LVHITTESTINFO
    On Local Error Resume Next
    
    ' .... If the listVieW is empty do Nothing
    If lstTitles.ListItems.Count = 0 Then Exit Sub
    
    ' .... Select the line where the cursor is pointed and extract the text/url
    lvhti.pt.x = x / Screen.TwipsPerPixelX
    lvhti.pt.Y = Y / Screen.TwipsPerPixelY
    lItemIndex = SendMessage(lstTitles.hWnd, LVM_HITTEST, 0, lvhti) + 1 ' .... {+ 1} because the ListViews have a recursion starting at 0
    
    ' .... URL and Title
    URLVideoTitle = lstTitles.ListItems(lItemIndex).Text
    URLFlashVideo = lstTitles.ListItems(lItemIndex).SubItems(1)
    
    sStringText = "Title: " & URLVideoTitle & vbCrLf & "URL: " & URLFlashVideo
    
    ' .... Build the menu on the fly
    mnuContextMenu(0).Caption = "Copy &URL [" & lstTitles.ListItems(lItemIndex).SubItems(1) & "] to Clipboard"
    mnuContextMenu(1).Caption = "Copy this whole list into the Clipboard"
    
    ' .... Now the var URLFlashVideo contains the video URL
    URLFlashVideo = lstTitles.ListItems(lItemIndex).SubItems(1)
    
    ' .... Show the popup menu
    If Button = 2 Then PopupMenu mnuFile
End Sub

Private Sub lstTitles_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lvhti As LVHITTESTINFO
    Dim lItemIndex As Long
    On Local Error GoTo ErrorHandler
    lvhti.pt.x = x / Screen.TwipsPerPixelX
    lvhti.pt.Y = Y / Screen.TwipsPerPixelY
    lItemIndex = SendMessage(lstTitles.hWnd, LVM_HITTEST, 0, lvhti) + 1
    If m_lCurItemIndex <> lItemIndex Then
        m_lCurItemIndex = lItemIndex
        If m_lCurItemIndex = 0 Then
            TT.Destroy
        Else
            TT.Title = "Video Title:"
            TT.TipText = lstTitles.ListItems(m_lCurItemIndex).Text & vbCrLf _
            & "Video URL: " & lstTitles.ListItems(m_lCurItemIndex).SubItems(1)
            TT.Create lstTitles.hWnd
        End If
    End If
Exit Sub
ErrorHandler:
    Err.Clear
End Sub

Private Sub lstUrlTitles_Click()
    Dim strBuffering As String: Dim TempArray(20) As String
    Dim pos1 As Long: Dim pos2 As Long
    Dim FSO As New FileSystemObject: Dim FS As File
    Dim strLines() As String: Dim tempString As String
    Dim i, n, x As Integer
    Dim strLinks(10) As String: Dim strLinkString As String
    Dim FF As Variant: Dim tmpStrings As String
    Dim sTIT As String
    
    FF = FreeFile
    
    On Local Error GoTo ErrorHadler
    Screen.MousePointer = vbHourglass
    i = 0
    
    lstTitles.Enabled = False
    
    cmdUseLink.Enabled = False: cmdConvert.Enabled = False
    cmdNext.Enabled = False: cmdPrev.Enabled = False
    cmdNavigateURL.Enabled = False: cmbHash.Clear
    
    If lstUrlTitles.ListCount = 0 Then Exit Sub
    
    lblvideoTitle.Caption = "n.a": txtVideoTitle.Text = "n.a"
    lblDescription.Text = "n.a"
    
    URLFlashVideo = lstURLs.List(lstUrlTitles.ListIndex)
    
    lstURLs.Selected(lstUrlTitles.ListIndex) = True
    
    URLFlashIDVideo = StripLeft(lstURLs.List(lstUrlTitles.ListIndex), "=", False)
    
    SF.LoadMovie 0, "http://www.youtube.com/v/" & StripLeft(lstURLs.List(lstUrlTitles.ListIndex), "=", False) & "?version=3"
    SF.Play
    
    ' .... Open the Video page URL and get INFO
    
    strBuffering = Inet.OpenURL(URLFlashVideo)
    
    While Inet.StillExecuting
        DoEvents
    Wend

    sTIT = lstUrlTitles.List(lstUrlTitles.ListIndex)
    
    If Len(sTIT) > 40 Then sTIT = Mid$(sTIT, 1, 40) & "#"
    
    lblvideoTitle.Caption = sTIT: txtVideoTitle.Text = lstUrlTitles.List(lstUrlTitles.ListIndex)
    
    URLVideoTitle = LCase(sTIT)

    ' .... Extract Video URL
    If InStr(strBuffering, "<meta property=""" & "og:url""" & " content=""") > 0 Then
        pos1 = InStr(pos1 + 1, strBuffering, "<meta property=""" & "og:url""" & " content=""", vbTextCompare)
        pos2 = InStr(pos1 + 1, strBuffering, """>", vbTextCompare)
        TempArray(4) = Mid$(strBuffering, pos1 + 33, pos2 - pos1 - 33)
        txtVideoPicture.Text = TempArray(4)
    Else
        TempArray(4) = "n.a"
    End If
    
    strTemp = strTemp & "Video URL: " & TempArray(4) & vbCrLf
    
    ' .... Extract Video description
    If InStr(strBuffering, "<meta property=""" & "og:description""" & " content=""") > 0 Then
        pos1 = InStr(pos1 + 1, strBuffering, "<meta property=""" & "og:description""" & " content=""", vbTextCompare)
        pos2 = InStr(pos1 + 1, strBuffering, """>", vbTextCompare)
        TempArray(3) = Mid$(strBuffering, pos1 + 41, pos2 - pos1 - 41)
        TempArray(3) = Replace(TempArray(3), "&#39;", "'")
        TempArray(3) = Replace(TempArray(3), "SÃ¬*", "'")
        TempArray(3) = Replace(TempArray(3), "Ã¨", "è")
        TempArray(3) = Replace(TempArray(3), "&quot;", """")
        TempArray(3) = Replace(TempArray(3), "|", "-")
        TempArray(3) = Replace(TempArray(3), "Ã¹", "ù")
        TempArray(3) = Replace(TempArray(3), "&amp;", "&")
        TempArray(3) = Replace(TempArray(3), "&lt;", "<")
        TempArray(3) = Replace(TempArray(3), "&gt;", ">")
        TempArray(3) = Replace(TempArray(3), "&#64;", "@")
        TempArray(3) = Replace(TempArray(3), "&#96;", "`")
        TempArray(3) = Replace(TempArray(3), "&copy;", "©")
        TempArray(3) = Replace(TempArray(3), "&uml;", "¨")
        TempArray(3) = Replace(TempArray(3), "&pound;", "£")
        TempArray(3) = Replace(TempArray(3), "&reg;", "®")
        TempArray(3) = Replace(TempArray(3), "&macr;", "¯")
        TempArray(3) = Replace(TempArray(3), "&laquo;", "«")
        TempArray(3) = Replace(TempArray(3), "&not;", "¬")
        TempArray(3) = Replace(TempArray(3), "&raquo;", "»")
        TempArray(3) = Replace(TempArray(3), "&deg;", "°")
        TempArray(3) = Replace(TempArray(3), "&sup2;", "²")
        TempArray(3) = Replace(TempArray(3), "&sup3;", "³")
        TempArray(3) = Replace(TempArray(3), "&acute;", "´")
        TempArray(3) = Replace(TempArray(3), "&frac14;", "¼")
        TempArray(3) = Replace(TempArray(3), "&#189;", "½")
        TempArray(3) = Replace(TempArray(3), "&#190;", "¾")
        TempArray(3) = Replace(TempArray(3), "&Agrave;", "À")
        TempArray(3) = Replace(TempArray(3), "&Aacute;", "Á")
        TempArray(3) = Replace(TempArray(3), "&Egrave;", "È")
        TempArray(3) = Replace(TempArray(3), "&Eacute;", "É")
        TempArray(3) = Replace(TempArray(3), "&agrave;", "à")
        TempArray(3) = Replace(TempArray(3), "&aacute", "á")
        TempArray(3) = Replace(TempArray(3), "&auml;", "ä")
        TempArray(3) = Replace(TempArray(3), "&ccedil;", "ç")
        TempArray(3) = Replace(TempArray(3), "&egrave;", "è")
        TempArray(3) = Replace(TempArray(3), "&eacute;", "é")
        TempArray(3) = Replace(TempArray(3), "&igrave;", "ì")
        TempArray(3) = Replace(TempArray(3), "&iacute;", "í")
        TempArray(3) = Replace(TempArray(3), "&ograve;", "ò")
        TempArray(3) = Replace(TempArray(3), "&oacute;", "ó")
        TempArray(3) = Replace(TempArray(3), "&ugrave;", "ù")
        TempArray(3) = Replace(TempArray(3), "&uacute;", "ú")
        TempArray(3) = Replace(TempArray(3), "&yacute;", "ý")
        TempArray(3) = Replace(TempArray(3), "&yuml;", "ÿ")
        TempArray(3) = Replace(TempArray(3), "&euro;", "€")
        
        txtVideoPicture.Text = TempArray(3)
    Else
        TempArray(3) = "n.a"
    End If
    
    strTemp = strTemp & "Video description: " & TempArray(3) & vbCrLf
    
    If TempArray(3) <> "n.a" Then _
    lblDescription.Text = strTemp
    
    ' .... Extract Video Image
    If InStr(strBuffering, "<meta property=""" & "og:image""" & " content=""") > 0 Then
        pos1 = InStr(pos1 + 1, strBuffering, "<meta property=""" & "og:image""" & " content=""", vbTextCompare)
        pos2 = InStr(pos1 + 1, strBuffering, """>", vbTextCompare)
        TempArray(0) = Mid$(strBuffering, pos1 + 35, pos2 - pos1 - 35)
        txtVideoPicture.Text = TempArray(0)
    Else
        TempArray(0) = "n.a"
    End If
    
    strTemp = strTemp & "Video URL image: " & TempArray(0) & vbCrLf
    
    ' .... Download Preview Picture
    If TempArray(0) <> "n.a" Then
    
        Dloader.BeginDownload TempArray(0), App.Path + "\prevTmp.jpg"
        DelayTime 2, True
        If FileExists(App.Path + "\prevTmp.jpg") Then
            ' .... Display the Picture
            SImage.loadimg App.Path + "\prevTmp.jpg"
            ' .... Get the Size
            Set FS = FSO.GetFile(App.Path + "\prevTmp.jpg")
            lblSize.Caption = "Size: " & GetSizeBytes(FS.Size, DISP_BYTES_SHORT)
        If FileExists(App.Path + "\prevTmp.jpg") Then Call Kill(App.Path + "\prevTmp.jpg")
    End If
    End If
    Set FSO = Nothing: Set FS = Nothing

    ' .... Empty list
    cmbExtension.Clear
    
    ' .... Extract Download Links ;)
    YouTubeThanks.TranscodeURLs strBuffering
    
    If cmbExtension.ListCount > 0 Then
        ' .... Enable btn {Download}
        cmdUseLink.Enabled = True
        cmdNavigateURL.Enabled = True
        cmbExtension.ListIndex = 0
    Else
        ' .... Disable btn {Download}
        cmbExtension.Enabled = False
        cmdNavigateURL.Enabled = False
    End If
    
    lstTitles.Enabled = True
    
    cmdNext.Enabled = True: cmdPrev.Enabled = True
    
    PB.value = 0: Call PlaySoundResource(102)
    
    If FileExists(App.Path + "\prevTmp.jpg") Then
            SImage.loadimg App.Path + "\prevTmp.jpg"
        If FileExists(App.Path + "\prevTmp.jpg") Then Call Kill(App.Path + "\prevTmp.jpg")
    End If
    
    strTemp = Empty
    
    Screen.MousePointer = vbDefault
    
Exit Sub
ErrorHadler:
    Call PlaySoundResource(101)
    If Inet.StillExecuting Then Inet.Cancel
    PB.value = 0
    Screen.MousePointer = vbDefault
    cmdNext.Enabled = False: cmdPrev.Enabled = False
    lstTitles.Enabled = True
        MsgBox "Error: " & Err.Number & "." & vbCrLf & Err.Description, vbCritical, App.Title
    Err.Clear
End Sub



Private Function StripLeft(strString As String, strChar As String, Optional sLeftsRight As Boolean = True) As String
  On Local Error Resume Next
  Dim i As Integer
    If sLeftsRight Then
        For i = 1 To Len(strString)
            If Mid$(strString, i, 1) = strChar Then
                    StripLeft = Mid$(strString, 1, i - 1)
                Exit For
            End If
        Next
    Else
        For i = (Len(strString)) To 1 Step -1
        If Mid$(strString, i, 1) = strChar Then
                StripLeft = Mid$(strString, i + 1, Len(strString) - i + 1)
            Exit For
        End If
    Next
End If
End Function

Private Sub NextPrev()
    Dim ObjLink As HTMLLinkElement
    Dim objMSHTML As New MSHTML.HTMLDocument
    Dim objDocument As MSHTML.HTMLDocument
    Dim strString As String: Dim strString2 As String
    Dim strVideoTitle As String: Dim TempArray(20) As String
    Dim pos1 As Long: Dim pos2 As Long
    Dim i As Integer: Dim FF As Long
    Dim x As Integer: Dim j As Integer
    
    On Local Error GoTo ErrorHandler
    
    If Len(txtSearch.Text) = 0 Then
        Exit Sub
    ElseIf Len(txtSearch.Text) < 2 Then
            MsgBox "The Word you find is too short!", vbExclamation, App.Title
        Exit Sub
    End If
    
    lblpages.Caption = "Pages " & pPage & " of " & lastPage
    
    lstTitles.ListItems.Clear: frmMain.Tag = Empty: lstURLs.Clear: lstUrlTitles.Clear
    If FileExists(App.Path + "\no-foto.jpg") Then SImage.loadimg App.Path + "\no-foto.jpg"
    x = 0: i = 0: j = 0: lblPercentage.Caption = "0%"
    
    Screen.MousePointer = vbHourglass
    
    strString = txtSearch.Text
    strString = Replace(strString, " ", "+")
    strString = Replace(strString, "-", "+")
    strString = Replace(strString, "_", "+")
    
    STB.Panels(1).Text = "Gettting document via HTTP..."
    STB.Panels(1).Picture = imgs(3).Picture
    
    Set objDocument = objMSHTML.createDocumentFromUrl("http://www.youtube.com/results?search_query=" & strString & "&suggested_categories=1%2C24&page=" & pPage, vbNullString)
    
    While objDocument.ReadyState <> "complete"
        DoEvents
    Wend
    
    STB.Panels(1).Text = "Getting and parsing HTML document..."
    STB.Panels(1).Picture = imgs(3).Picture
    
    strString2 = Inet.OpenURL("http://www.youtube.com/results?search_query=" & strString & "&suggested_categories=1%2C24&page=" & pPage)
    
    While Inet.StillExecuting
        DoEvents
    Wend

    DoEvents
    
    STB.Panels(1).Text = "Please wait... extracting links..."
    STB.Panels(1).Picture = imgs(2).Picture
    
    For Each ObjLink In objDocument.links
        
        If InStr(ObjLink, "watch?v=") Then
            If InStr(ObjLink, "&") = 0 Then
            
                strVideoTitle = "<a href=" & sQuote & "/watch?v=" & Mid$(ObjLink, 32, Len(ObjLink)) & sQuote & " title=" & sQuote
                
                ' .... Thanks to Todd for the suggestion :)
                '********************************************************
                If InStr(ObjLink.innerText, "aAggiunto alla coda") = 0 _
                    And InStr(ObjLink.innerText, "Add to queue") = 0 _
                    And InStr(ObjLink.innerText, "to add this to a playlist") = 0 _
                    And InStr(ObjLink.innerText, "Aggiungi a") = 0 Then
                TempArray(0) = ObjLink.innerText
                Else
                    TempArray(0) = "n.a"
                End If
                ' .... ******************************************************
                
                If TempArray(0) <> Empty And Mid$(TempArray(0), 1, 3) <> "ltr" Then
                    
                    TempArray(0) = Replace(TempArray(0), "&amp;quot;", sQuote)
                    TempArray(0) = Replace(TempArray(0), "&amp;#39;", "'")
                    TempArray(0) = Replace(TempArray(0), "&#231;", "ç")
                    TempArray(0) = Replace(TempArray(0), "&#232;", "è")
                    TempArray(0) = Replace(TempArray(0), "&#233;", "é")
                    TempArray(0) = Replace(TempArray(0), "&#224;", "à")
                    TempArray(0) = Replace(TempArray(0), "&#242;", "ò")
                    TempArray(0) = Replace(TempArray(0), "&#249;", "ù")
                    TempArray(0) = Replace(TempArray(0), ":", "-")
                    TempArray(0) = Replace(TempArray(0), ";", "-")
                    TempArray(0) = Replace(TempArray(0), """", "'")
                    TempArray(0) = Replace(TempArray(0), "+", "-")
                    TempArray(0) = Replace(TempArray(0), ".", "_")
                    TempArray(0) = Replace(TempArray(0), "|", "-")
                    TempArray(0) = Replace(TempArray(0), "%", "x100")
                    TempArray(0) = Replace(TempArray(0), "$", "(dollar)")
                    TempArray(0) = Replace(TempArray(0), "£", "(lit)")
                    TempArray(0) = Replace(TempArray(0), "!", "(esclam)")
                    TempArray(0) = Replace(TempArray(0), "&quot;", "'")
                    
                    If TempArray(0) <> "n_a" Then
                        lstUrlTitles.AddItem TempArray(0)
                        lstURLs.AddItem ObjLink
                        With lstTitles.ListItems.Add(, , TempArray(0))
                            .SubItems(1) = ObjLink
                        End With
                        STB.Panels(1).Text = "Extracted links: " & ObjLink & "...."
                        STB.Panels(1).Picture = imgs(3).Picture
                    End If
                End If
                TempArray(0) = Empty
            End If
        End If
        DoEvents
    Next
    
    STB.Panels(1).Text = "Done...": STB.Panels(1).Picture = imgs(3).Picture
    STB.Panels(1).Picture = imgs(2).Picture
    
    Screen.MousePointer = vbDefault
    
Exit Sub
ErrorHandler:
    Screen.MousePointer = vbDefault
        MsgBox "Error: " & Err.Number & "." & vbCrLf & Err.Description, vbCritical, App.Title
    Err.Clear
End Sub

Private Sub PopulateCombo()
    On Local Error GoTo ErrorHandler
    With cmbFormato
        .AddItem "-- Video Encoding --"
        .AddItem "AVI DivX Compatible (avi)"
        .AddItem "AVI XviD Compatible (avi)"
        .AddItem "iPOD Video (320x240) (mpeg4)"
        .AddItem "iPOD Video2 (640x480) (mpeg4)"
        .AddItem "MPG DvD Video 4:3 (mpg)"
        .AddItem "MPG DvD Video 16:9 (mpg)"
        .AddItem "M2V Video (m2v)"
        .AddItem "3G2 Video (3g2)"
        .AddItem "3GP for Mobile (3gp)"
        .AddItem "MP4 for Mobile (mp4)"
        .AddItem "MKV Matroska Video (mkv)"
        .AddItem "MOV QuickTime video (mov)"
        .AddItem "-- Audio Encoding --"
        .AddItem "WMV Windows Media Compatible (wmv)"
        .AddItem "WAVE Sound (wav)"
        .AddItem "MP3 Sound (mp3)"
        .AddItem "MP2 Sound (mp2)"
        .AddItem "AC3 Sound (ac3)"
        .AddItem "FLAC Sound (flac)"
        .AddItem "AAC Sound (aac)"
        .AddItem "M4A Sound (m4a)"
        .AddItem "WMA Sound (wma)"
        .AddItem "OGG Sound (ogg)"
        .AddItem "YUV RAW Sound (yuv)"
        .AddItem "VOB Sound (vob)"
    End With
    cmbFormato.ListIndex = 19
Exit Sub
ErrorHandler:
        cmbFormato.Enabled = False
    Err.Clear
End Sub

Private Sub iSenseChange(tBox As TextBox)
    Dim iStart As Integer: Dim iSense As String
    On Local Error GoTo ErrorHandler
    iStart = tBox.SelStart
    iSense = IntelliSense(tBox, False)
    If Len(iSense) > 0 And Not WasDelete Then
        tBox.Text = iSense: tBox.SelStart = iStart: tBox.SelLength = Len(tBox.Text) - iStart
    End If
Exit Sub
ErrorHandler:
    Err.Clear
End Sub

Private Sub iSenseKeyPress(tBox As TextBox, KeyAscii As Integer)
    On Local Error GoTo ErrorHandler
    If KeyAscii = 13 And tBox.Text <> "" Then
        IntelliSense tBox, True
    ElseIf KeyAscii = 8 Then
        WasDelete = True
    Else
        WasDelete = False
    End If
Exit Sub
ErrorHandler:
    Err.Clear
End Sub

Private Function IntelliSense(tBox As TextBox, AddRecord As Boolean) As String
    Dim iChannel As Integer, iActive As Integer, iLength As Integer, i As Integer
    Dim iFile As String: Dim iSense As iSense: Dim Done As Boolean
    On Local Error GoTo ErrorHandler
    iFile = App.Path & "\" & App.EXEName & ".dat"
    iLength = Len(iSense): iChannel = FreeFile
    Open iFile For Random As iChannel Len = iLength
    Close iChannel
    iActive = FileLen(iFile) / iLength: iChannel = FreeFile
    Open iFile For Random As iChannel Len = iLength
        If AddRecord Then
            iSense.sOut = tBox.Text
            Put iChannel, iActive + 1, iSense
        Else
            Do While Not EOF(iChannel) And Done = False
                i = i + 1
                Get iChannel, i, iSense
                If tBox.Text = Mid(RTrim(iSense.sOut), 1, Len(tBox.Text)) Then
                    IntelliSense = RTrim(iSense.sOut)
                End If
            Loop
        End If
    Close iChannel
Exit Function
ErrorHandler:
    Err.Clear
End Function

Private Sub mnuContextMenu_Click(Index As Integer)
    On Local Error GoTo ErrorHandler
    Select Case Index
        Case 0 ' Copy to Clipboard
            Clipboard.Clear
            Clipboard.SetText sStringText
            MsgBox "Video URL [" & URLFlashVideo & "], was successfully copied to the Clipboard.", vbInformation, App.Title
    End Select
Exit Sub
ErrorHandler:
        MsgBox "Error #" & Err.Number & "." & vbCrLf & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub

Private Sub mnuContextMenuOption_Click(Index As Integer)
    On Local Error GoTo ErrorHandler
    Select Case Index
        Case 0 ' Print List
            Call PrintList
        Case Else
        
    End Select
Exit Sub
ErrorHandler:
        MsgBox "Error #" & Err.Number & "." & vbCrLf & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub

Private Sub PingServer_PingReturn(ipAddress As String, successStatus As String, roundTripMilliseconds As Long)
    On Local Error GoTo ErrorHandler
    'lstResponse.AddItem Time & "-" & ipAddress & "  -> " & successStatus
    'lstResponse.AddItem "Round Trip" & "  -> " & roundTripMilliseconds
    lblPing.Caption = "Wait... " & roundTripMilliseconds
    strPing = roundTripMilliseconds
Exit Sub
ErrorHandler:
    Err.Clear
End Sub

Private Sub tDownloaded_Timer()
    Me.Caption = lblSaved.Caption & " / " & lblPercentage.Caption
End Sub

Private Sub tEffect_Timer()
    On Local Error Resume Next
    If cmdAdvanced.Caption = "&Option" Then
        cmdFind.Enabled = False: lstTitles.Enabled = False: txtSearch.Enabled = False
        utcDown.Start = True: cmdAdvanced.Enabled = False
        picOption.Visible = True
        picOption.Width = picOption.Width + 100
        If picOption.Width >= 12600 Then
            tEffect.Enabled = False: cmdAdvanced.Caption = "&Hide"
            cmdAdvanced.ToolTipText = "Hide Advanced Option..."
            cmdAdvanced.Enabled = True: utcDown.Start = False
        End If
    ElseIf cmdAdvanced.Caption = "&Hide" Then
        utcDown.Start = True
        cmdAdvanced.Enabled = False
        picOption.Width = picOption.Width - 100
        If picOption.Width <= 90 Then
            tEffect.Enabled = False
            picOption.Visible = False: cmdAdvanced.Caption = "&Option"
            cmdAdvanced.ToolTipText = "Advanced Option Program..."
            cmdAdvanced.Enabled = True: utcDown.Start = False
            cmdFind.Enabled = True: lstTitles.Enabled = True: txtSearch.Enabled = True
        End If
    End If
End Sub

Private Sub tEffect2_Timer()
    On Local Error Resume Next
    If cmdOterOption.Caption = ">>" Then
        cmdAdvanced.Enabled = False: cmdOterOption.Enabled = False
        picOption2.Visible = True: utcDown.Start = True
        picOption2.Width = picOption2.Width + 100
        If picOption2.Width >= 11715 Then
            tEffect2.Enabled = False: cmdOterOption.Caption = "<<"
            cmdOterOption.ToolTipText = "Hide oter Option..."
            cmdOterOption.Enabled = True: utcDown.Start = False
        End If
    ElseIf cmdOterOption.Caption = "<<" Then
        utcDown.Start = True
        cmdOterOption.Enabled = False
        picOption2.Width = picOption2.Width - 100
        If picOption2.Width <= 90 Then
            tEffect2.Enabled = False
            picOption2.Visible = False: cmdOterOption.Caption = ">>"
            cmdOterOption.ToolTipText = "Show oter Option..."
            cmdAdvanced.Enabled = True: utcDown.Start = False
            cmdOterOption.Enabled = True
        End If
    End If
End Sub


Private Sub TimerDB_Timer()
    Call PingS: TimerDB.Enabled = False
End Sub

Private Sub tlastSearch_Timer()
    On Local Error Resume Next
    ' .... Find last video
    If Len(txtSearch.Text) > 2 Then cmdFind_Click: tlastSearch.Enabled = False
End Sub

Private Sub tStart_Timer()
    Call StopWatch(lblElapced): DoEvents
End Sub

Private Sub txtSearch_Change()
    On Local Error Resume Next
    If ChckIntelliSense.value = 1 Then iSenseChange txtSearch
    
    If Len(txtSearch.Text) > 50 Then
        txtSearch.Text = Mid$(txtSearch.Text, 1, 50)
        txtSearch.SelStart = Len(txtSearch.Text)
        txtSearch.SetFocus
        Exit Sub
    End If
    cmdFind.Enabled = Not Len(txtSearch.Text) < 3
End Sub


Private Sub txtSearch_GotFocus()
    SetFocusField txtSearch, True, 1
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    On Local Error GoTo ErrorHandler
    
    If ChckIntelliSense.value = 1 Then
        iSenseKeyPress txtSearch, KeyAscii
    End If
    
    If KeyAscii = 13 And Len(txtSearch.Text) > 2 Then
        ' .... Lst search
        INI.DeleteKey "SETTING", "LAST_SEARCH"
        If Len(txtSearch.Text) > 0 Then _
        INI.CreateKeyValue "SETTING", "LAST_SEARCH", txtSearch.Text
        
        cmdFind_Click
    End If
    
Exit Sub
ErrorHandler:
    Err.Clear
End Sub


Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    On Local Error GoTo ErrorHandler
    If Len(txtSearch.Text) > 50 Then
        txtSearch.Text = Mid$(txtSearch.Text, 1, 50)
        txtSearch.SelStart = Len(txtSearch.Text)
        txtSearch.SetFocus
        Exit Sub
    End If
    cmdFind.Enabled = Not Len(txtSearch.Text) < 3
ErrorHandler:
    Err.Clear
End Sub

Private Function NormalizeSpaces(ByVal strString As String) As String
    On Local Error Resume Next
    ' .... Removes all leading and trailing spaces and replaces any occurrence of 2 or more spaces with a single space
    strString = Trim(strString)
    Do While InStr(strString, String(2, " ")) > 0
        strString = Replace(strString, String(2, " "), " ")
    Loop
    NormalizeSpaces = strString
End Function



Private Sub StopWatch(WhatLabel As Label)
Dim addSeconds As Long
On Error GoTo ErrorHandler
If Seconds = 60 Then
    AddMinutes = True
    addSeconds = 0
    Seconds = 0
Else
    Seconds = Seconds + 1
End If
If AddMinutes = True Then
    If Minutes = 60 Then
    AddHours = True
    Minutes = 0
Else
        Minutes = Minutes + 1
    End If
    AddMinutes = False
End If
If AddHours = True Then
    If Hours = 24 Then
        AddDays = True
        Hours = 0
        Days = Days + 1
    Else
        Hours = Hours + 1
    End If
    AddHours = False
End If

If AddDays = True Then
    If Days = 999 Then
    Days = 0
Else
        Days = Days + 1
    End If
    AddDays = False
End If
WhatLabel.Caption = Format(Hours, "00") & ":" & Format(Minutes, "00") & ":" & Format(Seconds, "00")
Exit Sub
ErrorHandler:
    Err.Clear: Call Restart(lblElapced)
    tStart.Enabled = False: tStart.Interval = 0
End Sub

Private Sub StartCount()
    Call Restart(lblElapced): tStart.Interval = 1000: tStart.Enabled = True
    DoEvents
End Sub

Private Sub Restart(WhatLabel As Label)
    WhatLabel.Caption = "00:00:00": Seconds = 0: Minutes = 0: Hours = 0: Days = 0
End Sub


Private Sub txtVideoTitle_GotFocus()
    SetFocusField txtVideoTitle, True, 0
End Sub


Private Sub YouTubeThanks_Aborted(CancelDownload As Boolean)
    Screen.MousePointer = vbDefault
    Call PlaySoundResource(101)
    STB.Panels(1).Text = "Video downloaded aborted by User..."
    STB.Panels(1).Picture = imgs(4).Picture
    tStart.Enabled = False: lblElapced.Caption = "00:00:00"
    lblSize2.Caption = "0 bytes": lblSaved.Caption = "0 bytes": lblOf.Caption = "0 bytes"
    
    tDownloaded.Enabled = False
    Downloading = False: SoftwareRegistered
    cmdFind.Enabled = True: cmdAdvanced.Enabled = True: cmdClose.Enabled = True: lstTitles.Enabled = True
    cmdNext.Enabled = True: cmdPrev.Enabled = True
    cmdCopyToClipboard.Enabled = False
    YouTubeThanks.StartedDownload = False
    utcDown.Start = False
    Downloading = False
    lblPercentage.Caption = "0%"
    
    On Local Error Resume Next
    PB.value = 0
End Sub

Private Sub YouTubeThanks_DirScan(scanDir As String, fileName As String, totFiles As Integer, totFolders As Integer, workProgress As Integer)
    On Local Error Resume Next
    PB.value = workProgress
End Sub


Private Sub YouTubeThanks_DownloadCompleted(FinishDownload As Boolean)
    If FinishDownload = True Then
        Call PlaySoundResource(102)
        utcDown.Start = False
        Screen.MousePointer = vbDefault
        STB.Panels(1).Text = "Download finished success..."
        STB.Panels(1).Picture = imgs(2).Picture
        
        tStart.Enabled = False: lblElapced.Caption = "00:00:00"
        lblSize2.Caption = "0 bytes": lblSaved.Caption = "0 bytes": lblOf.Caption = "0 bytes"
        
        tDownloaded.Enabled = False
        Downloading = False: SoftwareRegistered
        cmdFind.Enabled = True: cmdAdvanced.Enabled = True: cmdClose.Enabled = True: lstTitles.Enabled = True
        cmdNext.Enabled = True: cmdPrev.Enabled = True
        
        YouTubeThanks.StartedDownload = False
        Downloading = False
        lblPercentage.Caption = "0%"
        
        If Me.WindowState = vbMinimized Then Me.WindowState = vbNormal
        
        If CheckHash.value = 1 Then
            If FileExists(YouTubeThanks.fileName) Then
                cmbHash.Clear
                cmbHash.AddItem "CRC32: " & YouTubeThanks.GetCRC32(YouTubeThanks.fileName)
                cmbHash.AddItem "MD4: " & YouTubeThanks.HashFile(YouTubeThanks.fileName, MD4)
                cmbHash.AddItem "MD5: " & YouTubeThanks.HashFile(YouTubeThanks.fileName, MD5)
                cmbHash.AddItem "SHA1: " & YouTubeThanks.HashFile(YouTubeThanks.fileName, SHA1)
            End If
            If cmbHash.ListCount > 0 Then
                cmbHash.ListIndex = 3
                cmdCopyToClipboard.Enabled = True
            Else
                cmdCopyToClipboard.Enabled = False
            End If
        End If
        
        On Local Error Resume Next
        PB.value = 0
    End If
End Sub

Private Sub YouTubeThanks_DownloadProgress(CurBytes As Long, RestBytes As Long)
    lblSaved.Caption = GetSizeBytes(CurBytes, DISP_BYTES_SHORT)
    lblOf.Caption = GetSizeBytes(RestBytes, DISP_BYTES_SHORT)
    STB.Panels(1).Text = "Download... " & GetSizeBytes(CurBytes, DISP_BYTES_SHORT) & " of " & GetSizeBytes(RestBytes, DISP_BYTES_SHORT)
    STB.Panels(1).Picture = imgs(3).Picture
End Sub

Private Sub YouTubeThanks_DownloadStarted(Started As Boolean)
    If Started Then
        tStart.Enabled = True: Downloading = True
    Else
        tStart.Enabled = False: Downloading = False
    End If
End Sub

Private Sub YouTubeThanks_ErrorDownload(DownloadError As Boolean)
    If DownloadError = True Then
        Call PlaySoundResource(101)
        Screen.MousePointer = vbDefault
        STB.Panels(1).Text = "Unaspected error to download this Video..."
        STB.Panels(1).Picture = imgs(4).Picture
        tStart.Enabled = False: lblElapced.Caption = "00:00:00"
        lblSize2.Caption = "0 bytes": lblSaved.Caption = "0 bytes": lblOf.Caption = "0 bytes"
        
        tDownloaded.Enabled = False
        Downloading = False
        cmdFind.Enabled = True: cmdAdvanced.Enabled = True: cmdClose.Enabled = True: lstTitles.Enabled = True
        cmdNext.Enabled = True: cmdPrev.Enabled = True
        
        YouTubeThanks.StartedDownload = False
        Downloading = False
        lblPercentage.Caption = "0%"
        utcDown.Start = False
        
        On Local Error Resume Next
        PB.value = 0
    End If
End Sub

Private Sub YouTubeThanks_ProgressBar(Progress As Long, MaxValue As Long)
    lblPercentage.Caption = Progress & " % "
    PB.Max = MaxValue: PB.value = Progress
End Sub

Private Sub YouTubeThanks_ProgressHash(Progress As Long, MaxProgress As Long)
    On Local Error GoTo ErrorHandler
    PB.Max = MaxProgress: PB.value = Progress
Exit Sub
ErrorHandler:
        PB.value = 0
    Err.Clear
End Sub

Private Sub YouTubeThanks_TotalFileSize(CurFileSize As Long)
    lblSize2.Caption = GetSizeBytes(CurFileSize, DISP_BYTES_SHORT)
End Sub


Private Sub YouTubeThanks_VideoConversion(strMessage As String, strSuccess As Boolean)
    If strSuccess Then
        STB.Panels(1).Text = strMessage
        STB.Panels(1).Picture = imgs(2).Picture
    Else
        STB.Panels(1).Picture = imgs(4).Picture
        STB.Panels(1).Text = "Error to convert the Video..."
    End If
    
    Select Case strMessage
        Case "Shell Application started... Wait for ended..."
            utcWait.Start = True
        Case "Shelled and Wait failed or abandoned..."
            utcWait.Start = False
        Case "The Application shell has ended..."
            utcWait.Start = False
    End Select
    
End Sub



Private Sub DelayTime(ByVal Second As Long, Optional ByVal Refresh As Boolean = True)
    On Error Resume Next
    Dim Start As Date
    Start = Now: Do
    If Refresh Then DoEvents
    Loop Until DateDiff("s", Start, Now) >= Second
End Sub

Private Sub DeleteFileToRecycleBin(fileName As String)
    On Local Error Resume Next
    Dim fop As SHFILEOPTSTRUCT
    With fop
        .wFunc = FO_DELETE: .pFrom = fileName: .fFlags = FOF_ALLOWUNDO
    End With
    SHFileOperation fop
End Sub

Private Function IsLoaded(strFormName As String) As Boolean
    Dim i As Integer
    For i = 0 To Forms.Count - 1
        If (Forms(i).Name = strFormName) Then
                IsLoaded = True
            Exit For
        End If
    Next
End Function

Private Function CheckAllItems(lstListViewName As ListView) As Boolean
    
    On Error GoTo err_EnhLitView_CheckAllItems
    
    CheckAllItems = True
    
    Dim LV          As LVITEM
    Dim lvCount     As Long
    Dim lvIndex     As Long
    Dim lvState     As Long
    Dim r           As Long
    
    lvState = IIf(True, &H2000, &H1000)
    lvCount = lstListViewName.ListItems.Count - 1
    Do
        With LV
            .mask = LVIF_STATE
            .State = lvState
            .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessageAny(lstListViewName.hWnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
    
    Exit Function
    
err_EnhLitView_CheckAllItems:
    CheckAllItems = False

    Exit Function
End Function

Private Function UnCheckAllItems(lstListViewName As ListView) As Boolean
                
    On Error GoTo err_UnCheckAllItems
    
    UnCheckAllItems = True
    
    Dim LV          As LVITEM
    Dim lvCount     As Long
    Dim lvIndex     As Long
    Dim lvState     As Long
    Dim r           As Long
    
    lvState = IIf(True, &H2000, &H1000)
    lvCount = lstListViewName.ListItems.Count - 1
    Do
        With LV
            .mask = LVIF_STATE
            .State = lvState
            .stateMask = LVIS_STATEIMAGEMASK
        End With
        r = SendMessageAny(lstListViewName.hWnd, LVM_SETITEMSTATE, lvIndex, LV)
        lvIndex = lvIndex + 1
    Loop Until lvIndex > lvCount
    
    Exit Function

err_UnCheckAllItems:
    UnCheckAllItems = False
    Exit Function
End Function

Private Function RemCheckBoxes(lstListViewName As ListView) As Boolean
    
    On Error GoTo err_CheckBoxes
    
    RemCheckBoxes = True

    Dim rStyle  As Long
    Dim r       As Long

    rStyle = SendMessageLong(lstListViewName.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    
    rStyle = rStyle Xor LVS_EX_CHECKBOXES
    
    SendMessageLong lstListViewName.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle
    Exit Function
err_CheckBoxes:
    RemCheckBoxes = False
End Function

Private Sub SaveSetting()
    On Local Error Resume Next
    
    ' .... File extension
    INI.DeleteKey "SETTING", "EXTENSION"
    INI.CreateKeyValue "SETTING", "EXTENSION", cmbExtension.ListIndex
    
    ' .... Intellisense
    INI.DeleteKey "SETTING", "INTELLISENSE"
    INI.CreateKeyValue "SETTING", "INTELLISENSE", ChckIntelliSense.value
    
    ' .... Lst search
    INI.DeleteKey "SETTING", "LAST_SEARCH"
    If Len(txtSearch.Text) > 0 Then _
    INI.CreateKeyValue "SETTING", "LAST_SEARCH", txtSearch.Text
    
    ' .... Format Video Conversion
    INI.DeleteKey "SETTING", "CONVERSION"
    INI.CreateKeyValue "SETTING", "CONVERSION", cmbFormato.ListIndex
    
    ' .... Autofind
    INI.DeleteKey "SETTING", "AUTOFIND"
    INI.CreateKeyValue "SETTING", "AUTOFIND", CheckAutoFind.value
    
    ' .... Get Hashed of video file?
    INI.DeleteKey "SETTING", "GETASH"
    INI.CreateKeyValue "SETTING", "GETASH", CheckHash.value
    
    ' .... Get video format
    INI.DeleteKey "SETTING", "16943"
    If optVideo.Item(0).value = True Then
        INI.CreateKeyValue "SETTING", "16943", 0
    Else
        INI.CreateKeyValue "SETTING", "16943", 1
    End If
    
    ' .... Application StartUp
    INI.DeleteKey "SETTING", "STARTUP"
    INI.CreateKeyValue "SETTING", "STARTUP", CheckStartUp.value
    
    ' .... Sound Downloads
    INI.DeleteKey "SETTING", "PLAYSOUNDOWNLOADS"
    INI.CreateKeyValue "SETTING", "PLAYSOUNDOWNLOADS", CheckSnd.value
    
    ' .... Printers Left Margin
    INI.DeleteKey "SETTING", "PRINTERSLEFTMARGIN"
    INI.CreateKeyValue "SETTING", "PRINTERSLEFTMARGIN", cmbLeftMargin.ListIndex
    
    ' .... Printers Top Margin
    INI.DeleteKey "SETTING", "PRINTERSTOPMARGIN"
    INI.CreateKeyValue "SETTING", "PRINTERSTOPMARGIN", cmbTopMargin.ListIndex
    
    ' .... Default Printer
    INI.DeleteKey "SETTING", "DEFAULTPRINTRER"
    INI.CreateKeyValue "SETTING", "DEFAULTPRINTRER", cmbPrinters.ListIndex
    
    ' .... Default Font printer
    INI.DeleteKey "SETTING", "DEFAULTFONTPRINTRER"
    INI.CreateKeyValue "SETTING", "DEFAULTFONTPRINTRER", cmbFonts.ListIndex
    
    ' .... Display Printer setup Dialog
    INI.DeleteKey "SETTING", "PRINTRERSETUPDIALOG"
    INI.CreateKeyValue "SETTING", "PRINTRERSETUPDIALOG", CheckDisplayPrinter.value
    
    ' .... Font size
    INI.DeleteKey "SETTING", "DEFAULTFONTSIZE"
    INI.CreateKeyValue "SETTING", "DEFAULTFONTSIZE", cmbFontSize.ListIndex
    
    ' .... Default download video Path
    INI.DeleteKey "SETTING", "DEFAULTVIDEOPATH"
    If Len(txtDefaultPath.Text) > 0 Then
        INI.CreateKeyValue "SETTING", "DEFAULTVIDEOPATH", txtDefaultPath.Text
    Else
        txtDefaultPath.Text = App.Path + "\Downloads\"
        INI.CreateKeyValue "SETTING", "DEFAULTVIDEOPATH", txtDefaultPath.Text
    End If
    
End Sub

Private Sub ReadSettings()
    
    On Local Error Resume Next
    
        ' .... File extension
        If INI.GetKeyValue("SETTING", "EXTENSION") <> Empty Then
            cmbExtension.ListIndex = INI.GetKeyValue("SETTING", "EXTENSION")
        Else
            cmbExtension.ListIndex = 2
        End If
        
        ' .... Intellisense
        If INI.GetKeyValue("SETTING", "INTELLISENSE") <> Empty Then
            ChckIntelliSense.value = INI.GetKeyValue("SETTING", "INTELLISENSE")
        Else
            ChckIntelliSense.value = 0
        End If
    
        ' .... Last search
        If INI.GetKeyValue("SETTING", "LAST_SEARCH") <> Empty Then
            txtSearch.Text = INI.GetKeyValue("SETTING", "LAST_SEARCH")
        End If
    
        ' .... Format Video Conversion
        If INI.GetKeyValue("SETTING", "CONVERSION") <> Empty Then
            cmbFormato.ListIndex = INI.GetKeyValue("SETTING", "CONVERSION")
        Else
            cmbFormato.ListIndex = 1
        End If
        
        ' .... Autofind
        If INI.GetKeyValue("SETTING", "AUTOFIND") <> Empty Then
            CheckAutoFind.value = INI.GetKeyValue("SETTING", "AUTOFIND")
        Else
            CheckAutoFind.value = 0
        End If
        
        ' .... Hashed file?
        If INI.GetKeyValue("SETTING", "GETASH") <> Empty Then
            CheckHash.value = INI.GetKeyValue("SETTING", "GETASH")
        Else
            CheckHash.value = 0
        End If
        
        ' .... Video format
        If Len(INI.GetKeyValue("SETTING", "16943")) > 0 Then
            optVideo.Item(INI.GetKeyValue("SETTING", "16943")).value = True
        Else
            optVideo.Item(1).value = True
        End If
        
        ' .... Application StartUp
        If INI.GetKeyValue("SETTING", "STARTUP") <> Empty Then
            CheckStartUp.value = INI.GetKeyValue("SETTING", "STARTUP")
        Else
            CheckStartUp.value = 0
        End If
        
        ' .... Sound Downloads
        If INI.GetKeyValue("SETTING", "PLAYSOUNDOWNLOADS") <> Empty Then
            CheckSnd.value = INI.GetKeyValue("SETTING", "PLAYSOUNDOWNLOADS")
        Else
            CheckSnd.value = 0
        End If
        
        ' .... Printers Left Margin
        If INI.GetKeyValue("SETTING", "PRINTERSLEFTMARGIN") <> Empty Then
            cmbLeftMargin.ListIndex = INI.GetKeyValue("SETTING", "PRINTERSLEFTMARGIN")
        Else
            cmbLeftMargin.ListIndex = 1
        End If
        
        ' .... Printers Top Margin
        If INI.GetKeyValue("SETTING", "PRINTERSTOPMARGIN") <> Empty Then
            cmbTopMargin.ListIndex = INI.GetKeyValue("SETTING", "PRINTERSTOPMARGIN")
        Else
            cmbTopMargin.ListIndex = 1
        End If
        
        ' .... Default Printer
        If INI.GetKeyValue("SETTING", "DEFAULTPRINTRER") <> Empty And cmbPrinters.ListCount > 0 Then
            cmbPrinters.ListIndex = INI.GetKeyValue("SETTING", "DEFAULTPRINTRER")
        Else
            If cmbPrinters.ListCount > 0 Then cmbPrinters.ListIndex = 0
        End If
        
        ' .... Default Font printer
        If INI.GetKeyValue("SETTING", "DEFAULTFONTPRINTRER") <> Empty And cmbFonts.ListCount > 0 Then
            cmbFonts.ListIndex = INI.GetKeyValue("SETTING", "DEFAULTFONTPRINTRER")
        Else
            If cmbFonts.ListCount > 0 Then cmbFonts.ListIndex = 0
        End If
        
        ' .... Display Printer setup Dialog
        If INI.GetKeyValue("SETTING", "PRINTRERSETUPDIALOG") <> Empty Then
            CheckDisplayPrinter.value = INI.GetKeyValue("SETTING", "PRINTRERSETUPDIALOG")
        Else
            CheckDisplayPrinter.value = 0
        End If
        
        ' .... Font size
        If INI.GetKeyValue("SETTING", "DEFAULTFONTSIZE") <> Empty And cmbFontSize.ListCount > 0 Then
            cmbFontSize.ListIndex = INI.GetKeyValue("SETTING", "DEFAULTFONTSIZE")
        Else
            If cmbFontSize.ListCount > 0 Then cmbFontSize.ListIndex = 3
        End If
        
        ' .... Default download video Path
        If INI.GetKeyValue("SETTING", "DEFAULTVIDEOPATH") <> Empty Then
            txtDefaultPath.Text = INI.GetKeyValue("SETTING", "DEFAULTVIDEOPATH")
        Else
            txtDefaultPath.Text = App.Path + "\Downloads\"
            INI.CreateKeyValue "SETTING", "DEFAULTVIDEOPATH", txtDefaultPath.Text
        End If
        
        DEFAULTVIDEOPATH = txtDefaultPath.Text
        
End Sub

Private Sub DisableCloseWindowButton(frm As Form)
    On Local Error GoTo ErrorHandler
    Dim hSysMenu As Long
    hSysMenu = GetSystemMenu(frm.hWnd, 0)
    RemoveMenu hSysMenu, 6, MF_BYPOSITION
    RemoveMenu hSysMenu, 5, MF_BYPOSITION
Exit Sub
ErrorHandler:
    Err.Clear
End Sub

Private Sub SetFocusField(setTXT As TextBox, Optional sFocus As Boolean = True, Optional startChar As Integer = 0)
    On Error Resume Next
    If sFocus Then
        setTXT.SelStart = startChar
        setTXT.SelStart = startChar: setTXT.SelLength = Len(setTXT.Text) - startChar: setTXT.SetFocus
    Else
        setTXT.SelStart = 0: setTXT.SetFocus
    End If
End Sub

Public Sub SoftwareRegistered()
    ' .... Software already registered ?
    If YouTubeThanks.RegisterdSoftware Then
        Me.Caption = "YouTube Downloader v2.0.3"
        CaptionForm = "YouTube Downloader v2.0.3"
    Else
        Me.Caption = "YouTube Downloader v2.0.3 - NOT ACTIVATED - " & YouTubeThanks.DaysLeft & " Days left before expires!"
        CaptionForm = "YouTube Downloader v2.0.3 - NOT ACTIVATED - " & YouTubeThanks.DaysLeft & " Days left before expires!"
    End If
End Sub

Private Sub EndPlaySound()
    On Error Resume Next
    sndPlaySound ByVal vbNullString, 0&
End Sub

Private Sub CreateLinksList()
    Dim FF As Variant: FF = FreeFile
    On Local Error Resume Next
    If FileExists(App.Path + "\links.txt") Then _
        Open App.Path + "\links.txt" For Append As FF _
    Else Open App.Path + "\links.txt" For Output As FF
        Print #FF, "Video Title: " & UCase(lstUrlTitles.List(lstUrlTitles.ListIndex)) _
        & vbCrLf & Format(Now, "long date") & " - " & Time & vbCrLf _
        & "------------------------------------------" _
        & vbCrLf & lblDescription.Text & vbCrLf & "------------------------------------------"
    Close FF
    strTemp = Empty
End Sub

Private Function ValidFormat(ByVal sFileName As String) As Boolean
    Dim FSO As New FileSystemObject
    Dim f As File
    On Local Error Resume Next
    Set f = FSO.GetFile(sFileName)
    If UCase(FileExtensionFromPath(f.ShortName)) = ".FLV" Then
        ValidFormat = True
    Else
        ValidFormat = False
    End If
End Function

Private Sub PingS()
    Dim x As Integer
    On Local Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    txtPing.Text = ""
    txtPing.Text = PingServer.funcGetIPFromHostName(txtConvert.Text)
    
    If Len(txtPing.Text) < 1 Then
            img.Picture = pic(0).Picture
            lblSend.Caption = "No Connection..."
        Exit Sub
    End If
    
    PingServer.Ping txtPing.Text

    If strPing >= 400 Then
        img.Picture = pic(2).Picture: lblSend.Caption = "Bad Connection..."
    ElseIf strPing < 400 Then
        img.Picture = pic(1).Picture: lblSend.Caption = "Good Connection..."
    End If
    Screen.MousePointer = vbDefault
Exit Sub
ErrorHandler:
        Screen.MousePointer = vbDefault
    Err.Clear
End Sub

Private Sub StartEncodingListFiles(Optional dispMsg As Boolean = True)
    Dim i As Integer: i = 0
    If dispMsg Then
        If MsgBox("There are " & lstFileName.ListCount & " files in the list ready to be converted." & vbCrLf _
            & "Convert hte files now?", _
            vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
    End If
    
            PB.Max = lstFileName.ListCount
            For i = 0 To lstFileName.ListCount
                ' .............................. Start Conversion
                If estensione = "mpg" Then
                    If optVideo.Item(1).value Then
                        If YouTubeThanks.StartConversion(lstFileName.List(i), XVID, DvD43, estensione, SW_SHOWNORMAL) Then:
                    ElseIf optVideo.Item(0).value Then
                        If YouTubeThanks.StartConversion(lstFileName.List(i), XVID, DvD169, estensione, SW_SHOWNORMAL) Then:
                    End If
                ElseIf estensione = "avi" Then
                        If YouTubeThanks.StartConversion(lstFileName.List(i), XVID, , estensione, SW_SHOWNORMAL) Then:
                Else
                        If YouTubeThanks.StartConversion(lstFileName.List(i), , , estensione, SW_SHOWNORMAL) Then:
                End If
                ' .............................. End Conversion
                PB.value = i: lblPercentage.Caption = i & "%"
                DoEvents
            Next i
        lblPercentage.Caption = "0%": lstFileName.Clear
    On Local Error Resume Next
    PB.value = 0
Exit Sub
End Sub
Private Sub PrintList()
    Dim LEFT_MARGIN As Long
    Dim TOP_MARGIN As Long: Dim BOTTOM_MARGIN As Single
    Dim i As Integer: Dim text_line As String: Dim FF As Variant
    LEFT_MARGIN = cmbLeftMargin.List(cmbLeftMargin.ListIndex)
    TOP_MARGIN = cmbTopMargin.List(cmbTopMargin.ListIndex)
    
    On Error GoTo ErrorHandler
    
    If CheckDisplayPrinter.value = 1 Then
        If DialogPrintSetup(Me.hWnd) Then:
    End If
    
    If MsgBox("Print " & lstTitles.ListItems.Count & " items of the list?", vbYesNo + vbInformation + vbDefaultButton2, App.Title) = vbNo Then Exit Sub
    
    MousePointer = vbHourglass
    
    BOTTOM_MARGIN = Printer.ScaleHeight - 1160 '2160
    
    Printer.CurrentX = LEFT_MARGIN: Printer.CurrentY = TOP_MARGIN
    
    Printer.Font = cmbFonts.List(cmbFonts.ListIndex): Printer.FontSize = cmbFontSize.List(cmbFontSize.ListIndex)
    
    PB.Max = lstTitles.ListItems.Count
    
    For i = 1 To lstTitles.ListItems.Count
    
    If Printer.CurrentY + Printer.TextHeight(lstTitles.ListItems(i).Text) > BOTTOM_MARGIN Then
        ' .... Start New page...
        Printer.NewPage
        Printer.CurrentX = LEFT_MARGIN: Printer.CurrentY = TOP_MARGIN
    End If
            
    ' .... Print text Line
    Printer.Print Space(6) & "Title: " & lstTitles.ListItems(i).Text
    Printer.Print Space(6) & "Link: " & lstTitles.ListItems(i).SubItems(1)
    Printer.Print vbNewLine
    
    Printer.CurrentX = LEFT_MARGIN
    
    DoEvents
        PB.value = i
    Next i
    
    Printer.EndDoc

    MousePointer = vbDefault
    
Exit Sub
ErrorHandler:
    MousePointer = vbDefault
        MsgBox "Error #" & Err.Number & "." & vbCrLf & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub

Private Sub GetPrinters()
    Dim x As Printer: Dim i As Integer
    On Local Error GoTo ErrorHandler
        For Each x In Printers
            If Len(x.DeviceName) > 0 Then
                cmbPrinters.AddItem x.DeviceName
                i = i + 1
            End If
        Next
    If cmbPrinters.ListCount > 0 Then cmbPrinters.ListIndex = 0
Exit Sub
ErrorHandler:
    Err.Clear
End Sub

Private Sub PrintFontName()
    Dim jk As Integer: On Error GoTo ErrorHandler
    For jk = 0 To Printer.FontCount - 1
            cmbFonts.AddItem Printer.Fonts(jk)
        DoEvents
    Next jk
    If cmbFonts.ListCount > 0 Then cmbFonts.ListIndex = 0
    lblTotFont.Caption = cmbFonts.ListCount
Exit Sub
ErrorHandler:
    cmbFonts.Enabled = False
    Err.Clear
End Sub

Private Sub YouTubeThanks_VideoLinkDownloadURL(ByVal strLinks As Collection, ByVal strTranscodeLink As Collection)
    ' strTranscodeLink > contains the all transcoded URLs download of selected video...
    ' .... the secret is revealed :))
    ' How to:
    ' .................................... cut here ;)
    'Dim x As Variant
    'For Each x In strTranscodeLink
    '    cmbExtension.AddItem x
    'Next
    ' ....................................
    
    Dim i As Variant
    On Local Error GoTo ErrorHandler
    For Each i In strLinks
        cmbExtension.AddItem i
    Next
    Screen.MousePointer = vbDefault
Exit Sub
ErrorHandler:
    Screen.MousePointer = vbDefault
        MsgBox "Error #" & Err.Number & "." & vbCrLf & Err.Description, vbExclamation, App.Title
Err.Clear
End Sub


