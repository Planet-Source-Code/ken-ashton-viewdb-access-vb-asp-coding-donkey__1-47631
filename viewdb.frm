VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000016&
   Caption         =   " ViewDB"
   ClientHeight    =   5730
   ClientLeft      =   5025
   ClientTop       =   2535
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "viewdb.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5730
   ScaleWidth      =   7935
   Begin VB.Frame frSetup 
      BackColor       =   &H80000004&
      ForeColor       =   &H00800000&
      Height          =   4665
      Left            =   2430
      TabIndex        =   10
      Top             =   540
      Visible         =   0   'False
      Width           =   6195
      Begin VB.TextBox txbS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   17
         Left            =   5370
         MultiLine       =   -1  'True
         TabIndex        =   41
         ToolTipText     =   " Character used to delimit Combo Templates "
         Top             =   2550
         Width           =   375
      End
      Begin VB.TextBox txbS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   525
         Index           =   18
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   4050
         Width           =   5835
      End
      Begin VB.TextBox txbS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   16
         Left            =   5370
         MultiLine       =   -1  'True
         TabIndex        =   38
         ToolTipText     =   " Increment Index - normally 1, can be +/- anything"
         Top             =   2190
         Width           =   375
      End
      Begin VB.TextBox txbS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   15
         Left            =   5370
         MultiLine       =   -1  'True
         TabIndex        =   36
         ToolTipText     =   " Field  Index starts at 0, useful for offset array indexes "
         Top             =   1830
         Width           =   345
      End
      Begin VB.TextBox txbS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   14
         Left            =   2880
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   3630
         Width           =   1065
      End
      Begin VB.TextBox txbS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   13
         Left            =   2880
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   3300
         Width           =   1065
      End
      Begin VB.TextBox txbS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   12
         Left            =   2880
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   2970
         Width           =   1065
      End
      Begin VB.TextBox txbS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   11
         Left            =   2880
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   2610
         Width           =   1065
      End
      Begin VB.TextBox txbS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   10
         Left            =   2880
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   2280
         Width           =   1065
      End
      Begin VB.TextBox txbS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   9
         Left            =   1710
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   3630
         Width           =   645
      End
      Begin VB.TextBox txbS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   8
         Left            =   1710
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   3300
         Width           =   645
      End
      Begin VB.TextBox txbS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   7
         Left            =   1710
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   2970
         Width           =   645
      End
      Begin VB.TextBox txbS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   6
         Left            =   1710
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   2610
         Width           =   645
      End
      Begin VB.TextBox txbS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   5
         Left            =   1710
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   2280
         Width           =   645
      End
      Begin VB.TextBox txbS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   4
         Left            =   870
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   1590
         Width           =   645
      End
      Begin VB.TextBox txbS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   3
         Left            =   870
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   1260
         Width           =   645
      End
      Begin VB.TextBox txbS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   2
         Left            =   870
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   930
         Width           =   645
      End
      Begin VB.TextBox txbS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   1
         Left            =   870
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   600
         Width           =   645
      End
      Begin VB.TextBox txbS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   0
         Left            =   870
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   270
         Width           =   645
      End
      Begin VB.Label lbs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delimiter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   12
         Left            =   4620
         TabIndex        =   42
         Top             =   2610
         Width           =   600
      End
      Begin VB.Label lbs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step  +/-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   11
         Left            =   4620
         TabIndex        =   39
         ToolTipText     =   " Increment Index - normally 1, can be +/- anything"
         Top             =   2250
         Width           =   630
      End
      Begin VB.Label lbs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Index"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   10
         Left            =   4530
         TabIndex        =   37
         ToolTipText     =   " Field  Index starts at 0, useful for offset array indexes "
         Top             =   1890
         Width           =   765
      End
      Begin VB.Label lbs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "in 'Date' fields, replace                   with  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   29
         Top             =   3660
         Width           =   2820
      End
      Begin VB.Label lbs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "in 'Bool' fields, replace                   with  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   27
         Top             =   3330
         Width           =   2790
      End
      Begin VB.Label lbs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "in 'Curr' fields, replace                    with  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   25
         Top             =   3000
         Width           =   2805
      End
      Begin VB.Label lbs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "in 'Num' fields, replace                   with  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   23
         Top             =   2640
         Width           =   2805
      End
      Begin VB.Label lbs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "in 'Text' fields, replace                   with  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   21
         Top             =   2310
         Width           =   2790
      End
      Begin VB.Label lbs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Replace                    with Defined Size   (Memo=256)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   19
         Top             =   1650
         Width           =   3750
      End
      Begin VB.Label lbs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Replace                    with Custom Type  (0=Txt 1=Num 2=Cur 3=Bool 4=Date)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   17
         Top             =   1320
         Width           =   5475
      End
      Begin VB.Label lbs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Replace                    with Field Type      (0-255 Microsoft 'Type' number)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   15
         Top             =   990
         Width           =   5055
      End
      Begin VB.Label lbs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Replace                    with Field Index      (0-255 Field Order in Database)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   13
         Top             =   660
         Width           =   5070
      End
      Begin VB.Label lbs 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Replace                    with Field Name     (Actual Name of database Field)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   11
         Top             =   330
         Width           =   5115
      End
   End
   Begin VB.Frame frHelp 
      BackColor       =   &H80000016&
      Caption         =   "Help text - Edit if required"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   945
      Left            =   90
      TabIndex        =   8
      Top             =   3510
      Visible         =   0   'False
      Width           =   1785
      Begin VB.TextBox txbHelp 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   585
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   210
         Width           =   1485
      End
   End
   Begin VB.Frame frSplit 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      DragIcon        =   "viewdb.frx":0442
      DragMode        =   1  'Automatic
      Height          =   135
      Left            =   150
      MouseIcon       =   "viewdb.frx":0594
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "Drag split to preferred position"
      Top             =   210
      Width           =   7875
   End
   Begin VB.Frame frPane 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   60
      TabIndex        =   4
      Top             =   2010
      Width           =   2115
      Begin VB.TextBox txbCode 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   585
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   720
         Width           =   1995
      End
      Begin VB.ComboBox cbo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   930
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   390
         Width           =   1095
      End
      Begin VB.Label lbRem 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   10
         Left            =   15
         TabIndex        =   53
         Top             =   0
         Width           =   75
      End
      Begin VB.Label lbRem 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   15
         TabIndex        =   52
         Top             =   0
         Width           =   75
      End
      Begin VB.Label lbRem 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   15
         TabIndex        =   51
         Top             =   0
         Width           =   75
      End
      Begin VB.Label lbRem 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   15
         TabIndex        =   50
         Top             =   0
         Width           =   75
      End
      Begin VB.Label lbRem 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   15
         TabIndex        =   49
         Top             =   0
         Width           =   75
      End
      Begin VB.Label lbRem 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   15
         TabIndex        =   48
         Top             =   0
         Width           =   75
      End
      Begin VB.Label lbRem 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   15
         TabIndex        =   47
         Top             =   0
         Width           =   75
      End
      Begin VB.Label lbRem 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   15
         TabIndex        =   46
         Top             =   0
         Width           =   75
      End
      Begin VB.Label lbRem 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   15
         TabIndex        =   45
         Top             =   0
         Width           =   75
      End
      Begin VB.Label lbRem 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   15
         TabIndex        =   44
         Top             =   0
         Width           =   75
      End
      Begin VB.Label lbRem 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   75
         TabIndex        =   43
         Top             =   0
         Width           =   75
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3600
      Top             =   420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.ListBox ListBox 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Index           =   1
      IntegralHeight  =   0   'False
      Left            =   1980
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   630
      Width           =   1575
   End
   Begin VB.ListBox ListBox 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Index           =   0
      IntegralHeight  =   0   'False
      Left            =   60
      TabIndex        =   0
      Top             =   630
      Width           =   1905
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   6840
      TabIndex        =   58
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sql Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   6660
      TabIndex        =   57
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sql Insert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   5730
      TabIndex        =   56
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Un-Chk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   4620
      TabIndex        =   55
      Top             =   1020
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Chk All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3750
      TabIndex        =   54
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Fields"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1980
      TabIndex        =   3
      Top             =   300
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Tables"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   30
      TabIndex        =   1
      Top             =   300
      Width           =   1905
   End
   Begin VB.Menu mnFile 
      Caption         =   "&File"
      Begin VB.Menu mnOpenDatabase 
         Caption         =   "&Open Database"
      End
      Begin VB.Menu FileMenuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnCloseDatabase 
         Caption         =   "C&lose database"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnFileMenuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnMRU 
         Caption         =   "E&xit"
         Index           =   0
      End
   End
   Begin VB.Menu mnCreate 
      Caption         =   "&Create"
      Begin VB.Menu mnCreates 
         Caption         =   "List Table Fields"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnCreates 
         Caption         =   "List All Tables && Fields"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnCreates 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnCreates 
         Caption         =   "&Code"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mnCreates 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnCreates 
         Caption         =   "Connection String Access97"
         Index           =   5
      End
      Begin VB.Menu mnCreates 
         Caption         =   "Connection String Access2000"
         Index           =   6
      End
      Begin VB.Menu mnCreates 
         Caption         =   "Connection String Custom"
         Index           =   7
      End
      Begin VB.Menu mnCreates 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnCreates 
         Caption         =   "Sql INSERT INTO"
         Index           =   9
      End
      Begin VB.Menu mnCreates 
         Caption         =   "Sql UPDATE"
         Index           =   10
      End
      Begin VB.Menu mnCreates 
         Caption         =   "Sql SELECT"
         Index           =   11
      End
   End
   Begin VB.Menu mnOption 
      Caption         =   "Options"
      Begin VB.Menu mnOptions 
         Caption         =   "Allow Aposthophe's"
         Index           =   0
      End
      Begin VB.Menu mnOptions 
         Caption         =   "Auto Code"
         Index           =   1
      End
      Begin VB.Menu mnOptions 
         Caption         =   "Auto Copy"
         Index           =   2
      End
      Begin VB.Menu mnOptions 
         Caption         =   "Auto Open"
         Index           =   3
      End
      Begin VB.Menu mnOptions 
         Caption         =   "MS Sans Serif Font"
         Index           =   4
      End
      Begin VB.Menu mnOptions 
         Caption         =   "Show Tool-Tips"
         Index           =   5
      End
      Begin VB.Menu mnOptions 
         Caption         =   "Access 97 default"
         Index           =   6
      End
      Begin VB.Menu mnOptions 
         Caption         =   "Always On-Top"
         Index           =   7
      End
      Begin VB.Menu mnOptions 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnOptions 
         Caption         =   "Remove Combo entry"
         Enabled         =   0   'False
         Index           =   9
      End
   End
   Begin VB.Menu mnSetup 
      Caption         =   "Setup"
   End
   Begin VB.Menu mnPrint 
      Caption         =   "Print"
      Begin VB.Menu mnPrints 
         Caption         =   "Print Setup"
         Index           =   0
         Begin VB.Menu mnPrintSetup 
            Caption         =   "Left margin"
            Index           =   0
         End
         Begin VB.Menu mnPrintSetup 
            Caption         =   "Top margin"
            Index           =   1
         End
         Begin VB.Menu mnPrintSetup 
            Caption         =   "Lines per page"
            Index           =   2
         End
      End
      Begin VB.Menu mnPrints 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnPrints 
         Caption         =   "Print ..."
         Index           =   2
      End
   End
   Begin VB.Menu mnHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnHelps 
         Caption         =   "&Help.txt (editable)"
         Index           =   0
      End
      Begin VB.Menu mnHelps 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnHelps 
         Caption         =   "About"
         Index           =   2
      End
   End
   Begin VB.Menu mnCancel 
      Caption         =   "Cancel"
      Visible         =   0   'False
   End
   Begin VB.Menu mnCloseHelp 
      Caption         =   "Close"
      Visible         =   0   'False
   End
   Begin VB.Menu mnSaveHelp 
      Caption         =   "Save changes"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Const MXU As Integer = 3            ' Max Filenames/Menu MRU entries
Const MXT As Integer = 10           ' Number of Template entries
Const MXL As Integer = 12           ' Max number of stored templates
Const vbGrey = &H8000000F           ' Default Form color

Dim IsOnTop As Boolean              ' Indicates current 'OnTop' state
Dim MsSans As Boolean               ' Indicates if Sans Serif Font selected
Dim Inhibit As Boolean              ' Stops control 'event code' execution
Dim Access97 As Boolean             ' True if Setup for Access 97 by default
Dim z(0 To 18) As Boolean           ' True for each valid entry in Setup array
Dim SaveCombo As Boolean            ' Flags when Combo needs DblClick saving
Dim SkipDialog As Boolean           ' Skips manual OpenFile Dialog if MRU used

Dim tt As String                    ' CSV Holds crude, unique Datatype names
Dim MRU As String                   ' Str holds MRU list  (delimited ',')
Dim TPL As String                   ' Str holds templates (delimited specs(17))
Dim MyDB As String                  ' Str name of currently open database
Dim TheConn As String               ' Connection String currently in use.
Dim TheTable As String              ' TableName currently is use.
Dim LastPathOpen As String          ' Path of last MDB opened
Dim LastPathConn As String          ' Path of last MDB for Conn string
Dim Selects(0 To 5, 0 To 255) As Boolean ' Crude Temporary Listbox memory for 6 tables

Dim Specs(0 To 18) As String        ' Setups array, to avoid accessing controls
Dim FldNam() As String              ' Hold the 'DefinedSize' for each field
Dim TxtTyp() As String              ' Hold the 'DefinedSize' for each field
Dim DefSiz() As Long                ' Hold the 'DefinedSize' for each field
Dim FldTyp() As Integer             ' Holds the MS Field 'Type' number (0-255)
Dim CusTyp() As Integer             ' Holds my custom field 'Type' numbers (0-4)
Dim TxtLens(0 To 1) As Integer      ' Save the length of the largest Fieldname & textType
Dim SplitTop As Single              ' Position of Form Horiz. Split bar
Dim TableCount As Integer           ' Number of 'User' Tables in database
Dim FieldCount As Integer           ' Number of Fields in currently selected Table
Dim prtTop As String                ' Printer Top margin  (in Lines)
Dim prtLeft As String               ' Printer Left margin (in Characters)
Dim prtLines As String              ' Printer Lines per page
Dim LB(0 To 1) As ListBox           ' Listbox object variables
Dim RemIndex As Integer             ' Current 'Reminder Tip'
Dim LastTable As Integer            ' Remembers last table selected in DB
'
' Used for the ini file
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
'
' Used to set a tabStop in Listbox
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const LB_SETTABSTOPS = &H192

' Used to 'Set Window on top'
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1

Sub ClearAll()

    TheConn = ""                ' Current Connection String
    TheTable = ""               ' Current table Name
    FieldCount = 0              '
    TableCount = 0              '
    ClearListBoxes 2            '
    FormCaption                 '
    
End Sub

Function FormatFieldItem(u As Integer)

    FormatFieldItem = Mid(CStr(u) & Space(5), 1, 5) & _
                            Mid(FldNam(u) & Space(TxtLens(0)), 1, TxtLens(0) + 2) & _
                            Mid(TxtTyp(u) & Space(TxtLens(1)), 1, TxtLens(1) + 2) & _
                            Mid(DefSiz(u) & Space(5), 1, 5) & _
                            Mid("Cus[" & CusTyp(u) & "]" & Space(3), 1, 8) & _
                            "Ms[" & FldTyp(u) & "]"

End Function

Sub QualifyDefaults()
'
' If first time user or .ini diskfile blew up, give user a starting set of values
'
    If txbS(0) = "" Then txbS(0) = "$$"           ' Field Names
    If txbS(1) = "" Then txbS(1) = "##"           ' Field Index
    If txbS(2) = "" Then txbS(2) = "{ft}"         ' Field Type
    If txbS(3) = "" Then txbS(3) = "{ct}"         ' Custom Field Type
    If txbS(4) = "" Then txbS(4) = "{ds}"         ' DefinedSize
    If txbS(5) = "" Then txbS(5) = "{nul}"        ' Text Fields
    If txbS(6) = "" Then txbS(6) = "{nul}"        ' Number Fields
    If txbS(7) = "" Then txbS(7) = "{nul}"        ' Currency Fields
    If txbS(8) = "" Then txbS(8) = "{nul}"        ' Boolean Fields
    If txbS(9) = "" Then txbS(9) = "{nul}"        ' Date Fields
    If txbS(10) = "" Then txbS(10) = "Nul"        ' Text Field Subs
    If txbS(11) = "" Then txbS(11) = "NulN"       ' Number Field Subs
    If txbS(12) = "" Then txbS(12) = "NulC"       ' Currency Field Subs
    If txbS(13) = "" Then txbS(13) = "NulB"       ' Boolean Field Subs
    If txbS(14) = "" Then txbS(14) = "NulD"       ' Date Field Subs
    If txbS(15) = "" Then txbS(15) = "0"          ' Field Index Number Offset (if desired)
    If txbS(16) = "" Then txbS(16) = "1"          ' Field Index Increment (Normally 1)
    If txbS(17) = "" Then txbS(17) = "?"          ' Templates Delimiter
    If txbS(18) = "" Then txbS(18) = Chr(34) & "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Chr(34) & " & MyDB & " & Chr(34) & ";Persist Security Info=False" & Chr(34)

End Sub

Sub ClearListBoxes(i As Integer)

'   i = 0  means Clear lb(0)
'   i = 1  means Clear Lb(1)
'   i = 2  means Clear Both

'   We use the appropriate moment to also reset Listbox 'Header' labels & colors

    Inhibit = True                      ' Stop the Listbox triggering events
    
    If i <> 1 Then
            LB(0).Clear
            SetBackColor LB(0)
            lbl(0) = "Tables"
            DoEvents
    End If
    
    If i <> 0 Then
            LB(1).Clear
            SetBackColor LB(1)
            lbl(1) = "Fields"
            FieldCount = 0
            DoEvents
    End If
    
    Inhibit = False
    
End Sub

Sub DoIni(Key As String, vData As Variant, WriteMode As Boolean)
Dim s As String

'   Quick & dirty function for simply reading and writing to ini file.
'   Incoming flag 'WriteMode', if true, Writes vData to the ini file
'   else if False, will read vData which is a Variant to allow
'   controls to be used for writing simplicity.
'
'   Rationale: When developing, can quickly add any variable you want to
'   save, and simply make the Key same as variable name if desired.
'
'   Example Write usage:    DoIni "SetOnTop", Check1.value, True
'   Example Read usage:     DoIni "SetOnTop", b, False: Check1.value = b
'
'   Important Note. The name of the ini file is the application name
'   Important Note. All entries are written under [Default] section
'

    On Error Resume Next

    If WriteMode Then
    
        Dim ret As Long
        
        s = CStr(vData)
        ret = WritePrivateProfileString("Default", Key, s, _
                (LCase(App.Path & "\" & App.EXEName) & ".ini"))
    
    Else
    
        Dim bf As String, Lx As Long, ln As Integer
        ln = 256                                    ' Get up to 256 chars
        bf = String(ln, Chr(0))                     ' Create buffer & Get
        Lx = GetPrivateProfileString("Default", Key, "", bf, _
                ln, (LCase(App.Path & "\" & App.EXEName) & ".ini"))
        vData = Left$(bf, Lx)                       ' Remove the nulls
        If Err Or Right(Lx, 1) <> Chr(0) Then       ' If error up overflowed
            Error.Clear                             ' buffer (real big string)
            On Error Resume Next                    ' then reset any errors
            ln = 8192                               ' bump the buffer size to 8K
            bf = String(ln, Chr(0))                 ' and try and Get again
            Lx = GetPrivateProfileString("Default", Key, "", bf, _
                ln, (LCase(App.Path & "\" & App.EXEName) & ".ini"))
            vData = Left(bf, Lx)                 ' Assign Value to vData
        End If
    End If
    
    If Err Then                                     ' Error messages
        MsgBoxErr Choose(Abs(WriteMode) + 1, "Read", "Write") & _
                " error in Routine DoIni()"
        Err.Clear
    End If

    DoEvents        ' Might be several successive calls to this routine
                    ' during startup, so give Windows some time
End Sub


Sub DeleteListItem(Lst, Itm As String, d As String)
    Dim Tmp As String
'
'   Removes item (Itm) from a list (lst) of delimited variables (Delimiter d)
'
    Tmp = d & Lst & d                           ' Copy list to tmp & append delimiters
    Tmp = Trim(Replace(Tmp, d & Itm & d, ""))   ' Append delimiters to item d & replace
    If Len(Tmp) > 0 Then                        ' any match found in the list (tmp)
        If Mid(Tmp, 1, 1) = d Then              ' If len(tmp) then it wasn't last listitem
            Lst = Mid(Tmp, 2)                   ' so remove and leading delimiter or
        ElseIf Mid(Tmp, Len(Tmp), 1) = d Then   ' trailing delimiter form the tmp list
            Lst = Mid(Tmp, 1, Len(Tmp) - 1)     ' string and return it.
        End If
    Else
        Lst = ""                                ' List now empty, we must have
    End If                                      ' removed the last item in list

End Sub

Function FormatVarStr(Varnam As String, t As Integer) As String
    Dim q As String
'
'   Used to assist in the repetitive assembling of the SQL statements
'
    q = Chr(34)
    
    Select Case t
        Case 0  ' Text
            FormatVarStr = "'" & q & " & " & Varnam & " & " & q & "'"
        Case 1  ' Number
            FormatVarStr = q & " & " & Varnam & " & " & q
        Case 2  ' Currency
           FormatVarStr = q & " & " & Varnam & " & " & q
        Case 3  ' Boolean
           FormatVarStr = q & " & " & Varnam & " & " & q
        Case 4  ' Date
            FormatVarStr = "#" & q & " & " & Varnam & " & " & q & "#"
        
    End Select
    
End Function

Sub FormCaption()

    If Len(MyDB) > 0 Then
        Me.Caption = " ViewDB  " & MyDB & " [" & CStr(TableCount) & " Tables]"
    Else
        Me.Caption = " ViewDB  [No File]"
    End If
    
    DoEvents
    
End Sub

Function GetMdbFileName(Pth As String) As String

        On Error Resume Next
        
        CommonDialog1.InitDir = Pth                     ' Preset Folder in use
        CommonDialog1.DefaultExt = "mdb"                ' Only interest in
        CommonDialog1.FileName = "*.mdb"                ' Access databases, ie
        CommonDialog1.Filter = "MDB (*.mdb)|ALL (*.*)"  ' *.mdb
        CommonDialog1.CancelError = True                ' Allow User cancelling
        
        CommonDialog1.Action = 1                        ' Show the 'Open' dialog
        
        If Err Then                                     ' If Error, the User cancelled
            GetMdbFileName = ""                         ' so return null string
        Else                                            ' else
            GetMdbFileName = CommonDialog1.FileName     ' Assign the dialog filename
        End If
        
        DoEvents                                        ' Repaint hole left by dialog
        
End Function

Function CreateSQL(Mode As Integer) As String
    Dim i As Integer
    Dim FieldList As String
    Dim InsFixed As String
    Dim InsArray As String
    Dim UpdFixed As String
    Dim UpdArray As String
    Dim Tbl As String
    Dim flg As Boolean
    Dim Tmp As String
    Dim Tmp2 As String
    Dim q As String
'
'   Generates 2 versions of each Sql command for Insert, Update & Select
'
    q = Chr(34)
    If LB(0).ListIndex >= 0 Then                    ' If we have the Table name
        Tbl = LB(0).List(LB(0).ListIndex)           ' then assign it to Tbl
    Else
        MsgBox "No Table selected!"
        Exit Function
    End If
    
    If Not HaveItems(LB(1)) Then                    ' Quit if no fields selected
        MsgBox "No Fields Selected!"
        Exit Function
    End If
    
    For i = 0 To FieldCount - 1                     ' Loop thro Fields
    
        If LB(1).Selected(i) Then                   ' If Field is Selected then go
        
            If flg Then                             ' Avoid ',' if no fields yet
                FieldList = FieldList & ", "        ' FieldList for INSERT/SELECT
                InsFixed = InsFixed & ", "          ' For Insert, FixedVars
                UpdFixed = UpdFixed & ", "          ' For Update/Select, Fixed vars
                InsArray = InsArray & ", "          ' For Insert, Array Vars
                UpdArray = UpdArray & ", "          ' For update/Select, Array Vars
            End If
            
            FieldList = FieldList & FldNam(i)       ' FieldList for INSERT, SELECT
            UpdFixed = UpdFixed & FldNam(i) & "="   ' For Update, Fixed Vars
            UpdArray = UpdArray & FldNam(i) & "="   ' For Update, Array Vars
                
            Select Case CusTyp(i)                   ' Here's where a Custom Type useful
                Case 0  ' Text
                    InsFixed = InsFixed & FormatVarStr("txt_" & FldNam(i), CusTyp(i))
                    UpdFixed = UpdFixed & FormatVarStr("txt_" & FldNam(i), CusTyp(i))
                    InsArray = InsArray & FormatVarStr("MyArr(" & CStr(i) & ")", CusTyp(i))
                    UpdArray = UpdArray & FormatVarStr("MyArr(" & CStr(i) & ")", CusTyp(i))
                Case 1  ' Number
                    InsFixed = InsFixed & FormatVarStr("num_" & FldNam(i), CusTyp(i))
                    UpdFixed = UpdFixed & FormatVarStr("num_" & FldNam(i), CusTyp(i))
                    InsArray = InsArray & FormatVarStr("MyArr(" & CStr(i) & ")", CusTyp(i))
                    UpdArray = UpdArray & FormatVarStr("MyArr(" & CStr(i) & ")", CusTyp(i))
                Case 2  ' Currency
                    InsFixed = InsFixed & FormatVarStr("cur_" & FldNam(i), CusTyp(i))
                    UpdFixed = UpdFixed & FormatVarStr("cur_" & FldNam(i), CusTyp(i))
                    InsArray = InsArray & FormatVarStr("MyArr(" & CStr(i) & ")", CusTyp(i))
                    UpdArray = UpdArray & FormatVarStr("MyArr(" & CStr(i) & ")", CusTyp(i))
                Case 3  ' Boolean
                    InsFixed = InsFixed & FormatVarStr("bool_" & FldNam(i), CusTyp(i))
                    UpdFixed = UpdFixed & FormatVarStr("bool_" & FldNam(i), CusTyp(i))
                    InsArray = InsArray & FormatVarStr("MyArr(" & CStr(i) & ")", CusTyp(i))
                    UpdArray = UpdArray & FormatVarStr("MyArr(" & CStr(i) & ")", CusTyp(i))
                Case 4  ' Date
                    InsFixed = InsFixed & FormatVarStr("date_" & FldNam(i), CusTyp(i))
                    UpdFixed = UpdFixed & FormatVarStr("date_" & FldNam(i), CusTyp(i))
                    InsArray = InsArray & FormatVarStr("MyArr(" & CStr(i) & ")", CusTyp(i))
                    UpdArray = UpdArray & FormatVarStr("MyArr(" & CStr(i) & ")", CusTyp(i))
            End Select
            If Not flg Then flg = True  ' ok we've had a field, so flag that
        End If                          ' comma delimiters are need from now on
    Next
        
        
    Select Case Mode
    
        Case 0  ' INSERT
            Tmp = q & "INSERT INTO " & Tbl & " (" & FieldList & ") VALUES (" & InsFixed & ");" & q
            Tmp2 = q & "INSERT INTO " & Tbl & " (" & FieldList & ") VALUES (" & InsArray & ");" & q
        
        Case 1  ' UPDATE
            Tmp = q & "UPDATE " & Tbl & " SET " & UpdFixed & " WHERE ID=" & q & " & CStr(MyId) & " & q & ";" & q
            Tmp2 = q & "UPDATE " & Tbl & " SET " & UpdArray & " WHERE ID=" & q & " & CStr(MyId) & " & q & ";" & q
        
        Case 2  ' SELECT
            Tmp = q & "SELECT " & FieldList & " FROM " & Tbl & " WHERE " & q & " & MyCriteraStr & " & q & " ORDER BY " & q & " & MySortStr & " & q & ";" & q
            Tmp2 = q & "SELECT " & FieldList & " FROM " & Tbl & " WHERE " & q & " & MyCriteriaStr & " & q & " ORDER BY " & q & " & MySortStr & " & q & ";" & q
    
    End Select
'
'   Format output to personal taste
    
    CreateSQL = "FixedVar/Object style -" & vbCrLf & vbCrLf & _
                Tmp & _
                vbCrLf & vbCrLf & vbCrLf & _
                "Array based style -" & vbCrLf & vbCrLf & _
                Tmp2 & _
                vbCrLf

End Function

Function HaveItems(Lbx As ListBox) As Boolean
    Dim i As Integer

'   Simple check to see if any items in the listbox are selected (checked)

    HaveItems = True
    
    For i = 0 To Lbx.ListCount - 1              ' Loop thro the items and
        If Lbx.Selected(i) Then Exit Function   ' Quit as soon as we find
    Next                                        ' a selected item
    
    HaveItems = False                           ' end of list, return false
    
End Function

Sub LoadSpecs()
    Dim i As Integer
'
'   Loads array with Setup values, and a boolean array to indicate values exist
'
    For i = 0 To 18
        Specs(i) = txbS(i)                          ' Get Setups into string array
        z(i) = Sgn(Len(Specs(i))) * -1              ' Set Z() true if value present.
        If i < 10 Then lbRem(i) = Specs(i)          ' Load Reminder labels
    Next

End Sub

Sub LoadCombo()
    Dim arr, i As Integer
'
'   Load the 'templates' combo from var TPL with previously saved user templates
'
    cbo.Clear                               ' Clear old combo entries.
    If Len(Specs(17)) > 0 Then              ' If we have a vaild delimiter
        arr = split(TPL, Specs(17))         ' Split delimited var to array
        For i = 0 To UBound(arr)            ' Loop for Templates
            If Len(arr(i)) > 0 Then         ' If valid template then
                cbo.AddItem arr(i)          ' add to the combo list
            End If
            DoEvents
        Next
    End If
    DoEvents
    Set arr = Nothing

End Sub

Sub MsgBoxErr(Msg As String)
    Dim b As String
'
'   Quick Standardized way for displaaying any program errors
'
    b = Msg & vbCrLf & vbCrLf
    b = b & "Error Number: " & CStr(Err) & vbCrLf
    b = b & "Error Message: " & Error
    Err.Clear
    Moff
    MsgBox b, , " Program Error!"
    
End Sub

Sub Remind(z As Integer)
    Dim L, i%

    If z < 10 Then                              ' Labels 0-9 serve as buttons
        If lbRem(0).ForeColor <> vbGrey Then    ' so if they are visible (not greyed
            lbRem(RemIndex).FontBold = False    ' out, then restore the last selection
            lbRem(RemIndex).ForeColor = vbBlack ' to Black Normal font and set the
            lbRem(z).FontBold = True            ' new selection to Blue Bold font
            lbRem(z).ForeColor = vbBlue
            lbRem(10).ForeColor = vbBlue
            RemIndex = z                        ' Update 'Last' selection
        End If
    End If
    
'   Reposition labels because their width varies with Font changes (we want this)
    
    L = 20                                      ' Add slight left margin
    Do While i < 10                             ' and loop set the Left
        lbRem(i).Left = L                       ' for each label and
        L = L + lbRem(i).Width + 10             ' add up the widths
        i = i + 1                               ' (they change with Bold)
    Loop
    
    lbRem(10).Left = L + 25                     ' and finally position the
    lbRem(10).Width = Width - L - 45            ' Reminder Text label
'
'   Simply reload the 'reminder' labels each time, as they can change with setup
'
    Select Case z
    
        Case 0
            lbRem(10) = " Inserts 'FieldName' in list"
        Case 1
            lbRem(10) = " Inserts field 'Index' (number) in list"
        Case 2
            lbRem(10) = " Inserts MS field 'DataType' (0-255) in list"
        Case 3
            lbRem(10) = " Inserts 'CustomType' (Txt=0 Num=1 Curr=2 Bool=3 Date=4)"
        Case 4
            lbRem(10) = " Inserts 'DefinedSize' (Text 1-255 max, Memo = 256)"
        Case 5
            lbRem(10) = " Gets replaced with '" & txbS(10) & "' if field is TEXT"
        Case 6
            lbRem(10) = " Gets replaced with '" & txbS(11) & "' if field is NUMBER"
        Case 7
            lbRem(10) = " Gets replaced with '" & txbS(12) & "' if field is CURRENCY"
        Case 8
            lbRem(10) = " Gets replaced with '" & txbS(13) & "' if field is BOOLEAN"
        Case 9
            lbRem(10) = " Gets replaced with '" & txbS(14) & "' if field is DATE"

End Select

    
End Sub

Sub LoadSaveSelects(b As String, SaveIt As Boolean)
    Dim j%, i%, k%
    On Error Resume Next
    
'   Primitive routine to load/save Listbox 'Selected' list into a string for ini file
'   Sufficient for 6 tables (0-5) and 256 fields (0-255). So the boolean array var
'   Selects(0 to 5, 0 to 255) is saved to, or restored from a 1536byte string (b)
'   It is only to provide the ability to move between tables without resetting the
'   the checkboxes on the Fields listbox. I save the array to ini, costs nothing and
'   at least while you are working, and Auto-reloading the same database, the
'   Fields Listbox 'Checked/Selected' settings are preserved.
'

    If SaveIt Then                                          ' If the incoming flag is
        b = ""                                              ' true then we are saving,
        For j = 0 To 5                                      ' clear and build b using
            For i = 0 To 255                                ' '0'=false '1'=True
                b = b & Chr(48 + Abs(Selects(j, i)))
            Next
        Next
    ElseIf Len(b) = 1536 Then                               ' If reading, and b=1536 bytes
        For j = 0 To 5                                      ' then loop thro' and restore
            For i = 0 To 255                                ' the Boolean array,
                k = k + 1                                   ' converting '0' to zero
                Selects(j, i) = 48 - Asc(Mid(b, k, 1))      ' and '1' to -1 (True)
            Next
        Next
    Else                                                    ' If we get a duff string
        For j = 0 To 5                                      ' simply reset all the
            For i = 0 To 255                                ' listbox 'Selected's to true
                Selects(j, i) = True
            Next
        Next
    End If

End Sub

Sub SaveNewComboItem()
    Dim Tmp As String
    
    Tmp = cbo.Text                                      ' Store any text to tmp
    If Len(Tmp) > 0 Then                                ' If text there, then
        UpdateDelimitedList TPL, Tmp, Specs(17), MXT    ' add to the list of stored
        LoadCombo                                       ' templates and reload the
        cbo.Text = Tmp                                  ' combo, restore the text
    End If

End Sub

Sub SetChecked(Lbx As ListBox, X As Boolean)
    Dim i As Integer, li As Integer
    
'   Set all Listbox checkboxes to x  (True or False)
'
    If Lbx.ListCount > 0 Then               ' If the listbox has content
        li = Lbx.ListIndex                  ' Note the current Listindex
        Inhibit = True                      ' Block Listbox_click events
        For i = 0 To Lbx.ListCount - 1      ' Loop thro' items
            Lbx.Selected(i) = X             ' Set True/False as requested
        Next
        Lbx.ListIndex = li                  ' Restore Listindex
        Inhibit = False                     ' Re-Enable events
    End If
    
    
End Sub

Sub SetFonts(MsSans As Boolean)

    Const MonoFont = "Courier New"
    Const PropFont = "MS Sans Serif"
    
    If MsSans Then                      ' Ms Sans Serif required then
        LB(0).FontName = PropFont       ' load relevant controls with
        LB(1).FontName = PropFont       ' the (proportional) font
        txbCode.FontName = PropFont
        txbHelp.FontName = PropFont
    Else
        LB(0).FontName = MonoFont       ' otherwise use the 'Mono-Spaced'
        LB(1).FontName = MonoFont       ' Courier New 'Fixed Width' font.
        txbCode.FontName = MonoFont
        txbHelp.FontName = MonoFont
    End If
    
End Sub

Sub SetMenus(X As Integer)
'
'       Enable and Disable Menus and Controls as requested by X

'       Normal x=0, Busy x=1, Setup x=2,  Help x=3

        If X = 1 Then Mon
        DoEvents
        mnFile.Enabled = (X = 0)
        mnCreate.Enabled = (X = 0)
        mnOption.Enabled = (X = 0)
        mnSetup.Enabled = (X = 0)
        mnPrint.Enabled = (X = 0)
        mnHelp.Enabled = (X = 0)
        frPane.Enabled = (X = 0)
        mnSaveHelp.Visible = False
        DoEvents
        frSetup.Visible = (X = 2)
        mnCloseHelp.Visible = (X = 3) Or (X = 2)
        mnCancel.Visible = (X = 2)
        frHelp.Visible = (X = 3)
        frPane.Visible = Not (X = 3)
        If X <> 1 Then Moff
        DoEvents
        
End Sub

Sub SetOnTop(frm As Form, tf As Boolean)
Dim Success As Long
Rem Success <> 0 When Successful
Rem Tf=True or False (On Top and NOT On Top
Rem Frm =  Form1, Form2 ... (Form Name) etc.
On Error Resume Next

    If tf = True Then
        Success = SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        IsOnTop = True
        If Err Then MsgBoxErr "SetOnTop(" & frm.Name & "," & CStr(tf) & ") (Note: Failed to put Form on Top)"
    Else
        Success = SetWindowPos(frm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        IsOnTop = False
        If Err Then MsgBoxErr "SetOnTop(" & frm.Name & "," & CStr(tf) & ") (Note: Failed to remove Form on Top)"
    End If '
    
End Sub

Function GetFields(Tbl As String) As Boolean
'   Needs reference to Microsoft ActiveX Data Objects 2.5 Library
    Dim cn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim u As Integer    ' Counter
    Dim b As String

    On Error Resume Next
    
    If Len(TheConn) = 0 _
        Or TableCount = 0 _
        Or Len(Tbl) = 0 Then Exit Function
'
'   Get everything into arrays, we want flexibilty - speed not so important
'
    ReDim FldNam(0 To 0)
    ReDim FldTyp(0 To 0)                                ' Clear any prev values
    ReDim CusTyp(0 To 0)                                ' (Doesn't clear items (0)
    ReDim DefSiz(0 To 0)                                '  but won't be a problem)
    ReDim TxtTyp(0 To 0)
'
'   Open a recordset and for each Field read in Name, Type & DefinedSize
'   Getting the full recordset with all fields obviously has a high
'   overhead, but avoids other problems and speed not critical in this app
'   Some recordset info is stored in arrays for later use in creating code.
'
'   Needs reference to Microsoft ActiveX Data Objects 2.5 Library
    Set cn = New ADODB.Connection       ' Prime the connection Object cn
    cn.ConnectionString = TheConn '
   ' cn.CursorLocation = adUseClient
    cn.Mode = adModeRead
    cn.Open                             ' Open the connection and use it
    Set rs = cn.Execute("SELECT * FROM [" & Tbl & "]", adOpenStatic, 1)
    
    FieldCount = rs.Fields.Count                        ' Read field count into a var
    If Not Err Then
    
        ReDim FldNam(0 To FieldCount - 1)               ' Re-dimension arrays to match
        ReDim FldTyp(0 To FieldCount - 1)               ' number of fields in the
        ReDim FldTyp(0 To FieldCount - 1)               ' currently selected table
        ReDim CusTyp(0 To FieldCount - 1)
        ReDim DefSiz(0 To FieldCount - 1)
        ReDim TxtTyp(0 To FieldCount - 1)
    
        With rs
            For u = 0 To FieldCount - 1
                FldNam(u) = .Fields(u).Name             ' get 'Name' of each field
                FldTyp(u) = .Fields(u).Type             ' get data 'Type' of field (0-255)
                DefSiz(u) = .Fields(u).DefinedSize      ' get 'DefinedSize' of each field
                '   x = .Fields(u).Attributes           '
                '   x = .Fields(u).NumericScale         ' Other useable stuff if reqd.
                '   x = .Fields(u).Precision            ' or if needed in future
            
                If DefSiz(u) > 255 Then                 ' If defined size>255, it must be
                    DefSiz(u) = 256                     ' a Memo field so simply force size
                End If                                  ' to 256 to preserve column size
            
                EnumerateDataType u                     ' Update other arrays with info
            
            Next u
        
            For u = 0 To FieldCount - 1                 ' Load Listbox with formatted
                LB(1).AddItem FormatFieldItem(u)        ' Field info. Waited till now, we
                LB(1).Selected(u) = Selects(LB(0).ListIndex, u)
            Next                                        ' needed to know longest string
        
        End With
    End If
    
    If Err Then                                         ' Better check for errors as
        Moff                                            ' we were reading ADO recordset
        MsgBoxErr "No Fields GetField(" & Tbl & ")"     ' If there were, send message"
        lbl(1) = "No Fields"                            ' Clear the labels, return false
        FieldCount = 0                                  ' Indicate zero fields
        ClearListBoxes 1
        GetFields = False                               ' to reset Database 'Open' flag
    Else
        GetFields = True                                ' Return true, DB is open
        lbl(1) = CStr(FieldCount) & " Fields"           ' Update Listbox header
        DoEvents
    End If

    SetBackColor LB(1)
    
    adoTidy rs                                          ' Disconnect recordset
    adoTidy cn                                          ' and close connection
    
End Function

Function GetTables(MyDB As String, Tbl As String, ByVal Access97 As Boolean) As Integer
'   Needs reference to Microsoft ActiveX Data Objects 2.5 Library
    Dim cn As New ADODB.Connection
    Dim n As String
    Dim rs As New ADODB.Recordset
    On Error Resume Next
    
'   Opens the database according to Version selected in 'Settings' menu
'   If Access 97 selected, and a Access 2000 database is opened, the
'   error is trapped, user warned, & attempt to re-open it in 2000 mode.
'   If successful, does NOT change the 'Settings' from 97 to 2000


    If Access97 Then ' Connection for Jet 3.5, Access 8.0, Office97, VB5
        TheConn = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=" & _
                    MyDB & ";Persist Security Info=False"
    Else
retry:         ' Connection for Jet 4.0, Access 2000, Office2000, VB6
        TheConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                    MyDB & ";Persist Security Info=False"
    End If

    Set cn = New ADODB.Connection       ' Prime the connection Object cn
    cn.ConnectionString = TheConn
    cn.CursorLocation = adUseClient
    cn.Mode = adModeRead
    
    cn.Open                             ' Open the connection and use it
    Set rs = cn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty))
    
    If Err Then                         ' If we get an error, it could be
        If Access97 Then                ' later Access ver. than selected
            MsgBox "Access 97 mode is enabled. (File menu)" & _
                    vbCrLf & vbCrLf & "Failed to Open " & MyDB & _
                    vbCrLf & vbCrLf & "Error was: " & Error & _
                    vbCrLf & vbCrLf & _
                   "Will try to open in Access 2000 mode."
                   
            Access97 = False                ' So switch to Access2000 mode.
            adoTidy rs                      ' Tidy up failed connection attempt
            Err.Clear                       ' Clear the error state
            GoTo retry                      ' Go back and try with Access2000
        End If
        Moff                                ' None correctable error, User msg.
        MsgBox "Failed to Open " & MyDB & vbCrLf & vbCrLf & Error
        Err.Clear
        TableCount = 0
        FieldCount = 0
        TheConn = ""
    Else                                    ' If no errors after opening Schema

        Do While Not rs.EOF                 ' Loop thro' Schema recordset
            n = rs!TABLE_NAME               ' Get current Table name
            If Mid(n, 1, 4) <> "MSys" Then  ' If not an MS'System' table
                LB(0).AddItem n             ' Add the user Table to listbox
                TableCount = TableCount + 1 ' and bump Table counter.
            End If                          ' until done
            rs.MoveNext
        Loop
    End If
    
    If Err Then                             ' Report any errors
        Moff
        MsgBoxErr "Error in Gettables(" & Tbl & ")"
        lbl(0) = "No Tables"
        TableCount = 0
        TheConn = ""
    Else
        If TableCount > 0 Then                  ' If we have tables in DB
            GetTables = TableCount              ' then show Table count
            lbl(0) = CStr(TableCount) & " Table(s)"
            DoEvents
        Else
            Moff
            MsgBox "No User tables in database", , "No Tables found!"
            lbl(0) = "No Tables"
        End If
    End If
    
    SetBackColor LB(0)
    
    adoTidy rs                                  ' Disconnect from database
    adoTidy cn
    
End Function

Sub UpdateDelimitedList(Lst As String, Itm As String, delim As String, max As Integer, Optional Remove As Boolean = False)
    Dim arr, i%, j%, Tmp$, cnt%
    On Error Resume Next
    
    If Len(Itm) = 0 Then Exit Sub               ' Quit if no item
    
    If Not ItemExists(Lst, Itm, delim) Then     ' If the incoming Item is not in list
        Lst = Itm & delim & Lst                 ' append it to top (start) of list.
        If Mid(Lst, Len(Lst), 1) = delim Then   ' If this happens to be first item in
            Lst = Mid(Lst, 1, Len(Lst) - 1)     ' list, remove trailing delimiter
        End If
    Else
        If Remove Then                          ' If Remove=True then call the routine
            DeleteListItem Lst, Itm, delim      ' that will do that.
        End If
    End If

    arr = split(Lst, delim)
    cnt = UBound(arr) + 1                       ' Get number of items in list
    If cnt > 1 And Not Remove Then              ' If we have more than 1 item
        If cnt > max Then cnt = max             ' Reduce if max exceeded
        If Itm <> Tmp Then                      ' If the Item not top of list
            For i = 0 To cnt - 1                ' loop thro list until the
                If arr(i) = Itm Then            ' item if found, then move
                    For j = i To 1 Step -1      ' all entries above match
                        arr(j) = arr(j - 1)     ' down one position in the
                    Next                        ' list and put the reqd Item
                    arr(0) = Itm                ' it at top of list and we
                    Exit For                    ' are done, quit the loop
                End If
            Next
        End If
    End If
    
    Tmp = ""                                    ' re-assemble the array
    For i = 0 To cnt - 1                        ' back to a delimited list
        Tmp = Tmp & arr(i)
        If i < cnt - 1 Then Tmp = Tmp & delim
    Next
    Lst = Tmp
    
    Set arr = Nothing
    
End Sub

Function ItemExists(Lst As String, Itm As String, d As String) As Boolean

'   Function checks if Item (Itm) exists in a delimited (D) List (Lst)
'   of items. We append delimiters for check, to ensure unique matches.

    ItemExists = Sgn(Abs(InStr(d & Lst & d, d & Itm & d) > 0)) * -1

End Function


Sub SetWindowSplit(Source As Control, SplitY As Single)

'   To save coding every control for the dragdrop event, I just
'   get that event to call this routine
'
    If Source = frSplit Then                      ' If source is 'Split'
        If SplitY > Height - cbo.Height * 3 Then  ' ensure split stays
            SplitY = Height - cbo.Height * 3      ' within a sensible
        ElseIf SplitY < Height * 0.15 Then        ' range of values
            SplitY = Height * 0.15
        End If
        
        frSplit.Move 0, SplitY                    ' Position the Split
        SplitTop = SplitY                         ' Note the position
        Form_Resize                               ' Use resize event to repaint
    End If                                        ' display

End Sub

Sub SetToolTips(ToolTipsOn As Boolean)
    Dim i As Integer
    
    Mon
'
'   Load or Nullify Tools tips as requested by incoming var 'ToolTipsOn' (True/False)
'
    If ToolTipsOn Then
            
        LB(0).ToolTipText = " Tables in current Database "
        LB(1).ToolTipText = " Fields in current Table "
        txbCode.ToolTipText = " Output box - DblClick copies Output to clipboard and saves Codebox "
        txbHelp.ToolTipText = " Help file. (If edited, you will be prompted to save any changes) "
        cbo.ToolTipText = " Codebox - Enter your code line (DblClick to save entry in list) "
        txbS(18).ToolTipText = " Custom Connection string (see Create menu)"
        For i = 0 To lbRem.Count - 1
            lbRem(i).ToolTipText = " DblClick to Hide/Show reminder tips "
        Next
        lbl(0).ToolTipText = " Number of User Tables in current Database "
        lbl(1).ToolTipText = " Number of Fields in currently selected Table "
        lbl(2).ToolTipText = " Click to Check All fields for Sql creation "
        lbl(3).ToolTipText = " Click to Un-Check All Fields used in Sql creation "
        lbl(4).ToolTipText = " Creates an Sql INSERT string using Fields data "
        lbl(5).ToolTipText = " Creates an Sql UPDATE string using Fields data "
        lbl(6).ToolTipText = " Create Code from Template combo "
        
    Else
    
        LB(0).ToolTipText = ""
        LB(1).ToolTipText = ""
        txbCode.ToolTipText = ""
        txbHelp.ToolTipText = ""
        cbo.ToolTipText = ""
        txbS(18).ToolTipText = ""
        For i = 0 To lbRem.Count - 1
            lbRem(i).ToolTipText = ""
        Next
        For i = 0 To lbl.Count - 1
            lbl(i).ToolTipText = ""
        Next
        
    End If

    Moff
    
End Sub

Sub LoadMenusMRU()
    Dim arr, j%, i%
'
'   Apply MRU Filenames to File Menu (Can set Max in Declarations section)
'
    For i = 1 To mnMRU.Count - 1            ' Unload any unused
        Unload mnMRU(i)                     ' menus
    Next                                    ' load MRU menus
    j = 1
    arr = split(MRU, ",")                   ' Delimited string to array
    For i = 0 To UBound(arr)                ' Loop for the MRU filenames
        If Len(arr(i)) > 0 Then             ' Ensure we have a filename entry
            If j = 1 Then                   ' We have a filename
                Load mnMRU(j)               ' Load a new runtime menu
                mnMRU(j).Visible = True     ' make it visible, this one is
                mnMRU(j).Caption = "-"      ' a seperator bar, set its caption.
            End If                          ' Note can't use i, may be gaps
            j = j + 1                       ' in arr, so we use j for counter
            Load mnMRU(j)                   ' Load the rest of the MRU menus
            mnMRU(j).Visible = True         ' making them visibible and
            mnMRU(j).Caption = arr(i)       ' setting captions to filenames
        End If
        DoEvents
    Next

End Sub


Function CreateCode(Template As String) As String
    Dim Tmp As String, u%, s%, j%, a%, ln%, a1%, a2%, f$, b$, final$, cmp$
    Dim Ctr As Integer, IndexStep As Integer
'
'   Take user entered template (code/text), and scan it for any of the
'   user definable special chars (or 'keys'). Replace the keys found,
'   with the appropriate Field name, Index number or setup defined Sub/Function
'
    If Len(Template) = 0 Then Exit Function         ' If no Text, nothing to do, quit
'
'   First check the template for any conflicts characters
'
    
    If Not mnOptions(0).Checked Then                ' If 'Allow Apostrophe' not selected
        a = InStr(Template, "'")                    ' in Settings menu then strip leading
        If a > 0 Then Template = Left(Template, a - 1)  ' Apostrophes
    End If

    If Trim(Specs(17)) <> "" Then                   ' If 'Template' saving delimiter
        If InStr(Template, Specs(17)) > 0 Then      ' in use, warn and quit
            MsgBox "Template delimiter '" & Specs(17) & "' not allowed! (Change in Setttings)"
            Exit Function                           ' Can be changed to any character
        End If                                      ' you prefer, that does not conflict
    End If
'
'  REM - FldTyp() holds DataTypes 0=Text 1=Numeric 2=Currency 3=Bool 4=Date
'                 (and it got loaded when Fields were read into the Fields listbox)

    Ctr = Val(Specs(15))                        ' Preset Index (Field) counter, normally 0
    IndexStep = Val(Specs(16))                  ' Preset Index Step +/- n, normally 1
    For u = 0 To LB(1).ListCount - 1            ' Loop through fields (held in LB(1))
        f = FldNam(u)                           ' Get Field name (strip field num)
        b = Template                            ' Get a copy of the
        Do                                      ' template, and loop
            ln = Len(b)                         ' Note Len before Replacing
            If z(0) Then b = Replace(b, Specs(0), f)                ' Fieldname
            If z(1) Then b = Replace(b, Specs(1), CStr(Ctr))        ' Index
            If z(2) Then b = Replace(b, Specs(2), CStr(FldTyp(u)))  ' Type
            If z(3) Then b = Replace(b, Specs(3), CStr(CusTyp(u)))  ' Custom Type
            If z(4) Then b = Replace(b, Specs(4), CStr(DefSiz(u)))  ' DefinedSize
            
            Select Case CusTyp(u)               ' Now do the 'Type' dependent stuff
                Case 0  ' Text field
                    If z(5) And z(10) Then
                        b = Replace(b, Specs(5), Specs(10))
                    End If
                    
                Case 1  ' Number field
                    If z(6) And z(11) Then
                        b = Replace(b, Specs(6), Specs(11))
                    End If
                    
                Case 2  ' Currency field
                    If z(7) And z(11) Then
                        b = Replace(b, Specs(7), Specs(12))
                    End If
                    
                Case 3  ' Boolean field
                    If z(8) And z(13) Then
                        b = Replace(b, Specs(8), Specs(13))
                    End If
                    
                Case 4  ' User field
                    If z(9) And z(11) Then
                        b = Replace(b, Specs(9), Specs(14))
                    End If

            End Select
            
            If Len(b) = ln Then                 ' If Len not changed. 'might' be done
                If cmp <> b Then                ' so compare with comparison string.
                    cmp = b                     ' If not done, update comparison string
                Else                            ' else if comparison matched, then the
                    Exit Do                     ' replacements for this field/line
                End If                          ' are done - quit loop for next field.
            End If
        Loop
        cmp = ""                                ' Zero the comparison string
        Tmp = Tmp & b & vbCrLf                  ' Add the fragment just generated for this
        Ctr = Ctr + IndexStep                   ' Field to tmp and bump the Counter
    Next                                        ' and continue in loop for next field.
    
    Tmp = Tmp & vbCrLf                          ' Loop done, Append cr/lf on end of Tmp

    If mnOptions(2).Checked Then                ' If 'AutoCopy' selected in Options,
        Clipboard.Clear                         ' synchronize the clipboard with
        Clipboard.SetText Tmp                   ' the newly generated code.
    End If
    
    SaveCombo = True                            ' Indicate Outputbox is combo sourced
                                                ' code, so Template gets saved on DblClick
    CreateCode = Tmp                            ' Return the generated code

End Function



Function CreateFieldsList(TableName As String) As String
    Dim Tmp As String, u As Integer

'   Field List & data for Printable reference, gets called for each table

    Tmp = Tmp + "DataBase: " & MyDB & vbCrLf                ' Create a Heading
    Tmp = Tmp + "TableName: " & TableName & vbCrLf & vbCrLf ' Using DB & Table Name
    For u = 0 To LB(1).ListCount - 1                        ' Loop fields in Table
        Tmp = Tmp & FormatFieldItem(u) & vbCrLf             ' Format entries
    Next
    CreateFieldsList = Tmp & vbCrLf                         ' Return the list
    
End Function

Function GetTableName() As String
    Dim u As Integer

    With LB(0)
        If .ListCount > 0 Then                      ' If Items exist, and one
            If .ListIndex >= 0 Then                 ' is selected by ListIndex
                GetTableName = .List(.ListIndex)    ' return the Item
            Else                                    ' else
                For u = 0 To .ListCount - 1         ' Scan the Listbox for a
                    If .Selected(u) Then            ' 'Selected' Item and
                        LB(0).ListIndex = u         ' Force listindex to selected entry
                        GetTableName = .List(u)     ' return first one found.
                        Exit For                    ' Quit loop
                    End If
                Next
            End If
        End If
    End With
 
End Function

Private Sub cbo_Change()

    If Len(cbo.Text) > 0 Then           ' Backcolor to grey if empty
        cbo.BackColor = vbWhite
        If mnOptions(1).Checked Then
            mnCreates_Click 3
        End If
    Else
        cbo.Tag = ""
        cbo.BackColor = vbGrey
    End If
    
    If cbo.Text = "" Then txbCode = ""              ' Clear any remnants
    
    mnCreates(3).Enabled = Sgn(Len(cbo.Text)) * -1 _
                        And FieldCount > 0 _
                        And Not mnOptions(1).Checked
    
End Sub



Private Sub cbo_DblClick()

    SaveNewComboItem

End Sub

Private Sub cbo_DragDrop(Source As Control, X As Single, Y As Single)

    SetWindowSplit Source, Y + frPane.Top + cbo.Top
    
End Sub

Private Sub cbo_GotFocus()
    cbo.SelLength = 0
End Sub

Private Sub lbl_Click(Index As Integer)
    Dim oc&
    
    If Index > 1 Then
        oc = lbl(Index).BackColor
        lbl(Index).BackColor = vbRed
        Wait 0.05
        lbl(Index).BackColor = oc
    End If
    
    Select Case Index
    
    Case 2  ' Set listbox to 'All items Checked'
        SetChecked LB(1), True
    Case 3  ' Set listbox to 'All items Un-Checked'
        SetChecked LB(1), False
    Case 4  ' Create an Sql 'Insert' string
        mnCreates_Click 9
    Case 5  ' Create an Sql 'Update' string
        mnCreates_Click 10
    Case 6  ' Create Code from Combo Template entry
        If mnCreates(3).Enabled Then mnCreates_Click 3
    End Select
        
End Sub

Private Sub ListBox_ItemCheck(Index As Integer, Item As Integer)

    Selects(LB(0).ListIndex, Item) = LB(1).Selected(Item)
    
End Sub

Private Sub mnCancel_Click()
    Dim i As Integer
    
    For i = 0 To 18
        txbS(i) = txbS(i).Tag        ' Restore previous settings frfom tag properties
    Next
    
    SetMenus 0                       ' Restore user menus
    
End Sub

Private Sub lbRem_Click(Index As Integer)
    
    If Index < 10 Then
        If lbRem(0).ForeColor <> vbGrey Then
            Remind Index
        End If
    End If
    
End Sub

Private Sub lbRem_DblClick(Index As Integer)
    Dim i As Integer
'
'   If any 'Hidden' (forecolor=backcolor) label gets clicked then
'   wenneed to Display or Hide (toggle) the current Reminder bar
'
    If lbRem(0).ForeColor = vbGrey Then             ' If Forecolor vbGrey, bar is
        For i = 0 To 10                             ' hidden, so loop thro
            If i = RemIndex Or i = 10 Then          ' Restore 'Selected' and 'Tip'
                lbRem(i).ForeColor = vbBlue         ' forecolors to Blue
            Else
                lbRem(i).ForeColor = vbBlack        ' and rest to vbBlack
            End If
        Next
        Remind RemIndex                             ' Restore Hilight by refreshing
        
    Else                                            ' else we need to Hide bar
    
        For i = 0 To 10                             ' making forecolor same as form
            lbRem(i).ForeColor = vbGrey             ' We don't use 'visible' property
        Next                                        ' cos' we need to keeps labels
    End If                                          ' 'live' to sense DblClicks.
    
End Sub


Private Sub mnCloseHelp_Click()
    Dim i As Integer

    If frSetup.Visible Then                             ' If we are closing setup mode
    
        Mon                                             ' Show Hourglass
        For i = 0 To 18
            DoIni "Specs" & CStr(i), txbS(i), True      ' Save Setup textbox values
            Specs(i) = txbS(i)                          ' Update the Setups Array
            z(i) = Sgn(Len(Specs(i))) * -1              ' Saves time and simpler code later
            txbS(i).Tag = ""                            ' Clear tags (Well, why not)
        Next
        LoadSpecs                                       ' rebuild Setups array (may changed)
        Remind 0                                        ' rebuild Reminder List
        SetMenus 0                                      ' Restore user menus
        Moff                                            ' Kill Hourglass
    ElseIf frHelp.Visible Then
    
        If txbHelp <> txbHelp.Tag Then
            If MsgBox("Your have changed text - Ok to quit without saving", vbOKCancel, " Changes Occurred") = vbCancel Then Exit Sub
        End If
    End If
    
    SetMenus 0                                          ' Restore User menus
    
End Sub

Sub mnCreate_Click()

    mnCreates(0).Enabled = (FieldCount > 0)
    mnCreates(1).Enabled = (FieldCount > 1)
    mnCreates(3).Enabled = (FieldCount > 0) And (Len(cbo.Text) > 0)
    
End Sub

Private Sub mnFile_Click()

    mnCloseDatabase.Enabled = Sgn(Len(MyDB)) * -1

End Sub

Private Sub mnOption_Click()

    mnOptions(9).Enabled = (cbo.ListIndex >= 0)
    
End Sub

Private Sub mnPrint_Click()

    mnPrintSetup(0).Caption = "Left Margin  [" & prtLeft & "]"      ' Add settings to
    mnPrintSetup(1).Caption = "Top Margin   [" & prtTop & "]"       ' menu for user to
    mnPrintSetup(2).Caption = "Lines/Page   [" & prtLines & "]"     ' see immediately.
    
End Sub

Private Sub mnPrints_Click(Index As Integer)
    Dim f As String, s, i As Integer, arr, lc&
'
'   Primitive printer routine - can be removed or improved if printing important
'
    If Index = 2 And Len(txbCode) > 0 Then              ' If something to print
    
        If MsgBox(" Ok to Print?  ", vbOKCancel, " Print Code") = vbCancel Then Exit Sub
        
        f = Printer.FontName                            ' note current printer
        s = Printer.FontSize                            ' font name and size
        Printer.FontName = "Courier New"                ' Set to courier
        
        For i = 1 To Val(prtTop)                        ' Print any top margin
            Printer.Print                               ' Bump line counter
            lc = lc + 1
        Next
        
        arr = split(txbCode, vbCrLf)                    ' Break text box into lines
        
        For i = LBound(arr) To UBound(arr)              ' Loop thro lines, and Add
            Printer.Print Space(Val(prtLeft)) & arr(i)  ' left margin as we print
            lc = lc + 1                                 ' Bump line counter
            If Val(prtLines) > 0 Then                   ' If settings limit lines/page
                If lc > Val(prtLines) Then              ' and we reached limit then
                    lc = 0                              ' reset counter and
                    Printer.NewPage                     ' throw a form feed.
                End If
            End If
        Next
        
        Printer.FontName = f                            ' Restore Printer font
        Printer.FontSize = s                            ' and size
        Printer.EndDoc                                  ' Close down printing
        
        Set arr = Nothing
        
    End If
    
End Sub

Private Sub mnPrintSetup_Click(Index As Integer)
    Dim m$, b$

'   We want to keep printing simple, its only 'useful' for this sort of program,
'   so we use Choose and one InputBox for requesting changes to for all settings.

    m = Choose(Index + 1, "Left Margin (0-20)", "Top Margin (0-20)", "Lines per Page" & vbCrLf & "(0=No Limit)")
    b = Choose(Index + 1, prtLeft, prtTop, prtLines)
    
    Do
        b = Trim(InputBox(m, " Printer Setup", b))
        If b = "" Then Exit Sub
        If Val(b) >= 0 Then
            If Val(b) < Choose(Index + 1, 20, 20, 1000) Then
                Exit Do
            Else
                MsgBox "Out of range!"
            End If
        End If
    Loop
    
'   Apply the Value from settings to the appropriate Var and save in .ini file
    
    Select Case Index
        Case 0
            prtLeft = b
            DoIni "PrintLeft", prtLeft, True
        Case 1
            prtTop = b
            DoIni "PrintTop", prtTop, True
        Case 2
            prtLines = b
            DoIni "PrintLines", prtLines, True
    End Select
    
End Sub

Private Sub mnSaveHelp_Click()
    Dim b$
    
    txbHelp.Tag = txbHelp                       ' Store updated text in tag
    txbHelp = Replace(txbHelp, vbCrLf, "||")    ' Strip cr/lf's to save in ini file
    DoIni "Help", txbHelp, True                 ' Save txbHelp to ini file
    txbHelp = txbHelp.Tag                       ' Restore txbHelp.Text

End Sub

Private Sub mnSetup_Click()
    Dim i As Integer
    
    Mon                             ' Hourglass
    For i = 0 To 18
        txbS(i).Tag = txbS(i)       ' Save current values in tag property
    Next                            ' in case we want to cancel setup
    
    SetMenus 2                      ' Show Setup Panel
    Moff                            ' Kill Hourglass
    
End Sub

Private Sub frPane_DragDrop(Source As Control, X As Single, Y As Single)
SetWindowSplit Source, Y + frPane.Top
End Sub


Private Sub ListBox_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

    SetWindowSplit Source, Y + LB(0).Top
    
End Sub

Private Sub mnCloseDatabase_Click()

    Mon
    ClearAll
    Moff
    
End Sub

Private Sub mnHelps_Click(Index As Integer)
    Dim b As String
    
    Select Case Index
    
        Case 0      ' Help was selected
        
            SetMenus 3
            DoIni "Help", b, False              ' Get the help text
            b = Replace(b, "||", vbCrLf)        ' Restore the crlf's
            txbHelp.Text = b                    ' Assign it to textbox
            txbHelp.Tag = b                     ' Save copy in tag property
            txbHelp_Change
            
        Case 2      ' 'About' was selected
            MsgBox "ViewDB Version 1.00" & vbCrLf & vbCrLf & _
                "by Ken Ashton" & vbCrLf & vbCrLf & _
                "Email: ken@eurosol.com", , " About ViewDB"

    End Select
    
End Sub


Private Sub mnOpenDatabase_Click()
Dim u%, j%

    On Error Resume Next
    
    Mon
    ClearAll
    If SkipDialog = True Then                       ' If we already have a filename
        SkipDialog = False                          ' from MRU, reset flag now and
        If Dir$(MyDB) = "" Then                     ' Make sure the MRU listed file
            Moff                                    ' still on disk, if not
            MsgBox "File " + MyDB + " not Found!"   ' then warn user and update the
            Mon                                     ' MRU by removing file from list
            UpdateDelimitedList MRU, MyDB, ",", MXU, True
            LoadMenusMRU                            ' Reload modified MRU and Menus
            SetMenus 0
            Exit Sub
        End If
        
    Else
    
        LoadSaveSelects "", False                   ' DB will change, reset LB Selects
        MyDB = GetMdbFileName(LastPathOpen)         ' Get Database Filename to Open
        If MyDB = "" Then                           ' If no name, was user cancelled
            SetMenus 0                              ' so return menus to user and
            Exit Sub                                ' do nothing
        End If
    End If
    
    Mon
    
    TableCount = GetTables(MyDB, "", Access97)      ' Attempt to read Table structure

    If TableCount > 0 Then
        LastPathOpen = ExtractPath(MyDB)
        If TableCount < LastTable + 1 Then          ' If the TableCount is not less
            LB(0).ListIndex = LastTable             ' than last Table index, update
        ElseIf LB(0).ListIndex < 0 Then
            LB(0).ListIndex = 0
        End If                                      ' to last Table index
        If FieldCount = 0 Then ListBox_Click 0      ' If event didn't get triggered
    End If                                          ' then force to refresh Fields List
    
    UpdateDelimitedList MRU, MyDB, ",", MXU         ' Save any Filename changes to MRU
    LoadMenusMRU                                    ' list and reload for any Menu change
    FormCaption                                     ' Update Database Name in Caption
    SetMenus 0                                      ' Return menus to user.

End Sub

Function ExtractPath(Path)
    
    If Len(Path) > 0 Then _
        If InStr(Path, "\") > 0 Then _
            ExtractPath = Left(Path, InStrRev(Path, "\") - 1)

End Function


Private Sub Form_Load()
    Dim b As String, i As Integer
    
    Mon                                                 ' Hourglass On
    Set LB(0) = ListBox(0)                              ' Set the two listboxes
    Set LB(1) = ListBox(1)                              ' as object variables
    SetMenus 1                                          ' Turn off user menus
    DoIniForm Me, False                                 ' Restore last Form position
    DoIni "MRU", MRU, False                             ' Get MRU (Most Recent Files) list
    DoIni "TPL", TPL, False                             ' Get saved Templates
    DoIni "PrintLeft", prtLeft, False                   ' Get Printer settings, Left
    DoIni "PrintTop", prtTop, False                     ' & top margins and
    DoIni "PrintLines", prtLines, False                 ' lines/page.
    DoIni "LastTable", b, False: LastTable = Val(b)
    DoIni "Apos", b, False: mnOptions(0).Checked = Val(b)       ' Restore Checkboxes
    DoIni "AutoCode", b, False: mnOptions(1).Checked = Val(b)   ' and menu settings
    DoIni "AutoCopy", b, False: mnOptions(2).Checked = Val(b)
    DoIni "AutoOpen", b, False: mnOptions(3).Checked = Val(b)
    DoIni "Selects", b, False: LoadSaveSelects b, False
    DoIni "MsSans", b, False: MsSans = Val(b)                   ' Font settings
    mnOptions(4).Checked = MsSans
    If MsSans Then SetFonts MsSans
    DoIni "ToolTips", b, False: mnOptions(5).Checked = Val(b)   ' ToolTips
    SetToolTips Val(b)
    DoIni "Access97", b, False: mnOptions(6).Checked = Val(b)   ' Databse type
    DoIni "IsOnTop", b, False: IsOnTop = Val(b)                 ' Form 'On Top'
    mnOptions(7).Checked = IsOnTop
    If IsOnTop Then SetOnTop Me, True
    For i = 0 To 18
        DoIni "Specs" & CStr(i), b, False               ' Read Setups
        txbS(i) = b
    Next
    QualifyDefaults                                     ' Qualidy in case any bad values
    LoadSpecs                                           ' Get setups into a string array
    DoIni "SplitPos", b, False: SplitTop = Val(b)       ' Get correct Screen Split postn.
    DoIni "LastPathOpen", LastPathOpen, False           ' Get last path of last open DB
    DoIni "LastPathConn", LastPathConn, False           ' Get last path for Dialog browse
    Remind 0                                            ' Build Reminder Info Label
    FormCaption                                         ' Apply a pre-open DB Caption
    LoadMenusMRU                                        ' Get MRU Menus sorted
    LoadCombo                                           ' Load saved list of templates
    DoIni "CboText", b, False: cbo.Text = b             ' Restore combo template to last
    Me.Show                                             ' Show Form, in case Autoload DB
    Me.Refresh                                          ' Try and get it all painted
    
    If mnOptions(3).Checked Then                        ' If 'AutoOpen' is selected then
     If mnMRU.Count > 1 Then                            ' and we have an MRU list then
      If Dir(mnMRU(2).Caption) <> "" Then               ' If the DB at the top of MRU list
        mnMRU_Click 2                                   ' exists - then open it
      End If
     End If
    End If
    
    SetMenus 0                                          ' Enable user Menus
    
End Sub

Sub DoIniForm(frm As Form, WriteMode As Boolean)
Dim s As String, arr, Keyname As String, ret As Long
    On Error Resume Next
    
'   Simply pass the Form Name and action required (Save if WriteMode=true)
'   Screen state, bad Form positions are corrected automatically
    
    With frm
    
        Keyname = .Name & "Position"   ' Create ini [Section], eg [Form1Position]
    
        If WriteMode Then                           ' If WRITING the settings
                                                    ' ensure we don't do it if
            If .WindowState = 0 Then                ' window minmised or maximised
                s = Trim(Str(.Left)) & ","
                s = s & Trim(Str(.Top)) & ","
                s = s & Trim(Str(.Width)) & ","
                s = s & Trim(Str(.Height))
                ret = WritePrivateProfileString("Default", Keyname, s, _
                        (LCase(App.Path & "\" & App.EXEName) & ".ini"))
            End If
            
        Else
    
            s = String(256, Chr(0))
            ret = GetPrivateProfileString("Default", Keyname, "", s, 256, _
                    (LCase(App.Path & "\" & App.EXEName) & ".ini"))
            s = Mid(s, 1, ret)
            If Len(s) = 0 Then s = "0,0,1000,1000"
            arr = split(s, ",")
            If UBound(arr) <> 3 Then
                s = "1000,1000,1000,1000"
                arr = split(s, ",")
            End If
            If Err Then
                MsgBoxErr "Error Reading in DoIniForm()"
            End If
            If .WindowState = 0 Then
                If Val(arr(0)) < 0 Then arr(0) = "0"        ' Make sure the window
                If Val(arr(1)) < 0 Then arr(1) = "0"        ' gets some sensible
                If Val(arr(2)) < 1000 Then arr(2) = "4000"  ' values (in case ini
                If Val(arr(3)) < 1000 Then arr(3) = "3000"  ' file was corrupt, etc)
                If .BorderStyle = 1 Then                    ' If its not a sizable
                    .Left = Val(arr(0))                     ' window, just assign
                    .Top = Val(arr(1))                      ' Left and Top
                Else                                        ' else assign the lot
                    .Move Val(arr(0)), Val(arr(1)), Val(arr(2)), Val(arr(3))
                End If
                DoEvents
            End If
        End If
        If Err Then
            MsgBoxErr "Error Reading in DoIniForm()" & vbCrLf & Error
        End If
    End With
    DoEvents

End Sub





Sub Mon()

    Screen.MousePointer = 11        ' HourGlass Mousepointer
    
End Sub

Sub Moff()

    Screen.MousePointer = 0     ' Normal Mousepointer
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i As Integer, b$
    
    If frHelp.Visible Then      ' Could be inadvertant shutdown, if
        Cancel = True           ' Help screen showing, just close it
        mnCloseHelp_Click
        Exit Sub                ' and quit unload event
    End If
    
    Mon
    DoIni "MRU", MRU, True      ' Save MRU (Most Recent Files) List
    DoIni "TPL", TPL, True      ' Save Last used Template
    DoIni "Apos", CStr(Sgn(mnOptions(0).Checked)), True
    DoIni "AutoCode", CStr(Sgn(mnOptions(1).Checked)), True
    DoIni "AutoCopy", CStr(Sgn(mnOptions(2).Checked)), True
    DoIni "AutoOpen", CStr(Sgn(mnOptions(3).Checked)), True
    DoIni "MsSans", CStr(Sgn(mnOptions(4).Checked)), True
    DoIni "ToolTips", CStr(Sgn(mnOptions(5).Checked)), True
    DoIni "Access97", CStr(Sgn(mnOptions(6).Checked)), True
    DoIni "IsOnTop", CStr(Sgn(mnOptions(7).Checked)), True
    DoIni "CboText", cbo.Text, True
    DoIni "LastPathOpen", LastPathOpen, True
    DoIni "LastPathConn", LastPathConn, True
    DoIni "SplitPos", CStr(SplitTop), True
    DoIni "LastTable", CStr(LB(0).ListIndex), True
    LoadSaveSelects b, True
    DoIni "Selects", b, True

    For i = 0 To 18
        DoIni "Specs" & CStr(i), txbS(i), True      ' Write Setups
    Next
    DoIniForm Me, True                      ' Restore last Form position
    If IsOnTop Then SetOnTop Me, False      ' Remove 'OnTop' state to avoid
    Set LB(0) = Nothing                     ' windows instability.
    Set LB(1) = Nothing
    Moff
    End
    
End Sub
Sub adoTidy(r As Object)

    On Error Resume Next

'
'   Close (gracefully?) any unused recordset or connection
'
    If r Is Nothing Then Exit Sub
    If r.State > 1 Then r.Close
    Set r = Nothing
    If Err Then MsgBoxErr "Error in adoTidy()"

End Sub


Private Sub Form_Resize()
    Dim w, h, bd, ww, i, lbHeight, split, frHeight, SplitHeight
    Dim bd6, bd7, bd8, bd9, sh
    
    On Error Resume Next

    If WindowState = 1 Then Exit Sub    ' Quit if minimised to avoid errors

'   Do some basics, assign common dimensions to vars
    If Me.Width < Screen.Width * 0.25 Then Me.Width = Screen.Width * 0.25
    If Me.Height < Screen.Height * 0.325 Then Me.Height = Screen.Height * 0.325
    bd = Screen.Height * 0.003 ' Set bd for an arbitrary 'border width'
    w = Width
    h = Height
    bd6 = bd * 6
    bd7 = bd * 7
    bd8 = bd * 8
    bd9 = bd * 9
    ww = w - bd * 3.4
    
'   I'm using a Frame with a DragIcon assigned for the 'user draggable' band
    If SplitHeight = 0 Then SplitHeight = bd * 2  ' Set 'draggable' thickness
    If SplitTop = 0 Then SplitTop = h * 0.4         ' Limit how high can split
    frSplit.Move 0, SplitTop, w, SplitHeight        ' Limit how low can split
    lbHeight = SplitTop - bd * 10
    
'   Listboxes and their Header labels
    LB(0).Move bd * 0.5, bd * 10, ww * 0.3 - bd, lbHeight - bd * 2
    LB(1).Move bd * 0.5 + ww * 0.3, bd * 10, ww * 0.7 - bd, lbHeight - bd * 2
    
    lbl(0).Move bd * 0.5, bd, ww * 0.3 - bd, bd * 7
    lbl(1).Move bd * 0.5 + ww * 0.3, bd, ww * 0.2 - bd, bd * 7
    
    lbl(2).Move bd + ww * 0.5, bd, ww * 0.1 - bd, bd7      ' These are used
    lbl(3).Move bd + ww * 0.6, bd, ww * 0.1 - bd, bd7     ' for buttons
    lbl(4).Move bd + ww * 0.7, bd, ww * 0.105 - bd, bd7
    lbl(5).Move bd + ww * 0.805, bd, ww * 0.12 - bd, bd7
    lbl(6).Move bd + ww * 0.925, bd, ww * 0.07 - bd, bd7
    
    
'   Sizable bottom pane
    frHeight = h - SplitTop - bd * 18
    frPane.Move 0, SplitTop, w, frHeight
    cbo.Move bd * 0.5, bd * 2, w - bd * 4
    For i = 0 To 10                                     ' Just set all the heights
        lbRem(i).Top = bd * 12                          ' the rest is done elsewhere
    Next

    txbCode.Move bd * 0.5, bd * 18, w - bd * 4, frHeight - bd * 19
    
'   Help Screen
    frHelp.Move 0, bd, w - bd * 3, h - bd * 21
    txbHelp.Move bd, bd6, w - bd6, h - bd * 28
    
'   Setup Panel

    frSetup.Move 0, 0, w - bd * 3.5, h - bd * 20
    sh = (h - bd * 20) * 0.07
    If sh < bd * 8.5 Then sh = bd * 8.5
    For i = 0 To 4
        lbs(i).Move bd * 4, sh * i + bd8
        txbS(i).Move bd * 26, sh * i + bd7, bd * 18
        lbs(i + 5).Move bd * 4, sh * (i + 5) + bd8
        txbS(i + 5).Move bd * 52, sh * (i + 5) + bd7
        txbS(i + 10).Move bd * 86, sh * (i + 5) + bd7
    Next
    lbs(10).Move w * 0.5 + bd * 40, sh * 5 + bd8
    lbs(11).Move w * 0.5 + bd * 43, sh * 6 + bd8
    lbs(12).Move w * 0.5 + bd * 43, sh * 7 + bd8
    txbS(15).Move w * 0.5 + bd * 66, sh * 5 + bd7, bd8
    txbS(16).Move w * 0.5 + bd * 66, sh * 6 + bd7, bd8
    txbS(17).Move w * 0.5 + bd * 66, sh * 7 + bd7, bd8
    txbS(18).Move bd * 3, sh * 9 + bd * 18, w - bd * 9, sh * 5.4 - bd * 22
    
End Sub
Private Sub ListBox_Click(Index As Integer)
    Dim Tbl As String
    
    On Error Resume Next
'
'   Index 0=Table Names, 1=Field Names
'
    If Inhibit Or Index > 0 Then Exit Sub ' Quit if events suppressed or not table listbox
    
    Mon
    Tbl = GetTableName                  ' Get the name of the Table (Tbl)
    If Len(Tbl) = 0 Then Exit Sub       ' Quit if no Tbl (impossible??, but safety first)
    ClearListBoxes 1                    ' Clear any prev 'Fields' in list box
    GetFields Tbl
    SetBackColor LB(0)
    If mnOptions(1).Checked Then        ' If AutoCode is selected in Options
        If Len(cbo.Text) > 0 Then       ' and we have Tempale code
            mnCreates_Click 3           ' then force a re-build
        End If
    End If
    Moff
    
End Sub
Private Sub mnCreates_Click(Index As Integer)
    Dim Tmp As String
    Dim Tbl As String
    Dim db As String
    Dim u As Integer
    Dim q As String
    Dim li As Integer
    q = Chr(34)
    
    SaveCombo = False
    
    Select Case Index
    
    Case 0  ' List the selected Table
    
        Tbl = GetTableName
        If Len(Tbl) = 0 Then
            Moff
            MsgBox "No Table selected!"             ' Nothing selected
        Else                                        '
            Tmp = CreateFieldsList(Tbl)
        End If
        

    Case 1  ' Lists All tables in DB
    
        If LB(0).ListCount = 0 Then
            Moff  '
            MsgBox "No Tables!"
            Exit Sub
        Else                                        '
            Mon                                     ' Hourglass on
            li = LB(0).ListIndex                    ' Note current Table
            For u = 0 To LB(0).ListCount - 1        ' Loop thro tables
                LB(0).ListIndex = u                 ' Force a LB_Click event
                Tbl = LB(0).List(u)                 ' Get the Table name
                If FieldCount > 0 Then              ' If table has fields
                    Tmp = Tmp & CreateFieldsList(Tbl) & vbCrLf
                End If                              ' and Add them to list                            '
            Next
            If LB(0).ListIndex <> li Then           ' If necessary, restore
                LB(0).ListIndex = li                ' the same table user
            End If                                  ' had selected before
        End If
        
    Case 3                                          ' Code using template
        If FieldCount > 0 And Len(cbo.Text) > 0 Then
            Tmp = CreateCode(cbo.Text)
        Else
            Exit Sub
        End If
        
    Case 5 ' Create a Access 97 Connection string
        If MsgBox("Insert Hard File Name", vbYesNo, " Access 97 Connection String") = vbYes Then
            Tmp = GetMdbFileName(LastPathConn)
            If Len(Tmp) = 0 Then Exit Sub
            LastPathConn = ExtractPath(Tmp)
            Tmp = q & "Provider=Microsoft.Jet.OLEDB.3,51;Data Source=" & _
                    Tmp & ";Persist Security Info=False" & q
        Else
            Tmp = q & "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=" & _
                    q & " & MyDB & " & q & ";Persist Security Info=False" & q
        End If
        Tmp = "cn = " & Tmp
        
    Case 6  ' Create a Access 2000 Connection String
        If MsgBox("Insert Hard File Name", vbYesNo, " Access 97 Connection String") = vbYes Then
            Tmp = GetMdbFileName(LastPathConn)
            If Len(Tmp) = 0 Then Exit Sub
            LastPathConn = ExtractPath(Tmp)
            Tmp = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                    Tmp & ";Persist Security Info=False" & q
        Else
            Tmp = q & "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & q & " & MyDB & " & q & ";Persist Security Info=False" & q
        End If
        Tmp = "cn = " & Tmp
        
    Case 7
        Tmp = Specs(8)
        
    Case 9, 10, 11 ' Create an Sql INSERT, Sql UPDATE or Sql SELECT
        Tmp = CreateSQL(Index - 9)
        If Len(Tmp) = 0 Then Exit Sub

End Select

    If Len(Tmp) > 0 Then                    ' If tmp has something in it
        txbCode = Tmp                       ' put it into the textbox
        If mnOptions(2).Checked Then        ' If 'AutoCopy' to clipboard
            Clipboard.Clear                 ' is selected then send tmp
            Clipboard.SetText Tmp           ' to the clipboard
        End If
    End If

Moff


End Sub

Private Sub mnMRU_Click(Index As Integer)
Dim MruCount As Integer

Select Case Index

    Case 0
        Unload Me
    Case 2 To (mnMRU.Count - 1)
        If Len(mnMRU(Index).Caption) > 0 Then   ' User clicked file menu
        SkipDialog = True                   ' MRU list. Get name from
        MyDB = mnMRU(Index).Caption             ' menu and open if
        mnOpenDatabase_Click
    End If
    
End Select
    
End Sub

Sub EnumerateDataType(u As Integer)
On Error Resume Next
Dim m As String, v As String
'
'   No need for speed, as this gets called during reading the database
'   so the overhead is unnoticeable - so lets keep the code simple.
'
'   FldTyp() is the MS Datatype number for the field.
'   DefSiz() is the defined size for the field.
'   u is the incoming Field index number.
'   CusTyp() groups the MS DataTypes nto 0=Text 1=Numeric 2=Currency 3=Bool 4=Date
'            (Can distinguish between Text and Memo later, if DefSiz > 255, its Memo
'
    Select Case FldTyp(u)
            
        Case 1
            TxtTyp(u) = "Y/N"
            CusTyp(u) = 3
            
        Case 2
            TxtTyp(u) = "Int"
            CusTyp(u) = 1
            
        Case 3
            TxtTyp(u) = "IntLong"
            CusTyp(u) = 1
            
        Case 4
            TxtTyp(u) = "Single"
            CusTyp(u) = 1
            
        Case 5
            TxtTyp(u) = "Double"
            CusTyp(u) = 1
            
        Case 6
            TxtTyp(u) = "Currency"
            CusTyp(u) = 2
            
        Case 7
            TxtTyp(u) = "DateTime"
            CusTyp(u) = 4
            
        Case 8
            TxtTyp(u) = "Unicode"
            CusTyp(u) = 0
            
        Case 9
            TxtTyp(u) = "IDesp"
            CusTyp(u) = 1
            
        Case 10
            TxtTyp(u) = "ErrCode"
            CusTyp(u) = 1
            
        Case 11
            TxtTyp(u) = "Boolean"
            CusTyp(u) = 3
            
        Case 12
            TxtTyp(u) = "Variant"
            CusTyp(u) = 0
            
        Case 16
            TxtTyp(u) = "IntSmall"
            CusTyp(u) = 1
            
        Case 17
            TxtTyp(u) = "IntSmall" '(Byte)
            CusTyp(u) = 1
            
        Case 20
            TxtTyp(u) = "IntLong"
            CusTyp(u) = 1
            
        Case 129
            TxtTyp(u) = "Text"
            CusTyp(u) = 0
            
        Case 133
            TxtTyp(u) = "DateTime"
            CusTyp(u) = 4
            
        Case 134
            TxtTyp(u) = "Time"
            CusTyp(u) = 4
            
        Case 135
            TxtTyp(u) = "DateStamp"
            CusTyp(u) = 4
            
        Case 200
            TxtTyp(u) = "Text"
            CusTyp(u) = 0
            
        Case 201
            TxtTyp(u) = "Memo"
            CusTyp(u) = 0
            DefSiz(u) = 256
            
        Case 202
            TxtTyp(u) = "Text"
            CusTyp(u) = 0
            
        Case 203
            TxtTyp(u) = "Memo"
            CusTyp(u) = 0
            DefSiz(u) = 256
           
        Case Else
            TxtTyp(u) = "???"
            CusTyp(u) = 0
            DefSiz(u) = 0
            
    End Select
    
    If Len(FldNam(u)) > TxtLens(0) Then             ' Note the Longest Fieldname
        TxtLens(0) = Len(FldNam(u))                 ' and keep in TxtLens(0)
    End If
    
    If Len(TxtTyp(u)) > TxtLens(1) Then             ' Longest 'Type' Description
        TxtLens(1) = Len(TxtTyp(u))                 ' and keep in TxtLens(1)
    End If

    
    If Err Then MsgBoxErr "Possibly just no records in table! (Field=" & CStr(u) & ")"

End Sub

Private Sub mnOptions_Click(Index As Integer)
    Dim s As String

    Select Case Index

    Case 0  ' Allow Apstrophe                                   ' Allow Apostrophe comments
        mnOptions(Index).Checked = Not mnOptions(Index).Checked
        
    Case 1  ' AutoCode  - As template typed                     ' Create code as user
        mnOptions(Index).Checked = Not mnOptions(Index).Checked ' types entries
        cbo_Change      ' Force update of button/menu enables
          
    Case 2  ' AutoCopy  - Code to clipboard                     ' Copy changes immediate
        mnOptions(Index).Checked = Not mnOptions(Index).Checked
    
    Case 3  ' AutoOpen  - last database                         ' Open last DB on startup
        mnOptions(Index).Checked = Not mnOptions(Index).Checked
        
    Case 4  ' MsSans - Default is Courier New (Mono spaced)     ' User selects Courier
        MsSans = Not MsSans                                     ' or MsSansSerif font
        mnOptions(4).Checked = MsSans
        SetFonts MsSans                                         ' Apply Font
        
    Case 5 ' Tool Tips - On or Off                              ' Give programmer option
        mnOptions(Index).Checked = Not mnOptions(Index).Checked ' of killing tooltips
        SetToolTips mnOptions(Index).Checked

    Case 6  ' Access 97 Default                                 ' 'Checked' property will
        mnOptions(6).Checked = Not mnOptions(6).Checked         ' be used to select
                                                                ' Connection String
    Case 7  ' Always On top
        mnOptions(7).Checked = Not mnOptions(7).Checked
        SetOnTop Me, mnOptions(7).Checked                       ' Call 'SetOnTop' API
        
    Case 9  ' Remove template from Template List
        UpdateDelimitedList TPL, cbo.Text, Specs(17), MXT, True ' Save to list
        LoadCombo                                               ' Reload combo
        cbo_Change                                              ' Force refresh

    End Select

End Sub

Private Sub txbCode_Change()

        If Len(txbCode) > 0 Then            ' Grey out if no text
            txbCode.BackColor = vbWhite
        Else
            txbCode.BackColor = vbGrey
            txbCode.Tag = ""
        End If
End Sub


Private Sub txbCode_DblClick()
        
'   In many instances, I love the ability to DblClick any text control
'   in order to place its contents on clipboard, this does that and gives
'   a quick 'Red Flash' for visual feedback that the copy has occurred

    DblClickCopy txbCode
    If SaveCombo Then SaveNewComboItem
    
End Sub


Sub DblClickCopy(ctrl As Control)
    Dim oc As Long
    
'   (1)  Copies any text/label control (incoming ctrl) content to clipboard,
'   (2)  Displays quick 'Red Flash' to user for visual feedback that copy happened.
    
    If Len(Trim(ctrl)) > 0 Then             ' IF the control has content
        oc = ctrl.BackColor                 ' save the current background
        ctrl.BackColor = vbRed              ' color and make background
        ctrl.Refresh                        ' red (refresh makes it immediate)
        DoEvents                            ' Give windows some time
        Clipboard.Clear                     ' We have content, so clear
        Clipboard.SetText ctrl              ' clipboard and copy data to it
        Wait 0.1                            ' Hang around a bit so user
        ctrl.BackColor = vbWhite            ' sees 1/10 sec red flash and
        ctrl.SelLength = 0                  ' then restore control's color etc.
        ctrl.Refresh
    End If
    
End Sub


Sub Wait(PauseInterval As Single)
    Dim EndTime As Single
    
    EndTime = Timer + PauseInterval         ' Add Interval to Time when sub called
    While Timer < EndTime                   ' wait in loop till end time reached
        DoEvents
    Wend
    
End Sub

Sub SetBackColor(Lbx As ListBox)

    If Lbx.ListCount > 0 Then               ' If listbox has items
        Lbx.BackColor = vbWhite             ' then White background
    Else                                    ' else turn it to Grey
        Lbx.BackColor = vbGrey              ' for visual feedback
    End If
    
End Sub

Private Sub txbCode_DragDrop(Source As Control, X As Single, Y As Single)
SetWindowSplit Source, Y + frPane.Top + txbCode.Top
End Sub


Private Sub txbHelp_Change()

    If Len(txbHelp) > 0 Then            ' Grey out if no text
        txbHelp.BackColor = vbWhite
    Else
        txbHelp.BackColor = vbGrey
    End If
    If frHelp.Visible Then
        mnSaveHelp.Visible = (txbHelp <> txbHelp.Tag)
    End If
End Sub

Private Sub txbHelp_DblClick()

'   In many instances, I love the ability to DblClick any text control
'   in order to place its contents on clipboard, this does that and gives
'   a quick 'Red Flash' for visual feedback that the copy has occurred

    DblClickCopy txbHelp
    

End Sub


Private Sub txbHelp_DragDrop(Source As Control, X As Single, Y As Single)
SetWindowSplit Source, Y + frPane.Top + txbHelp.Top
End Sub


Private Sub txbS_DblClick(Index As Integer)

    If Index = 18 Then DblClickCopy txbS(18)
    
End Sub


