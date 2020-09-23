VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Karen's Hasher"
   ClientHeight    =   6435
   ClientLeft      =   735
   ClientTop       =   75
   ClientWidth     =   10230
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   10230
   Begin VB.PictureBox picHash 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   4815
      Index           =   1
      Left            =   240
      ScaleHeight     =   4815
      ScaleWidth      =   9735
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   9735
      Begin VB.Label lblWelcomeVerify 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "• Verify previously saved hash values"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   840
         MouseIcon       =   "Main.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   147
         Top             =   3000
         Width           =   3900
      End
      Begin VB.Label lblWelcomeGroup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "• Compute one hash value for a group of files or folders"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   840
         MouseIcon       =   "Main.frx":0594
         MousePointer    =   99  'Custom
         TabIndex        =   146
         Top             =   2640
         Width           =   5685
      End
      Begin VB.Label lblWelcomeFile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "• Compute individual hash values of one or more files"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   840
         MouseIcon       =   "Main.frx":06E6
         MousePointer    =   99  'Custom
         TabIndex        =   145
         Top             =   2280
         Width           =   5490
      End
      Begin VB.Label lblWelcomeText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "• Compute hash value of text string"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   840
         MouseIcon       =   "Main.frx":0838
         MousePointer    =   99  'Custom
         TabIndex        =   144
         Top             =   1920
         Width           =   3555
      End
      Begin VB.Label lblWelcome 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to Karen's Hasher!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   8055
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "This program can compute, or check, the MD5 Hash of any text string, file, or group of files."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   9495
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblVersion"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   3360
         Width           =   675
      End
   End
   Begin VB.PictureBox picHash 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   4815
      Index           =   6
      Left            =   480
      ScaleHeight     =   4815
      ScaleWidth      =   9735
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   600
      Width           =   9735
      Begin VB.ListBox lstTest 
         Height          =   3705
         IntegralHeight  =   0   'False
         Left            =   6360
         TabIndex        =   120
         Top             =   600
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.CommandButton btnTest 
         Caption         =   "Test"
         Height          =   375
         Left            =   6360
         TabIndex        =   119
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cboHashFav 
         Height          =   315
         ItemData        =   "Main.frx":098A
         Left            =   240
         List            =   "Main.frx":098C
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Frame fraAssoc 
         Caption         =   "File Associations: "
         Height          =   2655
         Left            =   3240
         TabIndex        =   111
         Top             =   960
         Width           =   2055
         Begin VB.CheckBox chkExt 
            Caption         =   ".sha&512"
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   117
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CheckBox chkExt 
            Caption         =   ".sha&384"
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   116
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CheckBox chkExt 
            Caption         =   ".sha&256"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   115
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CheckBox chkExt 
            Caption         =   ".sha22&4"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   114
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CheckBox chkExt 
            Caption         =   ".sha&1"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   113
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox chkExt 
            Caption         =   ".&md5"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   112
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame fraAvail 
         Caption         =   "Available Algorithms: "
         Height          =   2655
         Left            =   240
         TabIndex        =   101
         Top             =   960
         Width           =   2055
         Begin VB.CheckBox chkAvail 
            Caption         =   "SHA-&512"
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   107
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CheckBox chkAvail 
            Caption         =   "SHA-&384"
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   106
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CheckBox chkAvail 
            Caption         =   "SHA-&256"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   105
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CheckBox chkAvail 
            Caption         =   "SHA-22&4"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   104
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CheckBox chkAvail 
            Caption         =   "SHA-&1"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   103
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox chkAvail 
            Caption         =   "&MD5"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   102
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Label lblAssoc 
         BackStyle       =   0  'Transparent
         Caption         =   "Automatically launch %%PROGRAM%% when file with one of these extensions is opened:"
         Height          =   855
         Left            =   3120
         TabIndex        =   148
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label lblSep2 
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         Height          =   4455
         Left            =   6000
         TabIndex        =   118
         Top             =   120
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Which Hash algorithms should be available when computing hash values?"
         Height          =   855
         Left            =   240
         TabIndex        =   100
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label lblSep1 
         BackColor       =   &H8000000D&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         Height          =   4455
         Left            =   2760
         TabIndex        =   110
         Top             =   120
         Width           =   75
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Preferred Algorithm: "
         Height          =   255
         Left            =   240
         TabIndex        =   108
         Top             =   3960
         Width           =   1815
      End
   End
   Begin VB.PictureBox picHash 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   4815
      Index           =   5
      Left            =   720
      ScaleHeight     =   4815
      ScaleWidth      =   9735
      TabIndex        =   128
      TabStop         =   0   'False
      Top             =   600
      Width           =   9735
      Begin VB.Frame fraVerifyHash 
         Caption         =   "Step 2: Verify Hash Values "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   0
         TabIndex        =   130
         Top             =   2760
         Width           =   9735
         Begin VB.PictureBox picVerifyHash 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   1575
            Left            =   120
            ScaleHeight     =   1575
            ScaleWidth      =   9495
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   240
            Width           =   9495
            Begin VB.CommandButton btnVerifyUpdate 
               Caption         =   "&Update Hash File"
               Enabled         =   0   'False
               Height          =   615
               Left            =   0
               TabIndex        =   143
               Top             =   960
               Width           =   2775
            End
            Begin VB.TextBox txtVerifyNewHash 
               Enabled         =   0   'False
               Height          =   375
               Left            =   2880
               Locked          =   -1  'True
               TabIndex        =   138
               Top             =   1200
               Width           =   6615
            End
            Begin VB.TextBox txtVerifyOldHash 
               Enabled         =   0   'False
               Height          =   375
               Left            =   2880
               Locked          =   -1  'True
               TabIndex        =   137
               Top             =   360
               Width           =   6615
            End
            Begin VB.CommandButton btnVerifyCalc 
               Caption         =   "&Verify Hashes"
               Enabled         =   0   'False
               Height          =   615
               Left            =   0
               TabIndex        =   136
               Top             =   120
               Width           =   2775
            End
            Begin VB.Label lblVerifyNewHash 
               BackStyle       =   0  'Transparent
               Caption         =   "&Current Hash Value of Group:"
               Enabled         =   0   'False
               Height          =   255
               Left            =   2880
               TabIndex        =   140
               Top             =   960
               Width           =   3375
            End
            Begin VB.Label lblVerifyOldHash 
               BackStyle       =   0  'Transparent
               Caption         =   "Original Hash Val&ue of Group:"
               Enabled         =   0   'False
               Height          =   255
               Left            =   2880
               TabIndex        =   139
               Top             =   120
               Width           =   3255
            End
         End
      End
      Begin VB.Frame fraVerifyIn 
         Caption         =   "Step 1: Open Previously Saved Hash File "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   0
         TabIndex        =   129
         Top             =   0
         Width           =   9735
         Begin VB.PictureBox picVerifyIn 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   2055
            Left            =   120
            ScaleHeight     =   2055
            ScaleWidth      =   9495
            TabIndex        =   131
            TabStop         =   0   'False
            Top             =   360
            Width           =   9495
            Begin VB.PictureBox picVerifyPrompt 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               HasDC           =   0   'False
               Height          =   1335
               Left            =   3240
               OLEDropMode     =   1  'Manual
               ScaleHeight     =   1335
               ScaleWidth      =   5895
               TabIndex        =   141
               TabStop         =   0   'False
               Top             =   360
               Width           =   5895
               Begin VB.Label lblVerifyPrompt 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Drag and Drop previously saved hash file here,\nor click ""Open Hash File""."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1035
                  Left            =   120
                  OLEDropMode     =   1  'Manual
                  TabIndex        =   142
                  Top             =   120
                  Width           =   5520
               End
            End
            Begin VB.CommandButton btnVerifyOpen 
               Caption         =   "&Open Hash File"
               Height          =   855
               Left            =   0
               TabIndex        =   132
               Top             =   0
               Width           =   2775
            End
            Begin MSComctlLib.ListView lvVerifyIn 
               Height          =   1935
               Left            =   2880
               TabIndex        =   133
               Top             =   0
               Width           =   6615
               _ExtentX        =   11668
               _ExtentY        =   3413
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Object.Width           =   2540
               EndProperty
            End
            Begin VB.Label lblVerifyResult 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Verification\nFailed"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   0
               TabIndex        =   134
               Top             =   1080
               Visible         =   0   'False
               Width           =   2775
            End
         End
      End
   End
   Begin VB.PictureBox picHash 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   4815
      Index           =   4
      Left            =   240
      ScaleHeight     =   4815
      ScaleWidth      =   9735
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   600
      Width           =   9735
      Begin VB.Frame fraGroupCalc 
         Caption         =   "Step 2: Select Type of Hash, then Click ""Compute Hash"" button "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   0
         TabIndex        =   77
         Top             =   2760
         Width           =   9735
         Begin VB.PictureBox picGroupCalc 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   615
            Left            =   3360
            ScaleHeight     =   615
            ScaleWidth      =   6015
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   960
            Width           =   6015
            Begin VB.Label lblGroupCalc 
               BackStyle       =   0  'Transparent
               Caption         =   "Select a hashing algorithm, using the tabs shown above,\nthen click the ""Compute Hash"" button."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   120
               OLEDropMode     =   1  'Manual
               TabIndex        =   98
               Tag             =   "6010"
               Top             =   0
               Width           =   5640
            End
         End
         Begin VB.PictureBox picFilesHash 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   855
            Index           =   6
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   9255
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   840
            Width           =   9255
            Begin VB.CommandButton btnFilesCalc 
               Caption         =   "&Compute SHA-512 Hash"
               Height          =   825
               Index           =   6
               Left            =   0
               TabIndex        =   95
               ToolTipText     =   "Compute SHA-512 hash of group"
               Top             =   0
               Width           =   2655
            End
            Begin VB.TextBox txtFilesCalc 
               Height          =   825
               Index           =   6
               Left            =   2760
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   96
               Top             =   0
               Width           =   6480
            End
         End
         Begin VB.PictureBox picFilesHash 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   855
            Index           =   5
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   9255
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   840
            Width           =   9255
            Begin VB.CommandButton btnFilesCalc 
               Caption         =   "&Compute SHA-384 Hash"
               Height          =   825
               Index           =   5
               Left            =   0
               TabIndex        =   92
               ToolTipText     =   "Compute SHA-384 hash of group"
               Top             =   0
               Width           =   2655
            End
            Begin VB.TextBox txtFilesCalc 
               Height          =   825
               Index           =   5
               Left            =   2760
               TabIndex        =   93
               Top             =   0
               Width           =   6480
            End
         End
         Begin VB.PictureBox picFilesHash 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   855
            Index           =   4
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   9255
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   840
            Width           =   9255
            Begin VB.TextBox txtFilesCalc 
               Height          =   825
               Index           =   4
               Left            =   2760
               MultiLine       =   -1  'True
               TabIndex        =   90
               Top             =   0
               Width           =   6480
            End
            Begin VB.CommandButton btnFilesCalc 
               Caption         =   "&Compute SHA-256 Hash"
               Height          =   825
               Index           =   4
               Left            =   0
               TabIndex        =   89
               ToolTipText     =   "Compute SHA-256 hash of group"
               Top             =   0
               Width           =   2655
            End
         End
         Begin VB.PictureBox picFilesHash 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   855
            Index           =   3
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   9255
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   840
            Width           =   9255
            Begin VB.TextBox txtFilesCalc 
               Height          =   825
               Index           =   3
               Left            =   2760
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   87
               Top             =   0
               Width           =   6480
            End
            Begin VB.CommandButton btnFilesCalc 
               Caption         =   "&Compute SHA-224 Hash"
               Height          =   825
               Index           =   3
               Left            =   0
               TabIndex        =   86
               ToolTipText     =   "Compute SHA-224 hash of group"
               Top             =   0
               Width           =   2655
            End
         End
         Begin VB.PictureBox picFilesHash 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   855
            Index           =   2
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   9255
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   840
            Width           =   9255
            Begin VB.CommandButton btnFilesCalc 
               Caption         =   "&Compute SHA-1 Hash"
               Height          =   825
               Index           =   2
               Left            =   0
               TabIndex        =   83
               ToolTipText     =   "Compute SHA1 Hash of group"
               Top             =   0
               Width           =   2655
            End
            Begin VB.TextBox txtFilesCalc 
               Height          =   825
               Index           =   2
               Left            =   2760
               TabIndex        =   84
               Top             =   0
               Width           =   6480
            End
         End
         Begin VB.PictureBox picFilesHash 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   855
            Index           =   1
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   9255
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   840
            Width           =   9255
            Begin VB.CommandButton btnFilesCalc 
               Caption         =   "&Compute MD5 Hash"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   825
               Index           =   1
               Left            =   0
               TabIndex        =   80
               ToolTipText     =   "Compute MD5 Hash of group"
               Top             =   0
               Width           =   2655
            End
            Begin VB.TextBox txtFilesCalc 
               Height          =   825
               Index           =   1
               Left            =   2760
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   81
               Top             =   0
               Width           =   6480
            End
         End
         Begin MSComctlLib.TabStrip tabGroup 
            Height          =   1455
            Left            =   120
            TabIndex        =   78
            Top             =   360
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   2566
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   6
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   " &MD5 "
                  Object.Tag             =   "1"
                  Object.ToolTipText     =   "Compute MD5 Hash of Group"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   " SHA-&1 "
                  Object.Tag             =   "2"
                  Object.ToolTipText     =   "Compute SHA-1 Hash of Group"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   " SHA-22&4 "
                  Object.Tag             =   "3"
                  Object.ToolTipText     =   "Compute SHA-224 Hash of Group"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   " SHA-&256 "
                  Object.Tag             =   "4"
                  Object.ToolTipText     =   "Compute SHA-256 Hash of Group"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   " SHA-&384 "
                  Object.Tag             =   "5"
                  Object.ToolTipText     =   "Compute SHA-384 Hash of Group"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   " SHA-&512 "
                  Object.Tag             =   "6"
                  Object.ToolTipText     =   "Compute SHA-512 Hash of Group"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame fraGroupIn 
         Caption         =   "Step 1: Select Files of Group to be Hashed: 0 files "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   0
         TabIndex        =   68
         Top             =   0
         Width           =   9735
         Begin VB.PictureBox picGroupIn 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   2055
            Left            =   120
            ScaleHeight     =   2055
            ScaleWidth      =   9495
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   360
            Width           =   9495
            Begin VB.CommandButton btnGroupFolder 
               Caption         =   "&Add Folder to Group ..."
               Height          =   375
               Left            =   0
               TabIndex        =   71
               ToolTipText     =   "Add a folder to the group"
               Top             =   480
               Width           =   2775
            End
            Begin VB.PictureBox picFilesPrompt 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               HasDC           =   0   'False
               Height          =   1335
               Left            =   3240
               OLEDropMode     =   1  'Manual
               ScaleHeight     =   1335
               ScaleWidth      =   5895
               TabIndex        =   75
               TabStop         =   0   'False
               Top             =   360
               Width           =   5895
               Begin VB.Label lblFilesPrompt 
                  BackStyle       =   0  'Transparent
                  Caption         =   $"Main.frx":098E
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1035
                  Left            =   120
                  OLEDropMode     =   1  'Manual
                  TabIndex        =   76
                  Top             =   120
                  Width           =   5520
               End
            End
            Begin VB.CommandButton btnGroupBrowse 
               Caption         =   "&Add File(s) to Group ..."
               Height          =   375
               Left            =   0
               TabIndex        =   70
               ToolTipText     =   "Add one or more files to the group"
               Top             =   0
               Width           =   2775
            End
            Begin VB.CommandButton btnGroupDelAll 
               Caption         =   "Remove &All Files && Folders"
               Enabled         =   0   'False
               Height          =   375
               Left            =   0
               TabIndex        =   73
               ToolTipText     =   "Remove all files and folders from group"
               Top             =   1560
               Width           =   2775
            End
            Begin VB.CommandButton btnGroupDelSel 
               Caption         =   "&Remove Selected Items"
               Enabled         =   0   'False
               Height          =   375
               Left            =   0
               TabIndex        =   72
               ToolTipText     =   "Remove selected files and folders from group"
               Top             =   1080
               Width           =   2775
            End
            Begin MSComctlLib.ListView lvGroupIn 
               Height          =   1965
               Left            =   2880
               TabIndex        =   74
               Top             =   0
               Width           =   6615
               _ExtentX        =   11668
               _ExtentY        =   3466
               View            =   3
               LabelEdit       =   1
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               HideColumnHeaders=   -1  'True
               OLEDropMode     =   1
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               Appearance      =   1
               OLEDropMode     =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "File"
                  Object.Width           =   2540
               EndProperty
            End
         End
      End
   End
   Begin VB.PictureBox picHash 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   4815
      Index           =   3
      Left            =   240
      ScaleHeight     =   4815
      ScaleWidth      =   9735
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   600
      Width           =   9735
      Begin VB.Frame fraFileCalc 
         Caption         =   "Step 2: Select Type of Hash, then Click ""Compute Hash of Each File"" button "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   0
         TabIndex        =   45
         Top             =   2760
         Width           =   9735
         Begin VB.PictureBox picFileHash 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   855
            Index           =   1
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   9255
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   840
            Width           =   9255
            Begin VB.CommandButton btnFileCalcAll 
               Caption         =   "&Compute MD5 Hashes"
               Enabled         =   0   'False
               Height          =   825
               Index           =   1
               Left            =   0
               TabIndex        =   48
               ToolTipText     =   "Compute the MD5 Hash of Each File"
               Top             =   0
               Width           =   2655
            End
            Begin VB.CommandButton btnFileCalcSel 
               Caption         =   "Com&pute MD5 Hash of Each Selected File"
               Enabled         =   0   'False
               Height          =   375
               Index           =   1
               Left            =   0
               TabIndex        =   49
               ToolTipText     =   "Compute MD5 Hash of Each Selected File"
               Top             =   480
               Visible         =   0   'False
               Width           =   2655
            End
         End
         Begin VB.PictureBox picFileCalc 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   735
            Left            =   3360
            ScaleHeight     =   735
            ScaleWidth      =   5295
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   960
            Width           =   5295
            Begin VB.Label lblFileCalc 
               BackStyle       =   0  'Transparent
               Caption         =   "Select a hashing algorithm, using the tabs shown above, then click the ""Compute Hash of Each File"" button."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   120
               OLEDropMode     =   1  'Manual
               TabIndex        =   66
               Tag             =   "6010"
               Top             =   0
               Width           =   5145
            End
         End
         Begin VB.PictureBox picFileHash 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   855
            Index           =   6
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   9255
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   840
            Width           =   9255
            Begin VB.CommandButton btnFileCalcSel 
               Caption         =   "Com&pute SHA-512 Hash of Each Selected File"
               Enabled         =   0   'False
               Height          =   375
               Index           =   6
               Left            =   0
               TabIndex        =   64
               ToolTipText     =   "Compute SHA-512 Hash of Each Selected File"
               Top             =   480
               Visible         =   0   'False
               Width           =   2655
            End
            Begin VB.CommandButton btnFileCalcAll 
               Caption         =   "&Compute SHA-512 Hashes"
               Enabled         =   0   'False
               Height          =   825
               Index           =   6
               Left            =   0
               TabIndex        =   63
               ToolTipText     =   "Compute the SHA-512 Hash of Each File"
               Top             =   0
               Width           =   2655
            End
         End
         Begin VB.PictureBox picFileHash 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   855
            Index           =   5
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   9255
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   840
            Width           =   9255
            Begin VB.CommandButton btnFileCalcSel 
               Caption         =   "&Compute SHA-384 Hash of Each Selected File"
               Enabled         =   0   'False
               Height          =   375
               Index           =   5
               Left            =   0
               TabIndex        =   61
               ToolTipText     =   "Compute SHA-384 Hash of Each Selected File"
               Top             =   480
               Visible         =   0   'False
               Width           =   2655
            End
            Begin VB.CommandButton btnFileCalcAll 
               Caption         =   "&Compute SHA-384 Hashes"
               Enabled         =   0   'False
               Height          =   825
               Index           =   5
               Left            =   0
               TabIndex        =   60
               ToolTipText     =   "Compute the SHA-384 Hash of Each File"
               Top             =   0
               Width           =   2655
            End
         End
         Begin VB.PictureBox picFileHash 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   855
            Index           =   4
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   9255
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   840
            Width           =   9255
            Begin VB.CommandButton btnFileCalcAll 
               Caption         =   "&Compute SHA-256 Hashes"
               Enabled         =   0   'False
               Height          =   825
               Index           =   4
               Left            =   0
               TabIndex        =   57
               ToolTipText     =   "Compute the SHA-256 Hash of Each File"
               Top             =   0
               Width           =   2655
            End
            Begin VB.CommandButton btnFileCalcSel 
               Caption         =   "&Compute SHA-256 Hash of Each Selected File"
               Enabled         =   0   'False
               Height          =   375
               Index           =   4
               Left            =   0
               TabIndex        =   58
               ToolTipText     =   "Compute SHA-256 Hash of Each Selected File"
               Top             =   480
               Visible         =   0   'False
               Width           =   2655
            End
         End
         Begin VB.PictureBox picFileHash 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   855
            Index           =   3
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   9255
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   840
            Width           =   9255
            Begin VB.CommandButton btnFileCalcSel 
               Caption         =   "&Compute SHA-224 Hash of Each Selected File"
               Enabled         =   0   'False
               Height          =   375
               Index           =   3
               Left            =   0
               TabIndex        =   55
               ToolTipText     =   "Compute SHA-224 Hash of Each Selected File"
               Top             =   480
               Visible         =   0   'False
               Width           =   2655
            End
            Begin VB.CommandButton btnFileCalcAll 
               Caption         =   "&Compute SHA-224 Hashes"
               Enabled         =   0   'False
               Height          =   825
               Index           =   3
               Left            =   0
               TabIndex        =   54
               ToolTipText     =   "Compute the SHA-224 Hash of Each File"
               Top             =   0
               Width           =   2655
            End
         End
         Begin VB.PictureBox picFileHash 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   855
            Index           =   2
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   9255
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   840
            Width           =   9255
            Begin VB.CommandButton btnFileCalcAll 
               Caption         =   "&Compute SHA-1 Hashes"
               Enabled         =   0   'False
               Height          =   825
               Index           =   2
               Left            =   0
               TabIndex        =   51
               ToolTipText     =   "Compute the SHA-1 Hash of Each File"
               Top             =   0
               Width           =   2655
            End
            Begin VB.CommandButton btnFileCalcSel 
               Caption         =   "&Compute SHA-1 Hash of Each Selected File"
               Enabled         =   0   'False
               Height          =   375
               Index           =   2
               Left            =   0
               TabIndex        =   52
               ToolTipText     =   "Compute SHA-1 Hash of each Selected File"
               Top             =   480
               Visible         =   0   'False
               Width           =   2655
            End
         End
         Begin MSComctlLib.TabStrip tabFile 
            Height          =   1455
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   2566
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   6
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   " &MD5 "
                  Object.Tag             =   "1"
                  Object.ToolTipText     =   "Compute MD5 Hash of Group"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   " SHA-&1 "
                  Object.Tag             =   "2"
                  Object.ToolTipText     =   "Compute SHA-1 Hash of Group"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   " SHA-22&4 "
                  Object.Tag             =   "3"
                  Object.ToolTipText     =   "Compute SHA-224 Hash of Group"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   " SHA-&256 "
                  Object.Tag             =   "4"
                  Object.ToolTipText     =   "Compute SHA-256 Hash of Group"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   " SHA-&384 "
                  Object.Tag             =   "5"
                  Object.ToolTipText     =   "Compute SHA-384 Hash of Group"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   " SHA-&512 "
                  Object.Tag             =   "6"
                  Object.ToolTipText     =   "Compute SHA-512 Hash of Group"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame fraFileIn 
         Caption         =   "Step 1: Select Individual Files to be Hashed: 0 files "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   9735
         Begin VB.PictureBox picFileIn 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   2055
            Left            =   120
            ScaleHeight     =   2055
            ScaleWidth      =   9495
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   360
            Width           =   9495
            Begin VB.CommandButton btnFileFolder 
               Caption         =   "&Add Files(s) in Folder ..."
               Height          =   375
               Left            =   0
               TabIndex        =   39
               ToolTipText     =   "Select all files in a folder"
               Top             =   480
               Width           =   2775
            End
            Begin VB.PictureBox picFilePrompt 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               HasDC           =   0   'False
               Height          =   1095
               Left            =   3240
               OLEDropMode     =   1  'Manual
               ScaleHeight     =   1095
               ScaleWidth      =   5775
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   360
               Width           =   5775
               Begin VB.Label lblFilePrompt 
                  BackStyle       =   0  'Transparent
                  Caption         =   $"Main.frx":0A42
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   675
                  Left            =   120
                  OLEDropMode     =   1  'Manual
                  TabIndex        =   44
                  Top             =   120
                  Width           =   5520
               End
            End
            Begin VB.CommandButton btnFileDelSel 
               Caption         =   "&Remove Selected Files"
               Enabled         =   0   'False
               Height          =   375
               Left            =   0
               TabIndex        =   40
               ToolTipText     =   "Remove Selected Files from list of Files to be Hashed"
               Top             =   1080
               Width           =   2775
            End
            Begin VB.CommandButton btnFileDelAll 
               Caption         =   "Remove &All Files"
               Enabled         =   0   'False
               Height          =   375
               Left            =   0
               TabIndex        =   41
               ToolTipText     =   "Remove All Files from list of Files to be Hashed"
               Top             =   1560
               Width           =   2775
            End
            Begin VB.CommandButton btnFileBrowse 
               Caption         =   "&Add File(s) ..."
               Height          =   375
               Left            =   0
               TabIndex        =   38
               ToolTipText     =   "Select one or more files whose Hash will be computed"
               Top             =   0
               Width           =   2775
            End
            Begin MSComctlLib.ListView lvFileIn 
               Height          =   1965
               Left            =   2880
               TabIndex        =   42
               Top             =   0
               Width           =   6615
               _ExtentX        =   11668
               _ExtentY        =   3466
               View            =   3
               LabelEdit       =   1
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               OLEDropMode     =   1
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               Appearance      =   1
               OLEDropMode     =   1
               NumItems        =   7
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "File"
                  Object.Width           =   882
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   1
                  Key             =   "MD5"
                  Text            =   "MD5 Hash"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   2
                  Key             =   "SHA1"
                  Text            =   "SHA-1 Hash"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   3
                  Key             =   "SHA224"
                  Text            =   "SHA-224 Hash"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   4
                  Key             =   "SHA256"
                  Text            =   "SHA-256 Hash"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   5
                  Key             =   "SHA384"
                  Text            =   "SHA-384 Hash"
                  Object.Width           =   0
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   6
                  Key             =   "SHA512"
                  Text            =   "SHA-512 Hash"
                  Object.Width           =   0
               EndProperty
            End
         End
      End
   End
   Begin VB.PictureBox picHash 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   4815
      Index           =   2
      Left            =   240
      ScaleHeight     =   4815
      ScaleWidth      =   9735
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   9735
      Begin VB.Frame fraTextCalc 
         Caption         =   "Step 2: Select Type of Hash, then Click ""Compute Hash"" button "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   0
         TabIndex        =   13
         Top             =   2760
         Width           =   9735
         Begin VB.PictureBox picTextHash 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   855
            Index           =   2
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   9255
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   840
            Width           =   9255
            Begin VB.TextBox txtTextCalc 
               Height          =   825
               Index           =   2
               Left            =   2760
               TabIndex        =   20
               Top             =   0
               Width           =   6480
            End
            Begin VB.CommandButton btnTextCalc 
               Caption         =   "&Compute SHA-1 Hash"
               Height          =   825
               Index           =   2
               Left            =   0
               TabIndex        =   19
               ToolTipText     =   "Compute SHA1 Hash of Text"
               Top             =   0
               Width           =   2655
            End
         End
         Begin VB.PictureBox picTextHash 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   855
            Index           =   1
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   9255
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   840
            Width           =   9255
            Begin VB.TextBox txtTextCalc 
               Height          =   825
               Index           =   1
               Left            =   2760
               TabIndex        =   17
               Top             =   0
               Width           =   6480
            End
            Begin VB.CommandButton btnTextCalc 
               Caption         =   "&Compute MD5 Hash"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   825
               Index           =   1
               Left            =   0
               TabIndex        =   16
               ToolTipText     =   "Compute MD5 Hash of Text"
               Top             =   0
               Width           =   2655
            End
         End
         Begin VB.PictureBox picTextCalc 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   675
            Left            =   3360
            ScaleHeight     =   675
            ScaleWidth      =   6015
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   960
            Width           =   6015
            Begin VB.Label lblTextCalc 
               BackStyle       =   0  'Transparent
               Caption         =   "Select a hashing algorithm, using the tabs shown above,\nthen click the ""Compute Hash"" button."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   120
               OLEDropMode     =   1  'Manual
               TabIndex        =   34
               Tag             =   "6010"
               Top             =   0
               Width           =   5640
            End
         End
         Begin VB.PictureBox picTextHash 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   855
            Index           =   6
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   9255
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   840
            Width           =   9255
            Begin VB.CommandButton btnTextCalc 
               Caption         =   "&Compute SHA-512 Hash"
               Height          =   825
               Index           =   6
               Left            =   0
               TabIndex        =   31
               ToolTipText     =   "Compute SHA-512 hash of text"
               Top             =   0
               Width           =   2655
            End
            Begin VB.TextBox txtTextCalc 
               Height          =   825
               Index           =   6
               Left            =   2760
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   32
               Top             =   0
               Width           =   6480
            End
         End
         Begin VB.PictureBox picTextHash 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   855
            Index           =   5
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   9255
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   840
            Width           =   9255
            Begin VB.CommandButton btnTextCalc 
               Caption         =   "&Compute SHA-384 Hash"
               Height          =   825
               Index           =   5
               Left            =   0
               TabIndex        =   28
               ToolTipText     =   "Compute SHA-384 hash of text"
               Top             =   0
               Width           =   2655
            End
            Begin VB.TextBox txtTextCalc 
               Height          =   825
               Index           =   5
               Left            =   2760
               MultiLine       =   -1  'True
               TabIndex        =   29
               Top             =   0
               Width           =   6480
            End
         End
         Begin VB.PictureBox picTextHash 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   855
            Index           =   4
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   9255
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   840
            Width           =   9255
            Begin VB.TextBox txtTextCalc 
               Height          =   825
               Index           =   4
               Left            =   2760
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   26
               Top             =   0
               Width           =   6480
            End
            Begin VB.CommandButton btnTextCalc 
               Caption         =   "&Compute SHA-256 Hash"
               Height          =   825
               Index           =   4
               Left            =   0
               TabIndex        =   25
               ToolTipText     =   "Compute SHA-256 hash of text"
               Top             =   0
               Width           =   2655
            End
         End
         Begin VB.PictureBox picTextHash 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   855
            Index           =   3
            Left            =   240
            ScaleHeight     =   855
            ScaleWidth      =   9255
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   840
            Width           =   9255
            Begin VB.CommandButton btnTextCalc 
               Caption         =   "&Compute SHA-224 Hash"
               Height          =   825
               Index           =   3
               Left            =   0
               TabIndex        =   22
               ToolTipText     =   "Compute SHA-224 hash of text"
               Top             =   0
               Width           =   2655
            End
            Begin VB.TextBox txtTextCalc 
               Height          =   825
               Index           =   3
               Left            =   2760
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   23
               Top             =   0
               Width           =   6480
            End
         End
         Begin MSComctlLib.TabStrip tabText 
            Height          =   1455
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   2566
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   6
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   " &MD5 "
                  Object.Tag             =   "1"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   " SHA-&1 "
                  Object.Tag             =   "2"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   " SHA-22&4 "
                  Object.Tag             =   "3"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   " SHA-&256 "
                  Object.Tag             =   "4"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   " SHA-&384 "
                  Object.Tag             =   "5"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "SHA-&512 "
                  Object.Tag             =   "6"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame fraTextIn 
         Caption         =   "Step 1: Enter or Paste Text to be Hashed: 0 characters "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   9735
         Begin VB.PictureBox picTextIn 
            BorderStyle     =   0  'None
            HasDC           =   0   'False
            Height          =   2055
            Left            =   120
            ScaleHeight     =   2055
            ScaleWidth      =   9495
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   360
            Width           =   9495
            Begin VB.CommandButton btnPaste 
               Caption         =   "&Paste Text"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   0
               TabIndex        =   8
               ToolTipText     =   "Paste contents of clipboard"
               Top             =   0
               Width           =   2775
            End
            Begin VB.PictureBox picTextPrompt 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               HasDC           =   0   'False
               Height          =   1095
               Left            =   3240
               ScaleHeight     =   1095
               ScaleWidth      =   5535
               TabIndex        =   11
               Top             =   360
               Width           =   5535
               Begin VB.Label lblTextPrompt 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Enter Text Here,\nor click ""Paste Text"" to Paste Text from the Clipboard."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   675
                  Left            =   120
                  OLEDropMode     =   1  'Manual
                  TabIndex        =   12
                  Top             =   120
                  Width           =   5160
               End
            End
            Begin VB.TextBox txtTextIn 
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1965
               Left            =   2880
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   10
               ToolTipText     =   "Enter text whose MD5 Hash will be computed"
               Top             =   0
               Width           =   6615
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "The Hash Value of the entire Text String will be Computed."
               Height          =   735
               Left            =   0
               TabIndex        =   9
               Top             =   480
               Visible         =   0   'False
               Width           =   2775
            End
         End
      End
   End
   Begin MSComctlLib.TabStrip tabHash 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9551
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " &Welcome "
            Object.ToolTipText     =   "Welcome to Karen's Hasher!"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " Hash &Text "
            Object.ToolTipText     =   "Compute hash value of a some text"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " Hash &Individual Files "
            Object.ToolTipText     =   "Compute hash value of each file"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " Hash &Group of Files "
            Object.ToolTipText     =   "Compute one hash value for a entire group of files"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " &Verify Saved Hashes "
            Object.ToolTipText     =   "Verify previously computed hashes of files, folders, and groups of files"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   " &Settings "
            Object.ToolTipText     =   "Program settings"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btnHelp 
      Caption         =   "&Help"
      Height          =   375
      Left            =   6240
      TabIndex        =   123
      Tag             =   "6003"
      ToolTipText     =   "Display program's help file"
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save Results to Disk ..."
      Height          =   375
      Left            =   2880
      TabIndex        =   122
      Tag             =   "1002"
      ToolTipText     =   "Copy your computed hashes, and other information, to a disk file"
      Top             =   5640
      Width           =   2655
   End
   Begin VB.CommandButton btnCopy 
      Caption         =   "Co&py Results to Clipboard"
      Height          =   375
      Left            =   120
      TabIndex        =   121
      Tag             =   "1001"
      ToolTipText     =   "Copy your computed hashes, and other information, to Windows' clipboard"
      Top             =   5640
      Width           =   2655
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   126
      Top             =   6180
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14870
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1376
            MinWidth        =   18
            TextSave        =   "6/8/2007"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1217
            MinWidth        =   18
            TextSave        =   "6:42 AM"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btnAbout 
      Caption         =   "&About ..."
      Height          =   375
      Left            =   7560
      TabIndex        =   124
      Tag             =   "6002"
      ToolTipText     =   "About this program"
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton btnExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   8880
      TabIndex        =   125
      Tag             =   "6001"
      ToolTipText     =   "Exit this program"
      Top             =   5640
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdHash 
      Left            =   5640
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8880
      TabIndex        =   127
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Hash File"
      Visible         =   0   'False
      Begin VB.Menu mnuFileCopyAll 
         Caption         =   "Copy All Filenames and Hashes to Clipboard"
      End
      Begin VB.Menu mnuFileCopyFileAll 
         Caption         =   "Copy All Filenames to Clipboard"
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCopySel 
         Caption         =   "Copy Selected Filenames and Hashes to Clipboard"
      End
      Begin VB.Menu mnuFileCopyFileSel 
         Caption         =   "Copy Selected Filenames to Clipboard"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileDelSel 
         Caption         =   "Remove Selected  File(s)"
      End
      Begin VB.Menu mnuFileDelAll 
         Caption         =   "Remove All Files"
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCancel 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "Hash Files"
      Visible         =   0   'False
      Begin VB.Menu mnuFilesCopyAll 
         Caption         =   "Copy All Filenames, and Group's Hash, to Clipboard"
      End
      Begin VB.Menu mnuFilesCopyFileAll 
         Caption         =   "Copy All Filenames to Clipboard"
      End
      Begin VB.Menu mnuFilesSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilesCopyFileSel 
         Caption         =   "Copy Selected Filenames to Clipboard"
      End
      Begin VB.Menu mnuFilesSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilesDelSel 
         Caption         =   "Remove Selected  File(s) From Group"
      End
      Begin VB.Menu mnuFilesDelAll 
         Caption         =   "Remove All Files From Group"
      End
      Begin VB.Menu mnuFilesSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilesCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Copyright © 2002-2005 Karen Kenworthy
' All Rights Reserved
' http://www.karenware.com/

Private Const FRALEFT = 240
Private Const FRATOP = 720
Private Const FRAME_BORDER = 120
Private Const PAD = 120
Private Const PAD2 = 2 * PAD
Private Const PAD3 = 3 * PAD

Private Enum TAB_INX
    TAB_LBOUND = 1
    TAB_WELCOME = TAB_LBOUND
    TAB_TEXT
    TAB_FILE
    TAB_GROUP
    TAB_VERIFY
    TAB_SETTINGS
    TAB_UBOUND = TAB_SETTINGS
End Enum

Private Enum HASH_INX
    HASH_LBOUND = 1
    HASH_MD5 = TAB_LBOUND
    HASH_SHA1
    HASH_SHA224
    HASH_SHA256
    HASH_SHA384
    HASH_SHA512
    HASH_UBOUND = HASH_SHA512
End Enum

Private Enum COL_INX
    COL_FID = 0
    COL_MD5
    COL_FIRST_HASH = COL_MD5
    COL_SHA1
    COL_SHA224
    COL_SHA256
    COL_SHA384
    COL_SHA512
    COL_LAST_HASH = COL_SHA512
End Enum

Private Enum COL_VERIFY
    COL_VERIFY_NAME = 1
    COL_VERIFY_STATUS
    COL_VERIFY_NEW_HASH
    COL_VERIFY_OLD_HASH
End Enum

Private Const URL = 101 ' "http://www.karenware.com"
Private Const USING_MSG = 102 ' "Hash calculations performed by: "
Private Const WELCOME = 103 ' "Welcome to "
Private Const OK_MSG = 104 '"OK"
Private Const INVALID_MSG = 105 '"Invalid"
Private Const NOT_COMPUTED = 106 '"** Not Computed **"
Private Const NOT_VERIFIED = 107 '"Not Verified"
Private Const VERIFIED = 108 '"Verified"
Private Const COMPUTING = 109 '"Computing ..."
Private Const READY_MSG = 110 ' "Ready"
Private Const CANCEL_MSG = 111 ' "Computation Cancelled by User"
Private Const WRONG_HASH = 112 ' "Error - Wrong Hash"
Private Const FILE_NOT_FOUND = 113 '"File Not Found"
Private Const FOLDER_NOT_FOUND = 114 '"Folder Not Found"
Private Const ERROR_MSG = 115 '"Error"

Private Const NOHASHES_CONFIRM = 116 '"No Hashes have been computed. Do you want to %% anyway?"
Private Const OPEN_FAILED = 117 '"Could not open file:"

'Private Const FOLDER_HASH_MISMATCH = 118 '"This hash does not match the folders(s)." & vbCrLf & "Click ""OK"" to copy anyway" & vbCrLf & "Click ""Cancel"" to return to program without copying."
Private Const FILE_HASH_MISMATCH = 119 ' "This hash does not match the files(s)." & vbCrLf & "Click ""OK"" to copy anyway" & vbCrLf & "Click ""Cancel"" to return to program without copying."
Private Const TEXT_HASH_MISMATCH = 120 ' "This hash does not match the text." & vbCrLf & "Click ""OK"" to copy anyway" & vbCrLf & "Click ""Cancel"" to return to program without copying."
'Private Const RE_COMPUTE1 = 121 ' "Copy Cancelled -- Click """
'Private Const RE_COMPUTE2 = 122 ' """ button to Compute Hash"
Private Const SELECT_FILES = 123 ' "Select file(s) to Hash"
Private Const SAVE_FILE = 124 '"Select file to save Hash information"
'Private Const SELECT_FOLDERS = 125 ' "Select folder(s) to Hash"
'Private Const FOLDER_TIP1 = 126 ' "List of folders to Hash -- Click """
'Private Const FOLDER_TIP2 = 127 '""" button, or drag and drop new folders here"
Private Const FILE_TIP1 = 128 '"List of files to Hash -- Click """
Private Const FILE_TIP2 = 129 '""" button, or drag and drop new files here"
Private Const CANNOT_ACCESS = 130 ' "Cannot Access "
Private Const NOT_ADDED = 131 '"Not Added"
Private Const CANNOT_ADD_FOLDERS = 132 '"Can't Add Folders:"
Private Const DRAG_AND_DROP_CANCELLED = 133 ' "Drag and Drop Cancelled"

Private Const FILE_2B_HASHED = 152 ' '"Select Individual Files to be Hashed: "
Private Const GROUP_2B_HASHED = 134 '"Select Files of Group to be Hashed: "
Private Const TEXT_TOBE_HASHED = 136 ' "Text to be Hashed: "
Private Const CHARACTERS = 137 '" characters"
Private Const FOLDERS_MSG = 154 ' " folders "
Private Const REMOVE_SEL_FILES = 138 '"&Remove Selected Files"
Private Const REMOVE_ALL_FILES = 139 '"&Remove All Files"
'Private Const REMOVE_SEL_FOLDERS = 140 ' "&Remove Selected Folders"
'Private Const REMOVE_ALL_FOLDERS = 141 ' "&Remove All Folders"
Private Const COPY_CANCELLED = 142 '"Copy Cancelled"
Private Const SAVE_CANCELLED = 143 '"Save Cancelled"
Private Const HASH_COPIED = 144 '"Hash Copied to Clipboard"
Private Const HASH_SAVED = 145 '"Hash Saved to Disk"
Private Const FILES_REMOVED = 146 '" Files removed from list"
Private Const FOLDERS_REMOVED = 147 '" Folders removed from list"

Private Const SAVE_RESULTS = 149 ' "Save Results"
Private Const COPY_RESULTS = 150 ' "Copy Results"
Private Const FILES_MSG = 151 ' " files "
Private Const NEED_ONE_HASH = 161 ' "At Least One Hash Algorithm Must Be Available"
Private Const BAD_HEX = 148 ' "Hash can only include digits 0-9, and letters A-F"
Private Const ADDING_FILES_IN_FOLDER = 162 ' Adding Files in Folder
Private Const ADDING_FOLDER = 163 ' Adding Folder
Private Const INCLUDE_SUBFOLDERS_PARA = 164 ' Include subfolders of folder "%%FOLDER%%?
Private Const INCLUDE_SUBFOLDERS = 165 ' Include Subfolders?

Private Const HASH_TAB_BASE = 154

Private TextDirty As Boolean
Private TextHashDirty As Boolean
Private FileDirty As Boolean
Private FileHashDirty As Boolean
Private FilesDirty As Boolean
Private FilesHashDirty As Boolean
Private FolderDirty As Boolean
Private FolderHashDirty As Boolean
Private FoldersDirty As Boolean
Private FoldersHashDirty As Boolean
Private Status As Panel
Private PreviousTab As Long
Private GlobalBusy As Boolean

Private PrevShowNetwork As Boolean
Private PrevShowFiles As Boolean
Private PrevSubfolders As Boolean
Private PrevFolder As String

Private MinHeight As Long
Private MinWidth As Long
Private RightMargin As Long
Private BottomMargin As Long
Private HorzGap As Long
Private VertGap As Long
Private BlockGap As Long

Private BACK_NORMAL As OLE_COLOR
Private BACK_GOOD As OLE_COLOR
Private BACK_WARN As OLE_COLOR
Private BACK_ERROR As OLE_COLOR

Private VerifyFid As String
Private VerifyHashType As HASH_TYPE
Private VerifyGroup As Boolean
Private VerifyRelativePath As Boolean
Private VerifyDelimiter As String
Private VerifyPath As String
Private VerifyFailed As Boolean
Private Sub btnAbout_Click()
    Load frmAbout
    frmAbout.AboutExtra = "Settings Folder: " & ApiSettingsPath()
    frmAbout.Show vbModal
End Sub
Private Sub btnCancel_Click()
    Hasher.Cancel = True
End Sub
Private Sub btnCopy_Click()
    Dim b As Boolean

    Select Case tabHash.SelectedItem.Index
        Case TAB_TEXT
            b = TextCopy()
        Case TAB_FILE
            b = FileCopy()
        Case TAB_GROUP
            b = GroupCopy()
        Case TAB_VERIFY
            b = VerifyCopy()
    End Select

    If b Then
        UpdateStatus LoadResString(HASH_COPIED)
    Else
        UpdateStatus LoadResString(COPY_CANCELLED)
    End If
End Sub
Private Sub btnFileBrowse_Click()
    Dim s As String
    Dim sa() As String
    Dim i As Long
    Dim inx As Long
    Dim LstItm As ListItem

    cdHash.Flags = cdlOFNHideReadOnly Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNShareAware
    cdHash.Filter = "All files (*.*)|*.*"
    cdHash.FilterIndex = 1
    cdHash.DialogTitle = LoadResString(SELECT_FILES)
    cdHash.CancelError = True
    cdHash.MaxFileSize = 32000
    cdHash.Filename = ""
    On Error Resume Next
    cdHash.ShowOpen
    If Err <> 0 Then Exit Sub

    If Len(cdHash.Filename) <= 0 Then Exit Sub

    lvFileIn.Sorted = False

    sa = Split(cdHash.Filename, Chr(0))
    If UBound(sa) > 0 Then
        For inx = 1 To UBound(sa)
            s = sa(0) & "\" & sa(inx)
            For i = 1 To lvFileIn.ListItems.Count
                If StrComp(lvFileIn.ListItems(i).Text, s, vbTextCompare) = 0 Then  ' dup
                    Exit For
                End If
            Next i
            If i > lvFileIn.ListItems.Count Then
                Set LstItm = lvFileIn.ListItems.Add()
                LstItm.Text = s
'                LstItm.Selected = True
'                LstItm.SubItems(COL_STATUS) = resstring(NOT_COMPUTED)
            End If
        Next inx
    Else
        s = sa(0)
        For i = 1 To lvFileIn.ListItems.Count
            If StrComp(lvFileIn.ListItems(i).Text, s, vbTextCompare) = 0 Then   ' dup
                Exit For
            End If
        Next i
        If i > lvFileIn.ListItems.Count Then
            Set LstItm = lvFileIn.ListItems.Add()
            LstItm.Text = s
'            LstItm.Selected = True
'            LstItm.SubItems(COL_STATUS) = resstring(NOT_COMPUTED)
        End If
    End If

    lvFileIn.SortKey = 0
    lvFileIn.SortOrder = lvwAscending
    lvFileIn.Sorted = True

    If lvFileIn.ListItems.Count > 0 Then
        picFilePrompt.Visible = False
    Else
        picFilePrompt.Visible = True
    End If

    LvAdjust lvFileIn
    FileMarkDirty True
    FileEnable
End Sub
Private Sub FileMarkDirty(b As Boolean)
    Dim i As Long
'    Dim Col As Long
'    Dim Row As Long
'    Dim LstItm As ListItem
'
'    For Row = 1 To lvFileIn.ListItems.Count
'        Set LstItm = lvFileIn.ListItems(Row)
'        For Col = 2 To lvFileIn.ColumnHeaders.Count
'            lvFileIn.ColumnHeaders(Col).Width = 0
'            LstItm.SubItems(Col) = ""
'        Next Col
'    Next Row

'    For i = btnFileCalcAll.LBound To btnFileCalcAll.UBound
'        btnFileCalcAll(i).Default = False
'    Next i
    picFileCalc.Visible = True
    FileDirty = b
End Sub
Private Sub btnFileCalcAll_Click(Index As Integer)
    Dim result As Long
    Dim sig() As Byte
    Dim s As String
    Dim i As Long
    Dim Col As Long
    Dim HashType As HASH_TYPE
    Dim LstItm As ListItem
    Dim lim As Long
    Dim Cnt As Long

    Hasher.Cancel = False
    FileEnable False
    UpdateStatus LoadResString(COMPUTING)
    picFileCalc.Visible = False
    DoEvents

    s = ""
    Col = COL_FIRST_HASH + Index - 1
    HashType = Index
    lim = lvFileIn.ListItems.Count
    HashFileTot = lim
    HashFileCnt = 0
    lvFileIn.ColumnHeaders(Col + 1).Width = 10
    For i = 1 To lim
        If Hasher.Cancel Then Exit For
        Set LstItm = lvFileIn.ListItems(i)
        s = LstItm.Text
        sig = HashFile(HashType, s)
        If Hasher.HashResult = HASH_OK Then
            LstItm.SubItems(Col) = HashSig2Text(sig)
            Cnt = Cnt + 1
        Else
            LstItm.SubItems(Col) = HashErrorMsg(Hasher.HashResult)
            ApiBeep
        End If
        If (i Mod 8) = 1 Then
            LvAdjust lvFileIn
            LstItm.EnsureVisible
            DoEvents
        End If

        DoEvents
    Next i

    If Not (LstItm Is Nothing) Then LstItm.EnsureVisible
    LvAdjust lvFileIn

    If Hasher.Cancel Then
        UpdateStatus LoadResString(CANCEL_MSG)
        Hasher.Cancel = False
        picFileCalc.Visible = True
    Else
        UpdateStatus LoadResString(READY_MSG)
        FileMarkDirty False
        If Cnt = lim Then
            picFileCalc.Visible = False
        Else
            picFileCalc.Visible = True
        End If
    End If

    tabHash.Enabled = True
    FileEnable
    Hasher.Cancel = False
End Sub
Private Sub btnFileCalcSel_Click(Index As Integer)
    Dim result As Long
    Dim sig() As Byte
    Dim s As String
    Dim i As Long
    Dim HashType As HASH_TYPE
    Dim Col As Long
    Dim LstItm As ListItem
    Dim lim As Long
    Dim Cnt As Long

    Hasher.Cancel = False
    FileEnable False
    UpdateStatus LoadResString(COMPUTING)
    DoEvents

    Col = COL_FIRST_HASH + Index - 1
    HashType = Index
    s = ""
    lim = lvFileIn.ListItems.Count
    For i = 1 To lim
        If Hasher.Cancel Then Exit For
        Set LstItm = lvFileIn.ListItems(i)
        If LstItm.Selected Then
            lvFileIn.ColumnHeaders(Col + 1).Width = 10
            LvAdjust lvFileIn
            DoEvents
            s = LstItm.Text
            sig = HashFile(HashType, s)
            If Hasher.HashResult = HASH_OK Then
                LstItm.SubItems(Col) = HashSig2Text(sig)
                Cnt = Cnt + 1
            Else
                LstItm.SubItems(Col) = HashErrorMsg(Hasher.HashResult)
                ApiBeep
            End If
        End If
        LvAdjust lvFileIn
        DoEvents
    Next i

    If Hasher.Cancel Then
        UpdateStatus LoadResString(CANCEL_MSG)
        Hasher.Cancel = False
        picFileCalc.Visible = True
    Else
        UpdateStatus LoadResString(READY_MSG)
        FileMarkDirty False
        If Cnt = lim Then
            picFilePrompt.Visible = False
        Else
            picFilePrompt.Visible = True
        End If
    End If

    tabHash.Enabled = True
    FileEnable
    Hasher.Cancel = False
End Sub
Private Sub btnFileDelAll_Click()
    mnuFileDelAll_Click
End Sub
Private Sub btnFileDelSel_Click()
    mnuFileDelSel_Click
End Sub
Private Sub btnFileFolder_Click()
    Dim Cancelled As Boolean
    Dim Recurse As Boolean
    
    Load frmBrowse
    If Len(PrevFolder) > 0 Then
        frmBrowse.ShowFiles = PrevShowFiles
        frmBrowse.ShowNetwork = PrevShowNetwork
        frmBrowse.AddSubfolders = PrevSubfolders
        frmBrowse.SelFolder = PrevFolder
    End If

    frmBrowse.Show vbModal
    Cancelled = frmBrowse.Cancelled

    If Not Cancelled Then
        PrevShowFiles = frmBrowse.ShowFiles
        PrevShowNetwork = frmBrowse.ShowNetwork
        PrevSubfolders = frmBrowse.AddSubfolders
        PrevFolder = frmBrowse.SelFolder
        Recurse = frmBrowse.AddSubfolders
    End If

    Unload frmBrowse

    If Cancelled Then Exit Sub

    picFilePrompt.Visible = False
    lvFileIn.Sorted = False
    DoEvents

    FileAddFolder PrevFolder, Recurse

    lvFileIn.SortKey = 0
    lvFileIn.SortOrder = lvwAscending
    lvFileIn.Sorted = True

    If lvFileIn.ListItems.Count > 0 Then
        picFilePrompt.Visible = False
    Else
        picFilePrompt.Visible = True
    End If

    LvAdjust lvFileIn
    FileMarkDirty True
    FileEnable
    UpdateStatus LoadResString(READY_MSG)
End Sub
Private Sub FileAddFolder(Folder As String, Recurse As Boolean)
    Dim i As Long
    Dim j As Long
    Dim LstItm As ListItem
    Dim fa As PT_FID_ARRAY
    Dim result As Long
    Dim lim As Long
    Dim lim2 As Long
    Dim Fid As String

    If Right(Folder, 1) <> "\" Then Folder = Folder & "\"
    UpdateStatus LoadResString(ADDING_FILES_IN_FOLDER) & Folder

    fa = ApiFindAll(Folder, "*.*")
    If fa.FidCnt > 0 Then
        lim2 = fa.FidCnt - 1
        For j = 0 To lim2
            Fid = fa.Fids(j)
            If Right(Fid, 1) = "\" Then ' folder
                If Recurse Then FileAddFolder Fid, Recurse
                DoEvents
            Else ' file
                lim = lvFileIn.ListItems.Count
                For i = 1 To lim
                    If StrComp(lvFileIn.ListItems(i).Text, Fid, vbTextCompare) = 0 Then    ' dup
                        Exit For
                    End If
                    If (i Mod 100) = 1 Then DoEvents
                Next i
                If i > lim Then
                    Set LstItm = lvFileIn.ListItems.Add()
                    LstItm.Text = Fid
    '                LstItm.Selected = True
                    If (i Mod 10) = 1 Then
                        LvAdjust lvFileIn
                        DoEvents
                    End If
                End If
            End If
        Next j
    End If
End Sub
Private Sub btnFilesCalc_Click(Index As Integer)
    Dim sig() As Byte
    Dim s As String
    Dim i As Long
    Dim j As Long
    Dim HashType As HASH_TYPE
    Dim lim As Long
    Dim FileCnt As Long
    Dim Fid As String
    Dim Files As String
'    Dim fd As PT_FILE_INFO
    Dim fa As PT_FID_ARRAY

    Hasher.Cancel = False
    GroupEnable False
    UpdateStatus LoadResString(COMPUTING)
    picGroupCalc.Visible = False
    txtFilesCalc(Index).Text = ""
    DoEvents

    HashType = Index
    HashInit (HashType)

    Files = ""
    FileCnt = 0
    lim = lvGroupIn.ListItems.Count
    For i = 1 To lim
        Fid = lvGroupIn.ListItems(i).Text
        If Right(Fid, 1) = "\" Then ' folder
            If FileCnt > 0 Then
                HashFilesInput HashType, Files
            End If
            Files = ""
            FileCnt = 0
            fa = ApiFindFiles(Fid, "*.*")
            If fa.FidCnt > 0 Then
                For j = 0 To fa.FidCnt - 1
                    Files = Files & fa.Fids(j) & vbNullChar
                Next j
                HashFilesInput HashType, Files
            End If
            Files = ""
            FileCnt = 0
        Else
            Files = Files & Fid & vbNullChar
            FileCnt = FileCnt + 1
        End If
    Next i
    If FileCnt > 0 Then
        HashFilesInput HashType, Files
    End If
    sig = HashFini(HashType)

    If Hasher.HashResult = HASH_OK Then
        txtFilesCalc(Index).Text = HashSig2Text(sig)
        UpdateStatus LoadResString(READY_MSG)
    Else
        txtFilesCalc(Index).Text = HashErrorMsg(Hash.HashResult)
        UpdateStatus HashErrorMsg(Hasher.HashResult)
        ApiBeep
    End If

    If Hasher.Cancel Then
        UpdateStatus LoadResString(CANCEL_MSG)
        GroupMarkDirty True
        Hasher.Cancel = False
        picGroupCalc.Visible = True
    Else
        UpdateStatus LoadResString(READY_MSG)
        GroupMarkDirty False
        If Hasher.HashResult = HASH_OK Then
            picGroupCalc.Visible = False
        Else
            picGroupCalc.Visible = True
        End If
    End If

    tabHash.Enabled = True
    GroupEnable
    Hasher.Cancel = False
End Sub
Private Sub btnGroupBrowse_Click()
    Dim s As String
    Dim sa() As String
    Dim i As Long
    Dim inx As Long
    Dim LstItm As ListItem

    cdHash.Flags = cdlOFNHideReadOnly Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNShareAware
    cdHash.Filter = "All files (*.*)|*.*"
    cdHash.FilterIndex = 1
    cdHash.DialogTitle = LoadResString(SELECT_FILES)
    cdHash.CancelError = True
    cdHash.MaxFileSize = 32000
    cdHash.Filename = ""
    On Error Resume Next
    cdHash.ShowOpen
    If Err <> 0 Then Exit Sub

    If Len(cdHash.Filename) <= 0 Then Exit Sub

    lvGroupIn.Sorted = False

    sa = Split(cdHash.Filename, Chr(0))
    If UBound(sa) > 0 Then
        For inx = 1 To UBound(sa)
            s = sa(0) & "\" & sa(inx)
            For i = 1 To lvGroupIn.ListItems.Count
                If StrComp(lvGroupIn.ListItems(i).Text, s, vbTextCompare) = 0 Then   ' dup
                    Exit For
                End If
            Next i
            If i > lvGroupIn.ListItems.Count Then
                Set LstItm = lvGroupIn.ListItems.Add()
                LstItm.Text = s
'                LstItm.Selected = True
            End If
        Next inx
    Else
        s = sa(0)
        For i = 1 To lvGroupIn.ListItems.Count
            If StrComp(lvGroupIn.ListItems(i).Text, s, vbTextCompare) = 0 Then   ' dup
                Exit For
            End If
        Next i
        If i > lvGroupIn.ListItems.Count Then
            Set LstItm = lvGroupIn.ListItems.Add()
            LstItm.Text = s
'            LstItm.Selected = True
        End If
    End If

    lvGroupIn.SortKey = 0
    lvGroupIn.SortOrder = lvwAscending
    lvGroupIn.Sorted = True

    If lvGroupIn.ListItems.Count > 0 Then
        picFilesPrompt.Visible = False
    Else
        picFilesPrompt.Visible = True
    End If

    LvAdjust lvGroupIn
    GroupMarkDirty True
    GroupEnable
End Sub
Private Sub GroupMarkDirty(b As Boolean)
    Dim i As Long

    If b Then
        FilesDirty = True
        For i = txtFilesCalc.LBound To txtFilesCalc.UBound
            txtFilesCalc(i).Text = ""
        Next i
    Else
        FilesDirty = False
    End If
End Sub
Private Sub GroupEnable(Optional b As Boolean = True)
    Dim SelCnt As Long
    Dim i As Long

    If Not b Then
        btnGroupDelAll.Enabled = False
        btnGroupDelSel.Enabled = False
        ArrayEnable btnFilesCalc, False
        btnCopy.Enabled = False
        btnSave.Enabled = False
        tabHash.Enabled = False
        btnCancel.Visible = True
        btnExit.Visible = False
        btnCancel.Cancel = True
        Exit Sub
    End If

    If lvGroupIn.ListItems.Count > 0 Then
        btnGroupDelAll.Enabled = True
        SelCnt = 0
        For i = 1 To lvGroupIn.ListItems.Count
            If lvGroupIn.ListItems(i).Selected Then
                SelCnt = SelCnt + 1
                Exit For
            End If
        Next i
        If SelCnt > 0 Then
            btnGroupDelSel.Enabled = True
        Else
            btnGroupDelSel.Enabled = False
        End If
        ArrayEnable btnFilesCalc, True
        If ArrayLen(txtFilesCalc) > 0 Then
            btnCopy.Enabled = True
            btnSave.Enabled = True
        Else
            btnCopy.Enabled = False
            btnSave.Enabled = False
        End If
    Else
        btnCopy.Enabled = False
        btnSave.Enabled = False
        btnGroupDelAll.Enabled = False
        btnGroupDelSel.Enabled = False
        ArrayEnable btnFilesCalc, False
    End If
    If Len(txtFilesCalc(tabGroup.SelectedItem.Index).Text) > 0 Then
        picGroupCalc.Visible = False
    Else
        picGroupCalc.Visible = True
    End If

    btnExit.Visible = True
    btnCancel.Visible = False
    btnExit.Cancel = True

    fraGroupIn.Caption = LoadResString(GROUP_2B_HASHED) & FormatNumber(lvGroupIn.ListItems.Count, 0) & " files and folders"
End Sub
Private Sub FileEnable(Optional b As Boolean = True)
    Dim i As Long
    Dim SelCnt As Long
    Dim AllOK As Boolean
    Dim AllCalced As Boolean
    Dim SelCalced As Boolean

    If Not b Then
        btnFileDelAll.Enabled = False
        btnFileDelSel.Enabled = False
'        For i = btnFileCalcAll.LBound To btnFileCalcAll.UBound
'            btnFileCalcAll(i).Enabled = False
'            btnFileCalcAll(i).Default = False
'            btnFileCalcSel(i).Enabled = False
'        Next i
        btnCopy.Enabled = False
        btnSave.Enabled = False
        tabHash.Enabled = False
        btnCancel.Visible = True
        btnExit.Visible = False
        btnCancel.Cancel = True
        Exit Sub
    End If

    If lvFileIn.ListItems.Count > 0 Then
        btnFileDelAll.Enabled = True
        For i = btnFileCalcAll.LBound To btnFileCalcAll.UBound
            btnFileCalcAll(i).Enabled = True
        Next i
        AllOK = True
        AllCalced = True
        SelCalced = True
        For i = 1 To lvFileIn.ListItems.Count
            If lvFileIn.ListItems(i).Selected Then
                SelCnt = SelCnt + 1
            End If
        Next i
        If SelCnt > 0 Then
            btnFileDelSel.Enabled = True
            For i = btnFileCalcAll.LBound To btnFileCalcAll.UBound
                btnFileCalcSel(i).Enabled = True
            Next i
        Else
            btnFileDelSel.Enabled = False
            For i = btnFileCalcAll.LBound To btnFileCalcAll.UBound
                btnFileCalcSel(i).Enabled = False
            Next i
        End If
        If AllCalced Then
            btnCopy.Enabled = True
            btnSave.Enabled = True
        Else
            btnCopy.Enabled = False
            btnSave.Enabled = False
        End If
    Else
        btnCopy.Enabled = False
        btnSave.Enabled = False
        btnFileDelAll.Enabled = False
        btnFileDelSel.Enabled = False
        For i = btnFileCalcAll.LBound To btnFileCalcAll.UBound
            btnFileCalcAll(i).Enabled = False
        Next i
    End If

'    For i = btnFileCalcAll.LBound To btnFileCalcAll.UBound
'        If btnFileCalcAll(i).Visible And btnFileCalcAll(i).Enabled And (btnFileCalcAll(i).Left >= 0) Then
'            btnFileCalcAll(i).Default = True
'            Exit For
'        End If
'    Next i

    btnExit.Visible = True
    btnCancel.Visible = False
    btnExit.Cancel = True

    fraFileIn.Caption = LoadResString(FILE_2B_HASHED) & FormatNumber(lvFileIn.ListItems.Count, 0) & LoadResString(FILES_MSG)
End Sub
Private Sub VerifyEnable(Optional b As Boolean = True)
    If Not b Then
        btnCopy.Enabled = False
        btnSave.Enabled = False
        tabHash.Enabled = False
        btnCancel.Visible = True
        btnExit.Visible = False
        btnCancel.Cancel = False
        Exit Sub
    End If

    If btnVerifyUpdate.Enabled Then
        btnCopy.Enabled = True
        btnSave.Enabled = True
    Else
        btnCopy.Enabled = False
        btnSave.Enabled = False
    End If
    tabHash.Enabled = True
    btnCancel.Visible = False
    btnExit.Visible = True
    btnExit.Cancel = True
End Sub
Private Sub ArrayEnable(CntlArray As Object, b As Boolean)
    Dim i As Long

    For i = CntlArray.LBound To CntlArray.UBound
        CntlArray(i).Enabled = b
    Next i
End Sub
Private Function ArrayLen(CntlArray As Object) As Long
    Dim i As Long
    Dim TotLen As Long

    For i = CntlArray.LBound To CntlArray.UBound
        TotLen = TotLen + Len(CntlArray(i).Text)
    Next i
    ArrayLen = TotLen
End Function
Private Sub ArraySetFound(CntlArray As Object, FoundArray() As Boolean)
    Dim i As Long

    For i = CntlArray.LBound To CntlArray.UBound
        If HashLegal(CntlArray(i).Text, i) Then
            FoundArray(i) = True
        Else
            FoundArray(i) = False
        End If
    Next i
End Sub
Private Function ArrayAnyFound(CntlArray As Object) As Boolean
    Dim i As Long

    For i = CntlArray.LBound To CntlArray.UBound
        If HashLegal(CntlArray(i).Text, i) Then
            ArrayAnyFound = True
            Exit Function
        End If
    Next i
    ArrayAnyFound = False
End Function
Private Function ArrayAllFound(CntlArray As Object) As Boolean
    Dim i As Long

    For i = CntlArray.LBound To CntlArray.UBound
        If HashLegal(CntlArray(i).Text, i) Then
            ArrayAllFound = False
            Exit Function
        End If
    Next i
    ArrayAllFound = True
End Function
Private Sub TextEnable(Optional b As Boolean = True)

    If Not b Then
        btnPaste.Enabled = False
        ArrayEnable btnTextCalc, False
        btnCopy.Enabled = False
        btnSave.Enabled = False
        tabHash.Enabled = False
        btnCancel.Enabled = True
        btnExit.Enabled = False
        btnCancel.Cancel = True
        Exit Sub
    End If

    btnPaste.Enabled = True
    If Len(txtTextIn.Text) > 0 Then
        ArrayEnable btnTextCalc, True
        If ArrayLen(txtTextCalc) > 0 Then
            btnCopy.Enabled = True
            btnSave.Enabled = True
        Else
            btnCopy.Enabled = False
            btnSave.Enabled = False
        End If
    Else
        If ArrayLen(txtTextCalc) > 0 Then
            btnCopy.Enabled = True
            btnSave.Enabled = True
        Else
            btnCopy.Enabled = False
            btnSave.Enabled = False
        End If
        ArrayEnable btnTextCalc, True ' empty string is OK
    End If
    If Len(txtTextCalc(tabText.SelectedItem.Index).Text) > 0 Then
        picTextCalc.Visible = False
    Else
        picTextCalc.Visible = True
    End If

    btnExit.Enabled = True
    btnCancel.Enabled = False
    btnExit.Cancel = True

    fraTextIn.Caption = LoadResString(TEXT_TOBE_HASHED) & FormatNumber(Len(txtTextIn.Text), 0) & LoadResString(CHARACTERS)
End Sub
Private Sub btnGroupDelAll_Click()
    mnuFilesDelAll_Click
End Sub
Private Sub btnGroupDelSel_Click()
    mnuFilesDelSel_Click
End Sub
Private Sub DispHelp()
    Select Case tabHash.SelectedItem.Index
        Case TAB_WELCOME
            ApiHelpTopic 1000
        Case TAB_TEXT
            ApiHelpTopic 2000
        Case TAB_FILE
            ApiHelpTopic 3000
        Case TAB_GROUP
            ApiHelpTopic 4000
        Case TAB_VERIFY
            ApiHelpTopic 7000
        Case TAB_SETTINGS
            ApiHelpTopic 5000
        Case Else
            ApiHelpContents
    End Select
End Sub
Private Sub btnGroupFolder_Click()
    Dim Cancelled As Boolean
    Dim Recurse As Boolean

    Load frmBrowse
    If Len(PrevFolder) > 0 Then
        frmBrowse.ShowFiles = PrevShowFiles
        frmBrowse.ShowNetwork = PrevShowNetwork
        frmBrowse.AddSubfolders = PrevSubfolders
        frmBrowse.SelFolder = PrevFolder
    End If

    frmBrowse.Show vbModal
    Cancelled = frmBrowse.Cancelled

    If Not Cancelled Then
        PrevShowFiles = frmBrowse.ShowFiles
        PrevShowNetwork = frmBrowse.ShowNetwork
        PrevSubfolders = frmBrowse.AddSubfolders
        PrevFolder = frmBrowse.SelFolder
        Recurse = frmBrowse.AddSubfolders
    End If

    Unload frmBrowse

    If Cancelled Then Exit Sub

    picFilesPrompt.Visible = False
    lvGroupIn.Sorted = False
    DoEvents

    FilesAddFolder PrevFolder, Recurse
    lvGroupIn.SortKey = 0
    lvGroupIn.SortOrder = lvwAscending
    lvGroupIn.Sorted = True

    If lvGroupIn.ListItems.Count > 0 Then
        picFilesPrompt.Visible = False
    Else
        picFilesPrompt.Visible = True
    End If

    LvAdjust lvGroupIn
    GroupMarkDirty True
    GroupEnable
    UpdateStatus LoadResString(READY_MSG)
End Sub
Private Sub FilesAddFolder(Folder As String, Recurse As Boolean)
    Dim i As Long
    Dim LstItm As ListItem
    Dim lim As Long

    If Right(Folder, 1) <> "\" Then Folder = Folder & "\"
    UpdateStatus LoadResString(ADDING_FOLDER) & Folder
    lim = lvGroupIn.ListItems.Count
    
    For i = 1 To lim
        If StrComp(lvGroupIn.ListItems(i).Text, Folder, vbTextCompare) = 0 Then   ' dup
            Exit For
        End If
    Next i
    If i > lim Then
        Set LstItm = lvGroupIn.ListItems.Add()
        LstItm.Text = Folder
    End If

    If Not Recurse Then Exit Sub

    Dim fa As PT_FID_ARRAY

    fa = ApiFindFolders(Folder, "*.*")
    If fa.FidCnt > 0 Then
        lim = fa.FidCnt - 1
        For i = 0 To lim
            FilesAddFolder fa.Fids(i), Recurse
        Next i
    End If
End Sub
Private Sub btnHelp_Click()
    DispHelp
End Sub
Private Sub btnSave_Click()
    Dim b As Boolean
    Dim Fid As String
    Dim Ext As String
    Dim i As Long
    Dim s As String
    Dim Desc As String
    Dim Filters As String
    Dim sa() As String
    Dim HashType As HASH_TYPE

    Select Case tabHash.SelectedItem.Index
        Case TAB_TEXT:
            For i = txtTextCalc.LBound To txtTextCalc.UBound
                If HashLegal(txtTextCalc(i).Text, i) Then
                    Desc = HashDesc(i)
                    Filters = Filters & Desc & " Files (*."
                    Desc = LCase(Replace(Desc, "-", ""))
                    Filters = Filters & Desc & ")|*." & Desc & "|"
                End If
            Next i

        Case TAB_FILE:
            For i = COL_FIRST_HASH To COL_LAST_HASH
                If lvFileIn.ColumnHeaders(i + 1).Width > 0 Then
                    Desc = HashDesc(i - COL_FIRST_HASH + 1)
                    Filters = Filters & Desc & " Files (*."
                    Desc = LCase(Replace(Desc, "-", ""))
                    Filters = Filters & Desc & ")|*." & Desc & "|"
                End If
            Next i

        Case TAB_GROUP:
            For i = txtFilesCalc.LBound To txtFilesCalc.UBound
                If HashLegal(txtFilesCalc(i).Text, i) Then
                    Desc = HashDesc(i)
                    Filters = Filters & Desc & " Files (*."
                    Desc = LCase(Replace(Desc, "-", ""))
                    Filters = Filters & Desc & ")|*." & Desc & "|"
                End If
            Next i

        Case TAB_VERIFY
            Desc = HashDesc(VerifyHashType)
            Filters = Desc & " Files (*."
            Desc = LCase(Replace(Desc, "-", ""))
            Filters = Filters & Desc & ")|*." & Desc & "|"

    End Select

    cdHash.Flags = cdlOFNHideReadOnly Or cdlOFNExplorer Or cdlOFNLongNames Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist Or cdlOFNShareAware
    cdHash.DialogTitle = LoadResString(SAVE_FILE)
    cdHash.Filter = Filters & "Text files (*.txt)|*.txt|" & "All files (*.*)|*.*"
    cdHash.FilterIndex = 1
    cdHash.MaxFileSize = 32767
    cdHash.Filename = ""
    cdHash.CancelError = True

    On Error Resume Next
    cdHash.ShowSave
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If

    Fid = cdHash.Filename
    HashType = HASH_NONE
    i = InStrRev(Fid, ".")
    If i > 0 Then
        Ext = Mid(Fid, i)
        HashType = HashExt2Type(Ext)
    End If

    Select Case tabHash.SelectedItem.Index
        Case TAB_TEXT
            b = TextCopy(Fid, HashType)
        Case TAB_FILE
            b = FileCopy(Fid, HashType)
        Case TAB_GROUP
            b = GroupCopy(Fid, HashType)
        Case TAB_VERIFY
            b = VerifyCopy(Fid, HashType)
    End Select

    If b Then
        UpdateStatus LoadResString(HASH_SAVED)
    Else
        UpdateStatus LoadResString(SAVE_CANCELLED)
    End If
End Sub
Private Sub btnExit_Click()
    If Not SettingsOK() Then Exit Sub
    Unload Me
End Sub
Private Function FileCopy(Optional Fid As String = "", Optional HashType As HASH_TYPE = HASH_NONE) As Boolean
    Dim s As String
    Dim fn As Long
    Dim Hash As String
    Dim ColHdr As ColumnHeader
    Dim LstItm As ListItem
    Dim lim As Long
    Dim Col As Long
    Dim Row As Long

    FileCopy = False
    lim = lvFileIn.ListItems.Count

    If Len(Fid) > 0 Then
        fn = FreeFile()
        On Error Resume Next
        Open Fid For Output Access Write As fn
        If Err.Number <> 0 Then
            MsgBox LoadResString(OPEN_FAILED) & vbCrLf & Fid & vbCrLf & vbCrLf & Error & Err.Description & " (" & Err.Number & ")", vbExclamation Or vbOKOnly, App.FileDescription
            Err.Clear
            Exit Function
        End If
    End If

    If HashType = HASH_NONE Then
        s = CopyHeader() & vbCrLf
        s = s & "Files Hashed: " & FormatNumber(lim, 0) & vbCrLf
        s = s & vbCrLf
        s = s & "File Name"
        For Col = COL_FIRST_HASH To COL_LAST_HASH
            Set ColHdr = lvFileIn.ColumnHeaders(Col + 1)
            If ColHdr.Width > 0 Then s = s & vbTab & ColHdr.Text
        Next Col
    
        s = s & vbCrLf
        For Row = 1 To lim
            Set LstItm = lvFileIn.ListItems(Row)
            s = s & LstItm.Text
            For Col = COL_FIRST_HASH To COL_LAST_HASH
                Set ColHdr = lvFileIn.ColumnHeaders(Col + 1)
                If ColHdr.Width > 0 Then
                    If HashLegal(LstItm.SubItems(Col), Col - COL_FIRST_HASH + 1) Then
                        s = s & vbTab & LstItm.SubItems(Col)
                    Else
                        s = s & vbTab & LoadResString(NOT_COMPUTED)
                    End If
                End If
            Next Col
            s = s & vbCrLf
        Next Row
    Else
        Col = HashType + COL_FIRST_HASH - 1
        Set ColHdr = lvFileIn.ColumnHeaders(Col + 1)
        For Row = 1 To lim
            Set LstItm = lvFileIn.ListItems(Row)
'            If ColHdr.Width > 0 Then
                If HashLegal(LstItm.SubItems(Col), HashType) Then
                    s = s & LstItm.SubItems(Col)
                Else
                    s = s & LoadResString(NOT_COMPUTED)
                End If
                s = s & vbTab & LstItm.Text & vbCrLf
'            End If
        Next Row
    End If

    If Len(Fid) > 0 Then
        Print #fn, s;
        Close #fn
    Else
        Clipboard.Clear
        Clipboard.SetText Trim(s)
    End If
    FileCopy = True
End Function
Private Function CopyHeader() As String
    Dim ver As String

    ver = "v" & App.Major & "." & App.Minor
    If App.Revision > 0 Then ver = ver & "." & App.Revision

    ver = App.FileDescription & " " & ver & vbCrLf
    ver = ver & LoadResString(URL) & vbCrLf & vbCrLf
    ver = ver & "Date:" & vbTab & CStr(Now()) & vbCrLf
    ver = ver & "Computer:" & vbTab & ApiComputerName() & vbCrLf
    ver = ver & "User:" & vbTab & ApiUserName() & vbCrLf
    CopyHeader = ver
End Function
Private Function VerifyCopy(Optional Fid As String = "", Optional HashType As HASH_TYPE = HASH_NONE) As Boolean
    Dim fn As Long
    Dim s As String
    Dim i As Long
    Dim Ext As String
    Dim AddHeader As Boolean
    Dim lim As Long

    VerifyCopy = False

    s = ""
    If Len(Fid) > 0 Then
        fn = FreeFile()
        On Error Resume Next
        Open Fid For Output Access Write As fn
        If Err.Number <> 0 Then
            MsgBox LoadResString(OPEN_FAILED) & vbCrLf & Fid & vbCrLf & vbCrLf & Error & Err.Description & " (" & Err.Number & ")", vbExclamation Or vbOKOnly, App.FileDescription
            Err.Clear
            Exit Function
        End If
        i = InStrRev(Fid, ".")
        If i > 0 Then
            Ext = Mid(Fid, Ext)
        Else
            Ext = ""
        End If
        If HashExt2Type(Ext) = HASH_NONE Then AddHeader = True
    Else
        AddHeader = True
    End If

    lim = lvVerifyIn.ListItems.Count
    If AddHeader Then
        s = CopyHeader() & vbCrLf
        s = s & "Files or Folders Hashed: " & FormatNumber(lim, 0) & vbCrLf
        s = s & vbCrLf
        s = s & "File or Folder Name" & vbTab & "Current " & HashDesc(VerifyHashType) & " Hash" & vbCrLf
    End If

    Dim LstItm As ListItem

    If VerifyGroup Then
        s = s & txtVerifyNewHash.Text
        For i = 1 To lim
            s = s & vbTab & lvVerifyIn.ListItems(i).SubItems(COL_VERIFY_NAME)
        Next i
        s = s & vbCrLf
    Else
        For i = 1 To lim
            Set LstItm = lvVerifyIn.ListItems(i)
            s = s & LstItm.SubItems(COL_VERIFY_NEW_HASH)
            s = s & vbTab & LstItm.SubItems(COL_VERIFY_NAME) & vbCrLf
        Next i
    End If

    If Len(Fid) > 0 Then
        Print #fn, s;
        Close #fn
    Else
        Clipboard.Clear
        Clipboard.SetText Trim(s)
    End If
    VerifyCopy = True
End Function
Private Function GroupCopy(Optional Fid As String = "", Optional HashType As HASH_TYPE = HASH_NONE) As Boolean
    Dim s As String
    Dim yorn As VbMsgBoxResult
    Dim ver As String
    Dim i As Long
    Dim fn As Long
    Dim lim As Long
    Dim LstItm As ListItem

    GroupCopy = False
    If Not ArrayAnyFound(txtFilesCalc) Then
        If Len(Fid) > 0 Then
            s = Replace(LoadResString(NOHASHES_CONFIRM), "%%ACTION%%", LoadResString(SAVE_RESULTS))
        Else
            s = Replace(LoadResString(NOHASHES_CONFIRM), "%%ACTION%%", LoadResString(COPY_RESULTS))
        End If
        yorn = MsgBox(s, vbQuestion Or vbYesNo, App.FileDescription)
        If yorn <> vbYes Then Exit Function
    End If

    If Len(Fid) > 0 Then
        fn = FreeFile()
        On Error Resume Next
        Open Fid For Output Access Write As fn
        If Err.Number <> 0 Then
            MsgBox LoadResString(OPEN_FAILED) & vbCrLf & Fid & vbCrLf & vbCrLf & Error & Err.Description & " (" & Err.Number & ")", vbExclamation Or vbOKOnly, App.FileDescription
            Err.Clear
            Exit Function
        End If
    End If

    lim = lvGroupIn.ListItems.Count
    If HashType = HASH_NONE Then
        s = CopyHeader() & vbCrLf
        For i = txtFilesCalc.LBound To txtFilesCalc.UBound
            If HashLegal(txtFilesCalc(i).Text, i) Then s = s & HashDesc(i) & " Hash: " & vbTab & txtFilesCalc(i).Text & vbCrLf
        Next i
        s = s & vbCrLf
        s = s & "Files and Folders in Group: " & FormatNumber(lim, 0) & vbCrLf
        For i = 1 To lim
            s = s & lvGroupIn.ListItems(i).Text & vbCrLf
        Next i
    Else
        If HashLegal(txtFilesCalc(HashType).Text, HashType) Then
            s = txtFilesCalc(HashType).Text
        Else
            s = LoadResString(NOT_COMPUTED)
        End If
        For i = 1 To lim
            s = s & vbTab & lvGroupIn.ListItems(i).Text
        Next i
        s = s & vbCrLf
    End If

    If Len(Fid) > 0 Then
        Print #fn, s;
        Close #fn
    Else
        Clipboard.Clear
        Clipboard.SetText Trim(s)
    End If
    GroupCopy = True
End Function
Private Sub btnPaste_Click()
    txtTextIn.SelText = Clipboard.GetText()
    picTextPrompt.Visible = False
    TextMarkDirty True
    TextEnable
End Sub
Private Function TextCopy(Optional Fid As String = "", Optional HashType As HASH_TYPE = HASH_NONE) As Boolean
    Dim s As String
    Dim i As Long
    Dim yorn As VbMsgBoxResult
    Dim ver As String
    Dim fn As Long

    TextCopy = False
    If Not ArrayAnyFound(txtTextCalc) Then
        If Len(Fid) > 0 Then
            s = Replace(LoadResString(NOHASHES_CONFIRM), "%%ACTION%%", LoadResString(SAVE_RESULTS))
        Else
            s = Replace(LoadResString(NOHASHES_CONFIRM), "%%ACTION%%", LoadResString(COPY_RESULTS))
        End If
        yorn = MsgBox(s, vbQuestion Or vbYesNo, App.FileDescription)
        If yorn <> vbYes Then Exit Function
    End If

    If Len(Fid) > 0 Then
        fn = FreeFile()
        On Error Resume Next
        Open Fid For Output Access Write As fn
        If Err.Number <> 0 Then
            MsgBox LoadResString(OPEN_FAILED) & vbCrLf & Fid & vbCrLf & vbCrLf & Error & Err.Description & " (" & Err.Number & ")", vbExclamation Or vbOKOnly, App.FileDescription
            Err.Clear
            Exit Function
        End If
    End If

    If HashType = HASH_NONE Then
        s = CopyHeader() & vbCrLf
        For i = txtTextCalc.LBound To txtTextCalc.UBound
            If HashLegal(txtTextCalc(i).Text, i) Then s = s & HashDesc(i) & " Hash: " & vbTab & txtTextCalc(i).Text & vbCrLf
        Next i
        s = s & vbCrLf
        s = s & "Text (" & FormatNumber(Len(txtTextIn.Text), 0) & " characters):" & vbCrLf
        s = s & txtTextIn.Text
    Else
        If HashLegal(txtTextCalc(HashType).Text, HashType) Then
            s = txtTextCalc(HashType).Text
        Else
            s = LoadResString(NOT_COMPUTED)
        End If
        s = s & vbTab & txtTextIn.Text & vbCrLf
    End If

    If Len(Fid) > 0 Then
        Print #fn, s;
        Close #fn
    Else
        Clipboard.Clear
        Clipboard.SetText Trim(s)
    End If
    TextCopy = True
End Function
Private Sub btnTest_Click()
    lstTest.Clear
    HashTest lstTest
End Sub
Private Sub btnTextCalc_Click(Index As Integer)
    Dim sig() As Byte
    Dim s As String
    Dim HashType As HASH_TYPE

    Hasher.Cancel = False
    TextEnable False
    UpdateStatus LoadResString(COMPUTING)
    picTextCalc.Visible = False
    txtTextCalc(Index).Text = ""
    DoEvents

    s = txtTextIn.Text
    HashType = Index
    sig = HashString(HashType, s)

    If Hasher.HashResult = HASH_OK Then
        txtTextCalc(Index).Text = HashSig2Text(sig)
    Else
        txtTextCalc(Index).Text = HashErrorMsg(Hasher.HashResult)
        ApiBeep
    End If

    If Hasher.Cancel Then
        UpdateStatus LoadResString(CANCEL_MSG)
        TextMarkDirty True
        Hasher.Cancel = False
        picTextCalc.Visible = True
    Else
        UpdateStatus LoadResString(READY_MSG)
        TextMarkDirty False
        If Hasher.HashResult = HASH_OK Then
            picTextCalc.Visible = False
        Else
            picTextCalc.Visible = True
        End If
    End If

    tabHash.Enabled = True
    TextEnable
    Hasher.Cancel = False
End Sub
Private Sub btnVerifyCalc_Click()
    Dim sig() As Byte
    Dim s As String
    Dim i As Long
    Dim j As Long
    Dim HashType As HASH_TYPE
    Dim lim As Long
    Dim FileCnt As Long
    Dim Fid As String
    Dim Files As String
    Dim fa As PT_FID_ARRAY
    Dim LstItm As ListItem
    Dim Cnt As Long
    Dim result As HASH_RESULT
    Dim ErrsFound As Boolean
    Dim FailCnt As Long

    Hash.Cancel = False
'    VerifyEnable False
    UpdateStatus LoadResString(COMPUTING)
    txtVerifyNewHash.Text = ""
    DoEvents

    ErrsFound = False
    Files = ""
    FileCnt = 0
    lim = lvVerifyIn.ListItems.Count

    VerifyEnable False
    If VerifyGroup Then
        lvVerifyIn.ColumnHeaders(COL_VERIFY_STATUS + 1).Width = 100
        HashInit VerifyHashType
        For i = 1 To lim
            If Hash.Cancel Then Exit For
            Set LstItm = lvVerifyIn.ListItems(i)
            Fid = LstItm.SubItems(COL_VERIFY_NAME)
            If Right(Fid, 1) = "\" Then ' folder
                If ApiFolderExists(Fid) Then
                    If FileCnt > 0 Then
                        HashFilesInput VerifyHashType, Files
                    End If
                    Files = ""
                    FileCnt = 0
                    fa = ApiFindFiles(Fid, "*.*")
                    If fa.FidCnt > 0 Then
                        For j = 0 To fa.FidCnt - 1
                            Files = Files & fa.Fids(j) & vbNullChar
                        Next j
                        HashFilesInput VerifyHashType, Files
                    End If
                    Files = ""
                    FileCnt = 0
                Else
                    LstItm.SubItems(COL_VERIFY_STATUS) = LoadResString(FOLDER_NOT_FOUND)
                    ErrsFound = True
                End If
            Else
                If ApiFileExists(Fid) Then
                    Files = Files & Fid & vbNullChar
                    FileCnt = FileCnt + 1
                Else
                    LstItm.SubItems(COL_VERIFY_STATUS) = LoadResString(FILE_NOT_FOUND)
                    ErrsFound = True
                End If
            End If
            If (i Mod 8) = 1 Then
                LvAdjust lvVerifyIn
                LstItm.EnsureVisible
                DoEvents
            End If
        Next i
        If FileCnt > 0 Then
            HashFilesInput VerifyHashType, Files
        End If
        sig = HashFini(VerifyHashType)
        If Hash.HashResult <> HASH_OK Then
            txtVerifyNewHash.Text = HashErrorMsg(Hash.HashResult)
        ElseIf ErrsFound Then
            txtVerifyNewHash.Text = "One or more files could not be processed"
        Else
            txtVerifyNewHash.Text = HashSig2Text(sig)
        End If
        lblVerifyNewHash.Enabled = True
        txtVerifyNewHash.Enabled = True
    Else ' individual files
        FailCnt = 0
        HashFileTot = lim
        HashFileCnt = 0
        lvVerifyIn.ColumnHeaders(COL_VERIFY_STATUS + 1).Width = 100
        lvVerifyIn.ColumnHeaders(COL_VERIFY_NEW_HASH + 1).Width = 100
        For i = 1 To lim
            If Hash.Cancel Then Exit For
            Set LstItm = lvVerifyIn.ListItems(i)
            Fid = LstItm.SubItems(COL_VERIFY_NAME)
            If ApiFileExists(Fid) Then
                sig = HashFile(VerifyHashType, Fid)
                If Hasher.HashResult = HASH_OK Then
                    LstItm.SubItems(COL_VERIFY_NEW_HASH) = HashSig2Text(sig)
                    If StrComp(LstItm.SubItems(COL_VERIFY_OLD_HASH), LstItm.SubItems(COL_VERIFY_NEW_HASH), vbTextCompare) = 0 Then
                        LstItm.SubItems(COL_VERIFY_STATUS) = "Unchanged"
                    Else
                        LstItm.SubItems(COL_VERIFY_STATUS) = "CHANGED"
                        ErrsFound = True
                        FailCnt = FailCnt + 1
                    End If
                    Cnt = Cnt + 1
                Else
                    LstItm.SubItems(COL_VERIFY_NEW_HASH) = HashErrorMsg(Hash.HashResult)
                    ApiBeep
                End If
                If (i Mod 8) = 1 Then
                    LvAdjust lvVerifyIn
                    LstItm.EnsureVisible
                    DoEvents
                End If
            Else
                LstItm.SubItems(COL_VERIFY_NEW_HASH) = LoadResString(FILE_NOT_FOUND)
                LstItm.SubItems(COL_VERIFY_STATUS) = "ERROR"
                ErrsFound = True
                FailCnt = FailCnt + 1
            End If
            DoEvents
        Next i
    End If

    LvAdjust lvVerifyIn

    If VerifyGroup Then
        If StrComp(txtVerifyOldHash.Text, txtVerifyNewHash.Text, vbTextCompare) = 0 Then
            lblVerifyResult.Caption = "Verification" & vbCrLf & "Succeeded"
            lblVerifyResult.ForeColor = RGB(0, 0, 192) ' vbHighlight
            VerifyFailed = False
        Else
            lblVerifyResult.Caption = "Verification" & vbCrLf & "Failed"
            lblVerifyResult.ForeColor = RGB(192, 0, 0) ' vbHighlight
            VerifyFailed = True
        End If
    Else
        If (FailCnt = 0) And (Not ErrsFound) Then
            lblVerifyResult.Caption = "Verification" & vbCrLf & "Succeeded"
            lblVerifyResult.ForeColor = RGB(0, 0, 192) ' vbHighlight
            VerifyFailed = False
        Else
            lblVerifyResult.Caption = "Verification" & vbCrLf & "Failed"
            lblVerifyResult.ForeColor = RGB(192, 0, 0) ' vbHighlight
            VerifyFailed = True
        End If
    End If
    lblVerifyResult.Visible = True

'    If Not (VerifyGroup And VerifyFailed) Then
        btnVerifyUpdate.Enabled = True
'    End If

    UpdateStatus Replace(lblVerifyResult.Caption, vbCrLf, " ")
    ApiBeep
    VerifyEnable True
End Sub
Private Sub btnVerifyOpen_Click()
    Dim AllFiles As String
    Dim Desc As String
    Dim Filters As String
    Dim i As Long

    cdHash.Flags = cdlOFNFileMustExist Or cdlOFNLongNames Or cdlOFNShareAware
    cdHash.DialogTitle = "Open Previously Saved Hash File"
    For i = HASH_TYPE_LBOUND To HASH_TYPE_UBOUND
        Desc = HashDesc(i)
        Filters = Filters & Desc & " Files (*."
        Desc = LCase(Replace(Desc, "-", ""))
        Filters = Filters & Desc & ")|*." & Desc & "|"
        If Len(AllFiles) > 0 Then AllFiles = AllFiles & ";"
        AllFiles = AllFiles & "*." & Desc
    Next i
'    cdHash.Filter = "Text files (*.txt)|*.txt|" & Filters & "PTHash files (*.pthash)|*.pthash|All files (*.*)|*.*"
    cdHash.Filter = "All Hash Files|" & AllFiles & "|" & Filters '& "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    cdHash.FilterIndex = 1
    cdHash.MaxFileSize = 32767
    cdHash.CancelError = True
    On Error Resume Next
    cdHash.ShowOpen
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If

    VerifyFid = cdHash.Filename
    VerifyParse VerifyFid
    VerifyEnable
End Sub
Private Sub VerifyParse(Fid As String)
    Dim Ext As String
    Dim i As Long
    Dim Desc As String

    Fid = Replace(Fid, vbQuote, "")

    i = InStrRev(Fid, ".")
    If i <= 0 Then
        UpdateStatus "No file name extension -- Cannot Verify"
        ApiBeep
        Exit Sub
    End If

    Ext = Mid(Fid, i)
    VerifyHashType = HashExt2Type(Ext)
    If VerifyHashType = HASH_NONE Then
'        If StrComp(ext, ".txt", vbTextCompare) <> 0 Then
            UpdateStatus "Unrecognized file name extention (" & Ext & ") -- Cannot Verify"
            ApiBeep
            Exit Sub
'        End If
    End If
    Desc = HashDesc(VerifyHashType)

    Dim fn As Long

    fn = FreeFile()
    On Error Resume Next
    Open Fid For Input Access Read As fn
    If Err.Number <> 0 Then
        UpdateStatus "Could not open hash file: " & Fid
        ApiBeep
        Err.Clear
        Exit Sub
    End If

    i = InStrRev(Fid, "\")
    If i > 0 Then
        VerifyPath = Left(Fid, i)
    End If

    Dim s As String
    Dim sa() As String
    Dim LinCnt As Long
    Dim ColHdr As ColumnHeader
    Dim LstItm As ListItem
    Dim Cnt As Long
    Dim LineValid As Boolean

    LinCnt = 0
    Do While Not EOF(fn)
        Line Input #fn, s
        If InStr(1, s, vbTab) > 0 Then
            VerifyDelimiter = vbTab
            sa = Split(s, vbTab)
            LinCnt = LinCnt + 1
        ElseIf InStr(1, s, " ") > 0 Then
            VerifyDelimiter = " "
            sa = Split(s, VerifyDelimiter, 2)
            LinCnt = LinCnt + 1
        End If
        If LinCnt > 1 Then Exit Do
    Loop

    If LinCnt <= 0 Then
        UpdateStatus "Invalid Hash File -- Cannot Verify"
        Close #fn
        Exit Sub
    End If

    If LinCnt = 1 Then ' might be group
        If UBound(sa) = 1 Then ' one file or folder
            If Right(sa(1), 1) <> "\" Then ' one file
                VerifyGroup = False
            Else
                VerifyGroup = True
            End If
        Else
            VerifyGroup = True
        End If
    Else ' individual file(s)
        VerifyGroup = False
    End If

    lblVerifyResult.Caption = ""
    lvVerifyIn.ListItems.Clear
    lvVerifyIn.Sorted = False
    lvVerifyIn.ColumnHeaders.Clear

    Set ColHdr = lvVerifyIn.ColumnHeaders.Add() ' dummy
    ColHdr.Width = 0

    Set ColHdr = lvVerifyIn.ColumnHeaders.Add() ' file/folder name
    ColHdr.Text = "File or Folder Name"
    ColHdr.Width = 0

    Set ColHdr = lvVerifyIn.ColumnHeaders.Add() ' status
    ColHdr.Alignment = lvwColumnCenter
    ColHdr.Width = 0
    ColHdr.Text = "Status"

    btnVerifyCalc.Enabled = False
    lblVerifyOldHash.Enabled = False
    txtVerifyOldHash.Enabled = False
    lblVerifyNewHash.Enabled = False
    txtVerifyNewHash.Enabled = False
    txtVerifyOldHash.Text = ""
    txtVerifyNewHash.Text = ""
    btnVerifyUpdate.Enabled = False

    VerifyRelativePath = True
    If VerifyGroup Then
        picVerifyPrompt.Visible = False
        lvVerifyIn.ColumnHeaders(COL_VERIFY_NAME + 1).Width = 100
        Cnt = 0
        Seek #fn, 1
        Do While Not EOF(fn)
            Line Input #fn, s
            If InStr(1, s, vbTab) > 0 Then
                sa = Split(s, vbTab)
                Close #fn
                Exit Do
            ElseIf InStr(1, s, VerifyDelimiter) > 0 Then
                sa = Split(s, VerifyDelimiter, 2)
                Close #fn
                Exit Do
            End If
        Loop
        txtVerifyOldHash.Text = sa(0)
        For i = 1 To UBound(sa)
            Set LstItm = lvVerifyIn.ListItems.Add()
            If InStr(1, sa(i), "\") > 0 Then
                LstItm.SubItems(COL_VERIFY_NAME) = sa(i)
                VerifyRelativePath = False
            Else
                LstItm.SubItems(COL_VERIFY_NAME) = VerifyPath & sa(i)
            End If
            Cnt = Cnt + 1
            If (Cnt Mod 10) = 1 Then
                LvAdjust lvVerifyIn
                DoEvents
            End If
        Next i
        If lvVerifyIn.ListItems.Count > 0 Then
            lblVerifyOldHash.Enabled = True
            txtVerifyOldHash.Enabled = True
        End If
        s = Replace("Original %%HASHTYPE%% Hash Value of Group:", "%%HASHTYPE%%", Desc, , , vbTextCompare)
        lblVerifyOldHash.Caption = s
        s = Replace("Current %%HASHTYPE%% Hash Value of Group:", "%%HASHTYPE%%", Desc, , , vbTextCompare)
        lblVerifyNewHash.Caption = s

    Else ' individual files/folders
        picVerifyPrompt.Visible = False
        lvVerifyIn.ColumnHeaders(COL_VERIFY_NAME + 1).Width = 100

        Set ColHdr = lvVerifyIn.ColumnHeaders.Add()
        Set ColHdr = lvVerifyIn.ColumnHeaders.Add()
        
        lvVerifyIn.ColumnHeaders(COL_VERIFY_NEW_HASH + 1).Text = "Current " & Desc & " Hash"
        lvVerifyIn.ColumnHeaders(COL_VERIFY_NEW_HASH + 1).Width = 0
        lvVerifyIn.ColumnHeaders(COL_VERIFY_NEW_HASH + 1).Alignment = lvwColumnCenter
        
        lvVerifyIn.ColumnHeaders(COL_VERIFY_OLD_HASH + 1).Text = "Original " & Desc & " Hash"
        lvVerifyIn.ColumnHeaders(COL_VERIFY_OLD_HASH + 1).Width = 100
        lvVerifyIn.ColumnHeaders(COL_VERIFY_OLD_HASH + 1).Alignment = lvwColumnCenter

        Cnt = 0
        Seek #fn, 1
        Do While Not EOF(fn)
            Line Input #fn, s
            If InStr(1, s, vbTab) > 0 Then
                sa = Split(s, vbTab)
                LineValid = True
            ElseIf InStr(1, s, VerifyDelimiter) > 0 Then
                sa = Split(s, VerifyDelimiter, 2)
                LineValid = True
            Else
                LineValid = False
            End If
            If LineValid Then
                Set LstItm = lvVerifyIn.ListItems.Add()
                If InStr(1, sa(1), "\") > 0 Then
                    LstItm.SubItems(COL_VERIFY_NAME) = sa(1)
                    VerifyRelativePath = False
                Else
                    LstItm.SubItems(COL_VERIFY_NAME) = VerifyPath & sa(1)
                End If
                LstItm.SubItems(COL_VERIFY_OLD_HASH) = sa(0)
                Cnt = Cnt + 1
                If (Cnt Mod 10) = 1 Then
                    LvAdjust lvVerifyIn
                    DoEvents
                End If
            End If
        Loop
        Close #fn
        s = "Original Hash Value of Group:"
        lblVerifyOldHash.Caption = s
        s = "Current Hash Value of Group:"
        lblVerifyNewHash.Caption = s
    End If

    If lvVerifyIn.ListItems.Count > 0 Then
        s = "Verify " & Desc & " Hash"
        If lvVerifyIn.ListItems.Count > 1 Then s = s & "es"
        btnVerifyCalc.Caption = s
        btnVerifyCalc.Enabled = True
    End If

    If lvVerifyIn.ListItems.Count > 0 Then
        picVerifyPrompt.Visible = False
    Else
        picVerifyPrompt.Visible = True
    End If

    LvAdjust lvVerifyIn

    lvVerifyIn.SortKey = COL_VERIFY_NAME
    lvVerifyIn.SortOrder = lvwAscending
    lvVerifyIn.Sorted = True
End Sub
Private Sub btnVerifyUpdate_Click()
    Dim yorn As VbMsgBoxResult

    If VerifyFailed Then
        yorn = MsgBox("One or more hash values have changed, or could not be computed." & vbCrLf & vbCrLf & "Do you want to update the hash file with the current hash values," & vbCrLf & "and delete entries for any files or folders that could not be processed?", vbQuestion Or vbYesNo Or vbDefaultButton2, App.FileDescription)
        If yorn <> vbYes Then Exit Sub
    End If

    Dim fn As Long

    fn = FreeFile()
    On Error Resume Next
    Open VerifyFid For Output Access Write As fn
    If Err.Number <> 0 Then
        UpdateStatus "Could not open Hash File: " & VerifyFid
        ApiBeep
        Exit Sub
    End If

    Dim s As String
    Dim i As Long
    Dim j As Long
    Dim lim As Long
    Dim LstItm As ListItem

    lim = lvVerifyIn.ListItems.Count
    If VerifyGroup Then
        Print #fn, txtVerifyNewHash.Text;
        For i = 1 To lim
            Set LstItm = lvVerifyIn.ListItems(i)
            Print #fn, vbTab & LstItm.SubItems(COL_VERIFY_NAME);
        Next i
        Print fn, ""
    Else
        For i = 1 To lim
            Set LstItm = lvVerifyIn.ListItems(i)
            If HashLegal(LstItm.SubItems(COL_VERIFY_NEW_HASH), VerifyHashType) Then
                Print #fn, LstItm.SubItems(COL_VERIFY_NEW_HASH) & VerifyDelimiter;
                If VerifyRelativePath Then
                    s = LstItm.SubItems(COL_VERIFY_NAME)
                    j = InStrRev(s, "\")
                    If j > 0 Then
                        Print #fn, Mid(s, j + 1)
                    Else
                        Print #fn, s
                    End If
                Else
                    Print #fn, LstItm.SubItems(COL_VERIFY_NAME)
                End If
            End If
        Next i
    End If

    Close fn
    UpdateStatus "Hash File Updated"
    ApiBeep
    VerifyParse VerifyFid
End Sub
Private Sub UpdateStatus(Msg As String)
    Status.Text = Msg
    Status.ToolTipText = Msg
End Sub
Private Sub cboHashFav_Click()
    HashFav = cboHashFav.ItemData(cboHashFav.ListIndex)
    If HashFav <> HASH_NONE Then chkAvail(HashFav).Value = vbChecked
End Sub
Private Sub chkAvail_Click(Index As Integer)
    Static busy As Boolean

    If GlobalBusy Then Exit Sub
    GlobalBusy = True

    If chkAvail(Index).Value = vbUnchecked Then
        If cboHashFav.ItemData(cboHashFav.ListIndex) = Index Then cboHashFav.ListIndex = 0
        If fraAvail.Visible Then
            If chkExt(Index).Value = vbChecked Then
                chkExt(Index).Value = vbUnchecked
                UpdateStatus "File Association with " & HashType2Ext(CLng(Index)) & " Removed"
                ApiBeep
            End If
        End If
    End If
    GlobalBusy = False
End Sub
Private Sub chkExt_Click(Index As Integer)
    Static busy As Boolean

    If GlobalBusy Then Exit Sub
    GlobalBusy = True

    If fraAssoc.Visible Then
        If chkAvail(Index).Value = vbUnchecked Then
            chkAvail(Index).Value = vbChecked
            UpdateStatus "Algorithm " & HashDesc(CLng(Index)) & " made available"
            ApiBeep
        End If
    End If
    GlobalBusy = False
End Sub
Private Sub Form_Initialize()
    ApiInitCommonControls
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        If ApiHelpEnabled Then
            DispHelp
        Else
            frmAbout.Show vbModal
        End If
    End If
End Sub
Private Sub Form_Load()
    Dim result As Long
    Dim sig() As Byte
    Dim Ctrl As Control
    Dim i As Long
    Dim inx As Long
    Dim FirstAvail As HASH_TYPE

    BACK_NORMAL = vbWindowBackground
    BACK_GOOD = RGB(192, 255, 192)
    BACK_WARN = RGB(255, 255, 128)
    BACK_ERROR = RGB(255, 192, 192)

    Set Reg = New Registry
    Set Status = sbMain.Panels(1)
    Set HashPanel = frmMain.sbMain.Panels(1)
    Me.Caption = App.FileDescription
    Reg.Home

    On Error Resume Next
'    Intialize
    If Not Hasher.Present Then
        MsgBox Hasher.Banner, vbOKOnly Or vbCritical, App.FileDescription
        Unload Me
    End If

    lblVersion.Caption = LoadResString(USING_MSG) & Hasher.Banner

    MinWidth = Me.Width
    MinHeight = Me.Height
    RightMargin = Me.Width - (btnExit.Left + btnExit.Width)
    BottomMargin = (Me.ScaleHeight - sbMain.Height) - (btnExit.Top + btnExit.Height)
    HorzGap = btnExit.Left - (btnAbout.Left + btnAbout.Width)
    VertGap = fraTextCalc.Top - (fraTextIn.Top + fraTextIn.Height)
    BlockGap = tabText.Top - (fraTextIn.Top + fraTextIn.Height)

    Reg.GetFormSize Me
    Reg.GetFormPos Me
    lblWelcome.Caption = LoadResString(WELCOME) & App.FileDescription & "!"
    lblInfo.Caption = Replace(App.Comments, "\n", vbCrLf)
    lblTextPrompt.Caption = Replace(lblTextPrompt.Caption, "\n", vbCrLf)
    lblTextCalc.Caption = Replace(lblTextCalc.Caption, "\n", vbCrLf)
    lblFilePrompt.Caption = Replace(lblFilePrompt.Caption, "\n", vbCrLf)
    lblFileCalc.Caption = Replace(lblFileCalc.Caption, "\n", vbCrLf)
    lblFilesPrompt.Caption = Replace(lblFilesPrompt.Caption, "\n", vbCrLf)
    lblGroupCalc.Caption = Replace(lblGroupCalc.Caption, "\n", vbCrLf)

    ApiFormCaption Me
    On Error Resume Next
    ApiFormFont Me
    tabText.Font.Size = 8
    tabFile.Font.Size = 8
    tabGroup.Font.Size = 8
    lblVersion.FontSize = 8
    For i = txtTextCalc.LBound To txtTextCalc.UBound
        txtTextCalc(i).FontName = "Courier New"
        txtFilesCalc(i).FontName = "Courier New"
    Next i
    Err.Clear
    DoEvents

    cboHashFav.AddItem "None/Last Used"
    cboHashFav.ItemData(cboHashFav.NewIndex) = HASH_NONE
    FirstAvail = HASH_NONE
    For i = HASH_TYPE_LBOUND To HASH_TYPE_UBOUND
        chkAvail(i).Value = Reg.ReadValue("HashAvail " & HashDesc(i), vbChecked)
        If chkAvail(i).Value = vbChecked Then
'            tabText.Tabs(i).HighLighted = True
'            tabFile.Tabs(i).HighLighted = True
'            tabGroup.Tabs(i).HighLighted = True
            If FirstAvail = HASH_NONE Then FirstAvail = i
        Else
'            tabText.Tabs(i).HighLighted = False
'            tabFile.Tabs(i).HighLighted = False
'            tabGroup.Tabs(i).HighLighted = False
        End If
        cboHashFav.AddItem HashDesc(i)
        cboHashFav.ItemData(cboHashFav.NewIndex) = i
    Next i

    If FirstAvail = HASH_NONE Then ' no algorithms available
        For i = chkAvail.LBound To chkAvail.UBound
            chkAvail(i).Value = vbChecked
        Next i
        FirstAvail = HASH_SHA1
    End If

    HashFav = Reg.ReadValue("HashFav", HASH_SHA1)
    If HashFav <> HASH_NONE Then chkAvail(HashFav).Value = vbChecked

    For i = 0 To cboHashFav.ListCount - 1
        If cboHashFav.ItemData(i) = HashFav Then
            cboHashFav.ListIndex = i
            Exit For
        End If
    Next i

    For i = HASH_TYPE_LBOUND To HASH_TYPE_UBOUND
        If Reg.FileAssocOwned(HashType2Ext(i)) Then
            chkExt(i).Value = vbChecked
            chkAvail(i).Value = vbChecked
        Else
            chkExt(i).Value = vbUnchecked
        End If
    Next i
    lblAssoc.Caption = Replace(lblAssoc.Caption, "%%PROGRAM%%", App.FileDescription, , , vbTextCompare)

    If HashFav = HASH_NONE Then
        Set tabText.SelectedItem = tabText.Tabs(Reg.ReadValue("HashLast Text", FirstAvail))
        Set tabFile.SelectedItem = tabFile.Tabs(Reg.ReadValue("HashLast File", FirstAvail))
        Set tabGroup.SelectedItem = tabGroup.Tabs(Reg.ReadValue("HashLast Files", FirstAvail))
    Else
        Set tabText.SelectedItem = tabText.Tabs(HashFav)
        Set tabFile.SelectedItem = tabFile.Tabs(HashFav)
        Set tabGroup.SelectedItem = tabGroup.Tabs(HashFav)
    End If

    PrevShowNetwork = Reg.ReadValue("PrevShowNetwork", PrevShowNetwork)
    PrevShowFiles = Reg.ReadValue("PrevShowFiles", PrevShowFiles)
    PrevSubfolders = Reg.ReadValue("PrevSubfolders", PrevSubfolders)
    PrevFolder = Reg.ReadValue("PrevFolder", PrevFolder)

    LvAdjust lvFileIn

    If Not ApiHelpEnabled Then btnHelp.Enabled = False

    If ApiFileExists(App.Path & "\EnableTest.txt") Then
        lblSep2.Visible = True
        btnTest.Visible = True
        lstTest.Visible = True
    End If

    If SettingsOK() Then
        If Len(Command) > 0 Then
            VerifyParse Command
            Set tabHash.SelectedItem = tabHash.Tabs(TAB_VERIFY)
        Else
            Set tabHash.SelectedItem = tabHash.Tabs(TAB_WELCOME)
            UpdateStatus LoadResString(READY_MSG)
        End If
    Else
        Set tabHash.SelectedItem = tabHash.Tabs(TAB_SETTINGS)
        UpdateStatus LoadResString("Invalid Settings")
    End If
End Sub
Private Sub Form_Resize()
    Dim Para As Long
    Dim i As Long

    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < MinWidth Then
        Me.Width = MinWidth
        Exit Sub
    End If
    If Me.Height < MinHeight Then
        Me.Height = MinHeight
        Exit Sub
    End If

    btnExit.Left = Me.Width - RightMargin - btnExit.Width
    btnAbout.Left = btnExit.Left - btnAbout.Width - HorzGap
    btnHelp.Left = btnAbout.Left - btnHelp.Width - HorzGap
    tabHash.Width = Me.Width - RightMargin - tabHash.Left
    Para = tabHash.Width - (2 * tabHash.Left)
    For i = TAB_LBOUND To TAB_UBOUND
        picHash(i).Width = Para
    Next i

    Para = (Me.ScaleHeight - sbMain.Height) - BottomMargin - btnExit.Height
    btnExit.Top = Para
    btnAbout.Top = Para
    btnHelp.Top = Para
    btnCopy.Top = Para
    btnSave.Top = Para

    tabHash.Height = btnExit.Top - BottomMargin - tabHash.Top
    Para = tabHash.Height - picHash(1).Top
    For i = TAB_LBOUND To TAB_UBOUND
        picHash(i).Height = Para
    Next i

    Para = tabHash.Height - (fraTextCalc.Height + 600)
    fraTextCalc.Top = Para
    fraFileCalc.Top = Para
    fraGroupCalc.Top = Para
    fraVerifyHash.Top = Para

    Para = tabHash.Width - (FRAME_BORDER * 2)
    fraTextIn.Width = Para
    fraTextCalc.Width = Para
    fraFileIn.Width = Para
    fraFileCalc.Width = Para
    fraGroupIn.Width = Para
    fraGroupCalc.Width = Para
    fraVerifyIn.Width = Para
    fraVerifyHash.Width = Para

    Para = fraTextIn.Width - (picTextIn.Left + FRAME_BORDER)
    picTextIn.Width = Para
    picFileIn.Width = Para
    picGroupIn.Width = Para
    picVerifyIn.Width = Para

    Para = picTextIn.Width - (txtTextIn.Left + btnPaste.Left)
    txtTextIn.Width = Para
    lvFileIn.Width = Para
    lvGroupIn.Width = Para
    lvVerifyIn.Width = Para
    txtVerifyNewHash.Width = Para
    txtVerifyOldHash.Width = Para

    Para = fraTextCalc.Width - PAD2
    tabText.Width = Para
    tabFile.Width = Para
    tabGroup.Width = Para
    picVerifyHash.Width = Para

    Para = Para - PAD2
    For i = picFileHash.LBound To picFileHash.UBound
        picTextHash(i).Width = Para
        picFileHash(i).Width = Para
        picFilesHash(i).Width = Para
    Next i

    Para = tabText.Width - (btnTextCalc(1).Width + PAD3)
    For i = txtTextCalc.LBound To txtTextCalc.UBound
        txtTextCalc(i).Width = Para
        txtFilesCalc(i).Width = Para
    Next i

    Para = fraTextCalc.Top - (fraTextIn.Top + VertGap)
    fraTextIn.Height = Para
    fraFileIn.Height = Para
    fraGroupIn.Height = Para
    fraVerifyIn.Height = Para

    Para = fraTextIn.Height - (picTextIn.Top + FRAME_BORDER)
    picTextIn.Height = Para
    picFileIn.Height = Para
    picGroupIn.Height = Para
    picVerifyIn.Height = Para

    Para = picTextIn.Height - txtTextIn.Top
    txtTextIn.Height = Para
    lvFileIn.Height = Para
    lvGroupIn.Height = Para
    lvVerifyIn.Height = Para

    lblInfo.Width = tabHash.Width - (tabHash.Left * 4)
    lblWelcome.Width = lblInfo.Width

    btnFileDelAll.Top = (lvFileIn.Top + lvFileIn.Height) - btnFileDelAll.Height
    btnFileDelSel.Top = btnFileDelAll.Top - PAD - btnFileDelSel.Height

    btnGroupDelAll.Top = (lvGroupIn.Top + lvFileIn.Height) - btnGroupDelAll.Height
    btnGroupDelSel.Top = btnGroupDelAll.Top - PAD - btnGroupDelSel.Height

    lblVersion.Top = picHash(1).Height - lblVersion.Height '- PAD
    LvAdjust lvFileIn
    LvAdjust lvVerifyIn
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long

    Reg.SaveFormSize Me
    Reg.SaveFormPos Me
    For i = HASH_TYPE_LBOUND To HASH_TYPE_UBOUND
        Reg.WriteValue "HashAvail " & HashDesc(i), chkAvail(i).Value
    Next i
    Reg.WriteValue "HashFav", HashFav
    Reg.WriteValue "HashLast Text", tabText.SelectedItem.Index
    Reg.WriteValue "HashLast File", tabFile.SelectedItem.Index
    Reg.WriteValue "HashLast Files", tabGroup.SelectedItem.Index

    Reg.WriteValue "PrevShowNetwork", PrevShowNetwork
    Reg.WriteValue "PrevShowFiles", PrevShowFiles
    Reg.WriteValue "PrevSubfolders", PrevSubfolders
    Reg.WriteValue "PrevFolder", PrevFolder

    End
End Sub
Private Sub lblFilesPrompt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GroupContext Button, Shift, X, Y
End Sub
Private Sub lblTextCalc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picTextCalc.Visible = False
    txtTextCalc(tabText.SelectedItem.Index).SetFocus
End Sub
Private Sub lblTextPrompt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picTextPrompt.Visible = False
    txtTextIn.SetFocus
End Sub
Private Sub lblVerifyPrompt_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.Files.Count > 0 Then
        VerifyParse Data.Files(1)
    End If
End Sub
Private Sub lblWelcomeFile_Click()
    Set tabHash.SelectedItem = tabHash.Tabs(TAB_FILE)
End Sub
Private Sub lblWelcomeGroup_Click()
    Set tabHash.SelectedItem = tabHash.Tabs(TAB_GROUP)
End Sub
Private Sub lblWelcomeText_Click()
    Set tabHash.SelectedItem = tabHash.Tabs(TAB_TEXT)
End Sub
Private Sub lblWelcomeVerify_Click()
    Set tabHash.SelectedItem = tabHash.Tabs(TAB_VERIFY)
End Sub
Private Sub lvFileIn_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lvFileIn.ToolTipText = Item.Text
    FileEnable
End Sub
Private Sub lvFileIn_LostFocus()
    If lvFileIn.ListItems.Count > 0 Then
        picFilePrompt.Visible = False
    Else
        picFilePrompt.Visible = True
    End If
End Sub
Private Sub lvFileIn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FileContext Button, Shift, X, Y
End Sub
Private Sub FileContext(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim SelCnt As Long
    Dim i As Long

    If (Button And vbRightButton) = 0 Then Exit Sub

    SelCnt = 0
    For i = 1 To lvFileIn.ListItems.Count
        If lvFileIn.ListItems(i).Selected Then
            SelCnt = SelCnt + 1
            Exit For
        End If
    Next i

    If SelCnt > 0 Then
        mnuFileDelSel.Enabled = True
        mnuFileCopySel.Enabled = True
        mnuFileCopyFileSel.Enabled = True
    Else
        mnuFileDelSel.Enabled = False
        mnuFileCopySel.Enabled = False
        mnuFileCopyFileSel.Enabled = False
    End If
    If lvFileIn.ListItems.Count > 0 Then
        mnuFileDelAll.Enabled = True
        mnuFileCopyAll.Enabled = True
        mnuFileCopyFileAll.Enabled = True
    Else
        mnuFileDelAll.Enabled = False
        mnuFileCopyAll.Enabled = False
        mnuFileCopyFileAll.Enabled = False
    End If
    PopupMenu mnuFile
End Sub
Private Sub lvFileIn_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim Fids As New Collection

    On Error Resume Next
    For i = 1 To Data.Files.Count
        Fids.Add Data.Files(i)
    Next i

    FileDrop Fids
End Sub
Private Sub FileDrop(Fids As Collection)
    Dim i As Long
    Dim j As Long
    Dim Fid As String
    Dim yorn As VbMsgBoxResult
    Dim mbstyle As VbMsgBoxStyle
    Dim LstItm As ListItem
    Dim FoundOne As Boolean
    Dim fa As PT_FID_ARRAY
    Dim Recurse As YORN_VALUE
    Dim lim As Long
    Dim lim2 As Long

    lvFileIn.Sorted = False
    Recurse = YORN_NONE

    On Error Resume Next
    For i = 1 To Fids.Count
        Fid = Fids.Item(i)
        If (GetAttr(Fid) And vbDirectory) = 0 Then
            If Err.Number = 0 Then
                Set LstItm = lvFileIn.ListItems.Add()
                LstItm.Text = Fid
                LstItm.Selected = True
            Else
                Err.Clear
                If i = Fids.Count Then
                    mbstyle = vbOKOnly
                Else
                    mbstyle = vbOKCancel
                End If
                yorn = MsgBox(LoadResString(CANNOT_ACCESS) & vbCrLf & LoadResString(NOT_ADDED), mbstyle Or vbInformation, App.FileDescription)
                If yorn = vbCancel Then
                    UpdateStatus LoadResString(DRAG_AND_DROP_CANCELLED)
                    If lvFileIn.ListItems.Count > 0 Then
                        picFilePrompt.Visible = False
                    Else
                        picFilePrompt.Visible = True
                    End If
                    Exit Sub
                End If
            End If
        Else ' it's a folder
            If Right(Fid, 1) <> "\" Then Fid = Fid & "\"
            FoundOne = False
            fa = ApiFindFolders(Fid, "*.*")
            If fa.FidCnt > 0 Then
                lim2 = fa.FidCnt - 1
                For j = 0 To lim2
                    If Recurse = YORN_NONE Then
                        Load frmYesNo
                        frmYesNo.Question = Replace(LoadResString(INCLUDE_SUBFOLDERS_PARA), "%%FOLDER%%", Fid, , , vbTextCompare)
                        frmYesNo.Title = LoadResString(INCLUDE_SUBFOLDERS_PARA)
                        frmYesNo.Show vbModal
                        Recurse = frmYesNo.Answer
                    End If
                    If (Recurse And YORN_YES) = YORN_YES Then
                        FileAddFolder Fid, True
                    Else
                        FileAddFolder Fid, False
                    End If
                    If (Recurse And YORN_ALL) = 0 Then Recurse = YORN_NONE
                    FoundOne = True
                Next j
                If Not FoundOne Then FileAddFolder Fid, False
            End If
        End If
        DoEvents
    Next i

    lvFileIn.SortKey = 0
    lvFileIn.SortOrder = lvwAscending
    lvFileIn.Sorted = True

    If lvFileIn.ListItems.Count > 0 Then
        picFilePrompt.Visible = False
        LvAdjust lvFileIn
        ArrayEnable btnFileCalcAll, False
    Else
        picFilePrompt.Visible = True
    End If

    FileMarkDirty True
    FileEnable
End Sub
Private Sub lvGroupIn_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lvGroupIn.ToolTipText = Item.Text
    GroupEnable
End Sub
Private Sub lvGroupIn_LostFocus()
    If lvGroupIn.ListItems.Count > 0 Then
        picFilesPrompt.Visible = False
    Else
        picFilesPrompt.Visible = True
    End If
End Sub
Private Sub lvGroupIn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GroupContext Button, Shift, X, Y
End Sub
Private Sub GroupContext(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim SelCnt As Long
    Dim i As Long

    If (Button And vbRightButton) = 0 Then Exit Sub

    SelCnt = 0
    For i = 1 To lvGroupIn.ListItems.Count
        If lvGroupIn.ListItems(i).Selected Then
            SelCnt = SelCnt + 1
            Exit For
        End If
    Next i

    If SelCnt > 0 Then
        mnuFilesDelSel.Enabled = True
        mnuFilesCopyFileAll.Enabled = True
        mnuFilesCopyFileSel.Enabled = True
    Else
        mnuFilesDelSel.Enabled = False
        mnuFilesCopyFileAll.Enabled = False
        mnuFilesCopyFileSel.Enabled = False
    End If

    If lvGroupIn.ListItems.Count > 0 Then
        mnuFilesDelAll.Enabled = True
        mnuFilesCopyAll.Enabled = True
        mnuFilesCopyFileAll.Enabled = True
    Else
        mnuFilesDelAll.Enabled = False
        mnuFilesCopyAll.Enabled = False
        mnuFilesCopyFileAll.Enabled = False
    End If
    PopupMenu mnuFiles
End Sub
Private Sub lvGroupIn_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim Fids As New Collection

    On Error Resume Next
    For i = 1 To Data.Files.Count
        Fids.Add Data.Files(i)
    Next i

    GroupDrop Fids
End Sub
Private Sub GroupDrop(Fids As Collection)
    Dim i As Long
    Dim j As Long
    Dim Fid As String
    Dim yorn As VbMsgBoxResult
    Dim mbstyle As VbMsgBoxStyle
    Dim LstItm As ListItem
    Dim FoundOne As Boolean
    Dim fa As PT_FID_ARRAY
    Dim Recurse As YORN_VALUE
    Dim lim As Long
    Dim lim2 As Long

    lvGroupIn.Sorted = False
    Recurse = YORN_NONE

    On Error Resume Next
    For i = 1 To Fids.Count
        Fid = Fids.Item(i)
        If (GetAttr(Fid) And vbDirectory) = 0 Then
            If Err.Number = 0 Then ' it's a file
                Set LstItm = lvGroupIn.ListItems.Add()
                LstItm.Text = Fid
           Else
                Err.Clear
                If i = Fids.Count Then
                    mbstyle = vbOKOnly
                Else
                    mbstyle = vbOKCancel
                End If
                yorn = MsgBox(LoadResString(CANNOT_ACCESS) & vbCrLf & LoadResString(NOT_ADDED), mbstyle Or vbInformation, App.FileDescription)
                If yorn = vbCancel Then
                    UpdateStatus LoadResString(DRAG_AND_DROP_CANCELLED)
                    Exit Sub
                End If
            End If
        Else ' it's a folder
            If Right(Fid, 1) <> "\" Then Fid = Fid & "\"
            FoundOne = False
            fa = ApiFindFolders(Fid, "*.*")
            If fa.FidCnt > 0 Then
                lim2 = fa.FidCnt - 1
                For j = 0 To lim2
                    If Recurse = YORN_NONE Then
                        Load frmYesNo
                        frmYesNo.Question = Replace(LoadResString(INCLUDE_SUBFOLDERS_PARA), "%%FOLDER%%", Fid, , , vbTextCompare)
                        frmYesNo.Title = LoadResString(INCLUDE_SUBFOLDERS_PARA)
                        frmYesNo.Show vbModal
                        Recurse = frmYesNo.Answer
                    End If
                    If (Recurse And YORN_YES) = YORN_YES Then
                        FilesAddFolder Fid, True
                    Else
                        FilesAddFolder Fid, False
                    End If
                    If (Recurse And YORN_ALL) = 0 Then Recurse = YORN_NONE
                    FoundOne = True
                Next j
                If Not FoundOne Then FilesAddFolder Fid, False
            End If
        End If
        DoEvents
    Next i

    lvGroupIn.SortKey = 0
    lvGroupIn.SortOrder = lvwAscending
    lvGroupIn.Sorted = True

    If lvGroupIn.ListItems.Count > 0 Then
        LvAdjust lvGroupIn
        picFilesPrompt.Visible = False
    Else
        picFilesPrompt.Visible = True
    End If

    GroupMarkDirty True
    GroupEnable
End Sub
Private Sub lvVerifyIn_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.Files.Count > 0 Then
        VerifyParse Data.Files(1)
    End If
End Sub
Private Sub mnuFileCopyAll_Click()
    FileCopy
End Sub
Private Sub mnuFileCopyFileAll_Click()
    Dim s As String
    Dim fn As Long
    Dim Hash As String
    Dim ColHdr As ColumnHeader
    Dim LstItm As ListItem
    Dim lim As Long
    Dim Col As Long
    Dim Row As Long

    lim = lvFileIn.ListItems.Count

    s = CopyHeader() & vbCrLf
    s = s & Replace(Trim(fraFileIn.Caption), ": ", ":" & vbTab) & vbCrLf
    s = s & vbCrLf
    s = s & "File Name"
'    For Col = COL_FIRST_HASH To COL_LAST_HASH
'        Set ColHdr = lvFileIn.ColumnHeaders(Col + 1)
'        If ColHdr.Width > 0 Then s = s & vbTab & ColHdr.Text
'    Next Col

    s = s & vbCrLf
    For Row = 1 To lim
        Set LstItm = lvFileIn.ListItems(Row)
'        If LstItm.Selected Then
            s = s & LstItm.Text
'            For Col = COL_FIRST_HASH To COL_LAST_HASH
'                Set ColHdr = lvFileIn.ColumnHeaders(Col + 1)
'                If ColHdr.Width > 0 Then
'                    If HashLegal(LstItm.SubItems(Col),Col - COL_FIRST_HASH + 1) Then
'                        s = s & vbTab & LstItm.SubItems(Col)
'                    Else
'                        s = s & vbTab & resstring(NOT_COMPUTED)
'                    End If
'                End If
'            Next Col
            s = s & vbCrLf
'        End If
    Next Row

    Clipboard.Clear
    Clipboard.SetText Trim(s)
End Sub
Private Sub mnuFileCopyFileSel_Click()
    Dim s As String
    Dim fn As Long
    Dim Hash As String
    Dim ColHdr As ColumnHeader
    Dim LstItm As ListItem
    Dim lim As Long
    Dim Col As Long
    Dim Row As Long

    lim = lvFileIn.ListItems.Count

    s = CopyHeader() & vbCrLf
    s = s & Replace(Trim(fraFileIn.Caption), ": ", ":" & vbTab) & vbCrLf
    s = s & vbCrLf
    s = s & "File Name"
'    For Col = COL_FIRST_HASH To COL_LAST_HASH
'        Set ColHdr = lvFileIn.ColumnHeaders(Col + 1)
'        If ColHdr.Width > 0 Then s = s & vbTab & ColHdr.Text
'    Next Col

    s = s & vbCrLf
    For Row = 1 To lim
        Set LstItm = lvFileIn.ListItems(Row)
        If LstItm.Selected Then
            s = s & LstItm.Text
'            For Col = COL_FIRST_HASH To COL_LAST_HASH
'                Set ColHdr = lvFileIn.ColumnHeaders(Col + 1)
'                If ColHdr.Width > 0 Then
'                    If HashLegal(LstItm.SubItems(Col), Col - COL_FIRST_HASH + 1) Then
'                        s = s & vbTab & LstItm.SubItems(Col)
'                    Else
'                        s = s & vbTab & resstring(NOT_COMPUTED)
'                    End If
'                End If
'            Next Col
            s = s & vbCrLf
        End If
    Next Row

    Clipboard.Clear
    Clipboard.SetText Trim(s)
End Sub
Private Sub mnuFileCopySel_Click()
    Dim s As String
    Dim fn As Long
    Dim Hash As String
    Dim ColHdr As ColumnHeader
    Dim LstItm As ListItem
    Dim lim As Long
    Dim Col As Long
    Dim Row As Long

    lim = lvFileIn.ListItems.Count

    s = CopyHeader() & vbCrLf
    s = s & Replace(Trim(fraFileIn.Caption), ": ", ":" & vbTab) & vbCrLf
    s = s & vbCrLf
    s = s & "File Name"
    For Col = COL_FIRST_HASH To COL_LAST_HASH
        Set ColHdr = lvFileIn.ColumnHeaders(Col + 1)
        If ColHdr.Width > 0 Then s = s & vbTab & ColHdr.Text
    Next Col

    s = s & vbCrLf
    For Row = 1 To lim
        Set LstItm = lvFileIn.ListItems(Row)
        If LstItm.Selected Then
            s = s & LstItm.Text
            For Col = COL_FIRST_HASH To COL_LAST_HASH
                Set ColHdr = lvFileIn.ColumnHeaders(Col + 1)
                If ColHdr.Width > 0 Then
                    If HashLegal(LstItm.SubItems(Col), Col - COL_FIRST_HASH + 1) Then
                        s = s & vbTab & LstItm.SubItems(Col)
                    Else
                        s = s & vbTab & LoadResString(NOT_COMPUTED)
                    End If
                End If
            Next Col
            s = s & vbCrLf
        End If
    Next Row

    Clipboard.Clear
    Clipboard.SetText Trim(s)
End Sub
Private Sub mnuFileDelAll_Click()
    Dim i As Long
    Dim Cnt As Long

    For i = lvFileIn.ListItems.Count To 1 Step -1
        lvFileIn.ListItems.Remove i
        Cnt = Cnt + 1
    Next i

    For i = 2 To lvFileIn.ColumnHeaders.Count
        lvFileIn.ColumnHeaders(i).Width = 0
    Next i

    If Cnt > 0 Then
        UpdateStatus FormatNumber(Cnt, 0) & LoadResString(FILES_REMOVED)
    End If

    If lvFileIn.ListItems.Count > 0 Then
        picFilePrompt.Visible = False
    Else
        picFilePrompt.Visible = True
    End If

    lvFileIn.ToolTipText = LoadResString(FILE_TIP1) & Replace(btnFileBrowse.Caption, "&", "") & LoadResString(FILE_TIP2)
    FileMarkDirty True
    FileEnable
End Sub
Private Sub mnuFileDelSel_Click()
    Dim i As Long
    Dim Cnt As Long

    For i = lvFileIn.ListItems.Count To 1 Step -1
        If lvFileIn.ListItems(i).Selected Then
            lvFileIn.ListItems.Remove i
            Cnt = Cnt + 1
        End If
    Next i

    If Cnt > 0 Then
        UpdateStatus FormatNumber(Cnt, 0) & LoadResString(FILES_REMOVED)
    End If

    If lvFileIn.ListItems.Count > 0 Then
        picFilePrompt.Visible = False
    Else
        picFilePrompt.Visible = True
    End If

    lvFileIn.ToolTipText = LoadResString(FILE_TIP1) & Replace(btnFileBrowse.Caption, "&", "") & LoadResString(FILE_TIP2)
    FileMarkDirty True
    FileEnable
End Sub
Private Sub mnuFilesCopyAll_Click()
    GroupCopy
End Sub
Private Sub mnuFilesCopyFileAll_Click()
    Dim s As String
    Dim fn As Long
    Dim Hash As String
    Dim ColHdr As ColumnHeader
    Dim LstItm As ListItem
    Dim lim As Long
    Dim Col As Long
    Dim Row As Long

    lim = lvGroupIn.ListItems.Count

    s = CopyHeader() & vbCrLf
    s = s & Replace(Trim(fraGroupIn.Caption), ": ", ":" & vbTab) & vbCrLf
    s = s & vbCrLf
    s = s & "File Name"
'    For Col = COL_FIRST_HASH To COL_LAST_HASH
'        Set ColHdr = lvGroupIn.ColumnHeaders(Col + 1)
'        If ColHdr.Width > 0 Then s = s & vbTab & ColHdr.Text
'    Next Col

    s = s & vbCrLf
    For Row = 1 To lim
        Set LstItm = lvGroupIn.ListItems(Row)
'        If LstItm.Selected Then
            s = s & LstItm.Text
'            For Col = COL_FIRST_HASH To COL_LAST_HASH
'                Set ColHdr = lvGroupIn.ColumnHeaders(Col + 1)
'                If ColHdr.Width > 0 Then
'                    If HashLegal(LstItm.SubItems(Col), Col - COL_FIRST_HASH + 1) Then
'                        s = s & vbTab & LstItm.SubItems(Col)
'                    Else
'                        s = s & vbTab & resstring(NOT_COMPUTED)
'                    End If
'                End If
'            Next Col
            s = s & vbCrLf
'        End If
    Next Row

    Clipboard.Clear
    Clipboard.SetText Trim(s)
End Sub
Private Sub mnuFilesCopyFileSel_Click()
    Dim s As String
    Dim fn As Long
    Dim Hash As String
    Dim ColHdr As ColumnHeader
    Dim LstItm As ListItem
    Dim lim As Long
    Dim Col As Long
    Dim Row As Long

    lim = lvGroupIn.ListItems.Count

    s = CopyHeader() & vbCrLf
    s = s & Replace(Trim(fraGroupIn.Caption), ": ", ":" & vbTab) & vbCrLf
    s = s & vbCrLf
    s = s & "File Name"
'    For Col = COL_FIRST_HASH To COL_LAST_HASH
'        Set ColHdr = lvGroupIn.ColumnHeaders(Col + 1)
'        If ColHdr.Width > 0 Then s = s & vbTab & ColHdr.Text
'    Next Col

    s = s & vbCrLf
    For Row = 1 To lim
        Set LstItm = lvGroupIn.ListItems(Row)
        If LstItm.Selected Then
            s = s & LstItm.Text
'            For Col = COL_FIRST_HASH To COL_LAST_HASH
'                Set ColHdr = lvGroupIn.ColumnHeaders(Col + 1)
'                If ColHdr.Width > 0 Then
'                    If HashLegal(LstItm.SubItems(Col), Col - COL_FIRST_HASH + 1) Then
'                        s = s & vbTab & LstItm.SubItems(Col)
'                    Else
'                        s = s & vbTab & resstring(NOT_COMPUTED)
'                    End If
'                End If
'            Next Col
            s = s & vbCrLf
        End If
    Next Row

    Clipboard.Clear
    Clipboard.SetText Trim(s)
End Sub
Private Sub mnuFilesDelAll_Click()
    Dim i As Long
    Dim Cnt As Long

    For i = lvGroupIn.ListItems.Count To 1 Step -1
        lvGroupIn.ListItems.Remove i
        Cnt = Cnt + 1
    Next i

    If Cnt > 0 Then
        UpdateStatus FormatNumber(Cnt, 0) & LoadResString(FILES_REMOVED)
    End If

    If lvGroupIn.ListItems.Count > 0 Then
        picFilesPrompt.Visible = False
    Else
        picFilesPrompt.Visible = True
    End If

    lvGroupIn.ToolTipText = LoadResString(FILE_TIP1) & Replace(btnGroupBrowse.Caption, "&", "") & LoadResString(FILE_TIP2)
    GroupMarkDirty True
    GroupEnable
End Sub
Private Sub mnuFilesDelSel_Click()
    Dim i As Long
    Dim Cnt As Long
    Dim SelCnt As Long

    SelCnt = 0
    For i = 1 To lvGroupIn.ListItems.Count
        If lvGroupIn.ListItems(i).Selected Then
            SelCnt = SelCnt + 1
            Exit For
        End If
    Next i
    If SelCnt < 0 Then Exit Sub

    For i = lvGroupIn.ListItems.Count To 1 Step -1
        If lvGroupIn.ListItems(i).Selected Then
            lvGroupIn.ListItems.Remove i
            Cnt = Cnt + 1
        End If
    Next i

    If Cnt > 0 Then
        UpdateStatus FormatNumber(Cnt, 0) & LoadResString(FILES_REMOVED)
    End If

    If lvGroupIn.ListItems.Count > 0 Then
        picFilesPrompt.Visible = False
    Else
        picFilesPrompt.Visible = True
    End If

    lvGroupIn.ToolTipText = LoadResString(FILE_TIP1) & Replace(btnGroupBrowse.Caption, "&", "") & LoadResString(FILE_TIP2)
    GroupMarkDirty True
    GroupEnable
End Sub
'Private Sub mnuFolderDelAll_Click()
'    Dim i As Long
'    Dim Cnt As Long
'
'    For i = lvFolderIn.ListItems.Count To 1 Step -1
'        lvFolderIn.ListItems.Remove i
'        Cnt = Cnt + 1
'    Next i
'
'    If Cnt > 0 Then
'        UpdateStatus FormatNumber(Cnt, 0) & LoadResString(FOLDERS_REMOVED)
'    End If
'
'    If lvFolderIn.ListItems.Count > 0 Then
'        picFolderPrompt.Visible = False
'    Else
'        picFolderPrompt.Visible = True
'    End If
'
'    lvFolderIn.ToolTipText = LoadResString(FOLDER_TIP1) & Replace(btnFolderBrowse.Caption, "&", "") & LoadResString(FILE_TIP2)
'    FolderDirty = True
'    FolderEnable
'End Sub
'Private Sub mnuFolderDelSel_Click()
'    Dim i As Long
'    Dim Cnt As Long
'
'    For i = lvFolderIn.ListItems.Count To 1 Step -1
'        If lvFolderIn.ListItems(i).Selected Then
'            lvFolderIn.ListItems.Remove i
'            Cnt = Cnt + 1
'        End If
'    Next i
'
'    For i = 2 To lvFileIn.ColumnHeaders.Count
'        lvFileIn.ColumnHeaders(i).Width = 0
'    Next i
'
'    If Cnt > 0 Then
'        UpdateStatus FormatNumber(Cnt, 0) & LoadResString(FOLDERS_REMOVED)
'    End If
'
'    If lvFolderIn.ListItems.Count > 0 Then
'        picFolderPrompt.Visible = False
'    Else
'        picFolderPrompt.Visible = True
'    End If
'
'    lvFolderIn.ToolTipText = LoadResString(FOLDER_TIP1) & Replace(btnFolderBrowse.Caption, "&", "") & LoadResString(FOLDER_TIP2)
'    FolderDirty = True
'    FolderEnable
'End Sub
Private Sub picFilePrompt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FileContext Button, Shift, X, Y
End Sub
Private Sub picFilePrompt_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim Fids As New Collection

    On Error Resume Next
    For i = 1 To Data.Files.Count
        Fids.Add Data.Files(i)
    Next i

    FileDrop Fids
End Sub
Private Sub picFilesPrompt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GroupContext Button, Shift, X, Y
End Sub
Private Sub picFilesPrompt_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim Fids As New Collection

    On Error Resume Next
    For i = 1 To Data.Files.Count
        Fids.Add Data.Files(i)
    Next i

    GroupDrop Fids
End Sub
'Private Sub picFolderPrompt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    FolderContext Button, Shift, X, Y
'End Sub
Private Sub picTextCalc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picTextCalc.Visible = False
    txtTextCalc(tabText.SelectedItem.Index).SetFocus
End Sub
Private Sub picTextPrompt_Click()
    picTextPrompt.Visible = False
    txtTextIn.SetFocus
End Sub
Private Sub picVerifyPrompt_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Data.Files.Count > 0 Then
        VerifyParse Data.Files(1)
    End If
End Sub
Private Sub tabFile_Click()
    Dim Col As Long
    Dim inx As Long

    inx = CStr(tabFile.SelectedItem.Tag)
    Col = inx + COL_FIRST_HASH
    If (lvFileIn.ColumnHeaders(Col).Width > 0) And (Not FileDirty) Then
        picFileCalc.Visible = False
    Else
        picFileCalc.Visible = True
    End If
'    If btnFileCalcAll(Inx).Enabled Then btnFileCalcAll(Inx).Default = True

    Dim i As Long

    For i = HASH_LBOUND To HASH_UBOUND
        If i = inx Then
            picFileHash(i).Visible = True
        Else
            picFileHash(i).Visible = False
        End If
    Next i
End Sub
Private Sub tabGroup_Click()
    Dim i As Long
    Dim inx As Long

    inx = CStr(tabGroup.SelectedItem.Tag)
    For i = HASH_LBOUND To HASH_UBOUND
        If i = inx Then
            picFilesHash(i).Visible = True
        Else
            picFilesHash(i).Visible = False
        End If
    Next i

    If Len(txtFilesCalc(tabGroup.SelectedItem.Index).Text) > 0 Then
        picGroupCalc.Visible = False
    Else
        picGroupCalc.Visible = True
    End If
End Sub
Private Function SettingsOK() As Boolean
    Dim i As Long
    Dim inx As Long
    Dim FirstAvail As HASH_TYPE
    Dim Cnt As Long
    Dim FavInx As Long
    Dim Ext As String

    SettingsOK = False

    For i = chkExt.LBound To chkExt.UBound
        Ext = HashType2Ext(i)
        If chkExt(i).Value = vbChecked Then
            If Not Reg.FileAssocOwned(Ext) Then
                Reg.FileAssocAdd Ext, HashDesc(i) & " Hash File"
            End If
            chkAvail(i).Value = vbChecked
        Else
            If Reg.FileAssocOwned(Ext) Then
                Reg.FileAssocDel Ext
            End If
        End If
    Next i

    FirstAvail = HASH_NONE
    Cnt = 0
    For i = chkAvail.LBound To chkAvail.UBound
        If chkAvail(i).Value = vbChecked Then
            If FirstAvail = HASH_NONE Then FirstAvail = i
            Cnt = Cnt + 1
        End If
    Next i

    If (FirstAvail = HASH_NONE) Or (Cnt <= 0) Then
        PreviousTab = TAB_WELCOME
        Set tabHash.SelectedItem = tabHash.Tabs(TAB_SETTINGS)
        MsgBox LoadResString(NEED_ONE_HASH), vbOKOnly Or vbExclamation, App.FileDescription
        Exit Function
    End If

    Do While tabText.Tabs.Count > Cnt
        tabText.Tabs.Remove 1
        tabFile.Tabs.Remove 1
        tabGroup.Tabs.Remove 1
    Loop

    Do While tabText.Tabs.Count < Cnt
        tabText.Tabs.Add
        tabFile.Tabs.Add
        tabGroup.Tabs.Add
    Loop

    inx = 1
    FavInx = -1
    For i = HASH_TYPE_LBOUND To HASH_TYPE_UBOUND
        If chkAvail(i).Value = vbChecked Then
            tabText.Tabs(inx).Caption = LoadResString(HASH_TAB_BASE + i)
            tabText.Tabs(inx).Tag = CStr(i)
            tabFile.Tabs(inx).Caption = LoadResString(HASH_TAB_BASE + i)
            tabFile.Tabs(inx).Tag = CStr(i)
            tabGroup.Tabs(inx).Caption = LoadResString(HASH_TAB_BASE + i)
            tabGroup.Tabs(inx).Tag = CStr(i)
            If i = HashFav Then FavInx = inx
            inx = inx + 1
        End If
    Next i

    If FavInx <= 0 Then HashFav = HASH_NONE
    If HashFav = HASH_NONE Then
        Set tabText.SelectedItem = tabText.Tabs(1)
        Set tabFile.SelectedItem = tabFile.Tabs(1)
        Set tabGroup.SelectedItem = tabGroup.Tabs(1)
    Else
        Set tabText.SelectedItem = tabText.Tabs(FavInx)
        Set tabFile.SelectedItem = tabFile.Tabs(FavInx)
        Set tabGroup.SelectedItem = tabGroup.Tabs(FavInx)
    End If

    SettingsOK = True
End Function
Private Sub lblFilePrompt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FileContext Button, Shift, X, Y
End Sub
Private Sub lblFilePrompt_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim Fids As New Collection

    On Error Resume Next
    For i = 1 To Data.Files.Count
        Fids.Add Data.Files(i)
    Next i

    FileDrop Fids
End Sub
Private Sub lblFilesPrompt_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim Fids As New Collection

    On Error Resume Next
    For i = 1 To Data.Files.Count
        Fids.Add Data.Files(i)
    Next i

    GroupDrop Fids
End Sub
Private Sub tabHash_Click()
    If PreviousTab = TAB_SETTINGS Then
        If Not SettingsOK() Then Exit Sub
    End If

    Dim i As Long
    Dim inx As Long

    inx = tabHash.SelectedItem.Index
    For i = TAB_LBOUND To TAB_UBOUND
        If i = inx Then
            picHash(i).Top = 600
            picHash(i).Left = 240
            picHash(i).Visible = True
        Else
            picHash(i).Visible = False
        End If
    Next i

    Select Case inx
        Case TAB_TEXT
            TextEnable
        Case TAB_FILE
            FileEnable
        Case TAB_GROUP
            GroupEnable
        Case TAB_VERIFY
            VerifyEnable
        Case Else
            btnCopy.Enabled = False
            btnSave.Enabled = False
    End Select

    PreviousTab = tabHash.SelectedItem.Index
End Sub
Private Sub tabText_Click()
    If Len(txtTextCalc(tabText.SelectedItem.Index).Text) > 0 Then
        picTextCalc.Visible = False
    Else
        picTextCalc.Visible = True
    End If
'    If btnTextCalc(tabText.Tab).Enabled Then btnTextCalc(tabText.Tab).Default = True

    Dim i As Long
    Dim inx As Long

    inx = CStr(tabText.SelectedItem.Tag)

    For i = HASH_LBOUND To HASH_UBOUND
        If i = inx Then
            picTextHash(i).Visible = True
        Else
            picTextHash(i).Visible = False
        End If
    Next i
End Sub
Private Sub txtTextIn_Change()
    TextMarkDirty True
End Sub
Private Sub TextMarkDirty(b As Boolean)
    Dim i As Long

    If b Then
        TextDirty = True
        For i = txtTextCalc.LBound To txtTextCalc.UBound
            txtTextCalc(i).Text = ""
        Next i
    Else
        TextDirty = False
    End If
End Sub
Private Sub txtTextIn_GotFocus()
    picTextPrompt.Visible = False
End Sub
Private Sub txtTextIn_LostFocus()
    If Len(txtTextIn.Text) <= 0 Then picTextPrompt.Visible = True
End Sub
