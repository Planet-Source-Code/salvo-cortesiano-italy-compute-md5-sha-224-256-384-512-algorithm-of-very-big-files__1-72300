VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Power Tool"
   ClientHeight    =   5715
   ClientLeft      =   1185
   ClientTop       =   1485
   ClientWidth     =   7425
   ForeColor       =   &H80000008&
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5715
   ScaleWidth      =   7425
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.TextBox txtExtra 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.TextBox txtComments 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1440
      Width           =   6975
   End
   Begin VB.CommandButton btnOK 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   6240
      TabIndex        =   11
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label lblSepExtra 
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Height          =   75
      Left            =   180
      TabIndex        =   12
      Top             =   3120
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.Image imgLogo 
      Height          =   960
      Left            =   240
      Picture         =   "About.frx":212A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   960
   End
   Begin VB.Label lblSep 
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Height          =   75
      Left            =   180
      TabIndex        =   4
      Top             =   3120
      Width           =   7095
   End
   Begin VB.Label lblVersion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Version X.X"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   825
   End
   Begin VB.Label lblCDLink 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.karenware.com/cd.asp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      MouseIcon       =   "About.frx":26F9
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "Get Karen's Power Tools on CD!"
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Label lblCDInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"About.frx":284B
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   3360
      Width           =   6975
   End
   Begin VB.Label lblSubLink 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.karenware.com/subscribe/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      MouseIcon       =   "About.frx":291E
      MousePointer    =   99  'Custom
      TabIndex        =   8
      ToolTipText     =   "Subscribe to Karen's FREE newsletter"
      Top             =   4560
      Width           =   5175
   End
   Begin VB.Label lblSubInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Subscribe to Karen's FREE newsletter, and be the first to know of new and upgraded Power Tools: "
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   4320
      Width           =   6975
   End
   Begin VB.Label lblHomeLink 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.karenware.com/powertools/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   240
      MouseIcon       =   "About.frx":2A70
      MousePointer    =   99  'Custom
      TabIndex        =   10
      ToolTipText     =   "Visit Karen's home page"
      Top             =   5280
      Width           =   4215
   End
   Begin VB.Label lblHomeInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "For more information, and the latest version, visit Karen's Power Tools home page: "
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   5040
      Width           =   5775
   End
   Begin VB.Label lblCopyright 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 1993, 1994, 1999 Karen Kenworthy All Rights Reserved"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   4890
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Power Tool Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   1440
      TabIndex        =   0
      Top             =   180
      Width           =   3195
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Copyright © 2002-2005, 2007 Karen Kenworthy
' All Rights Reserved
' http://www.karenware.com/
' Version 2.14 5/31/2007

Private Const SW_NORMAL = 1
Private Const SW_SHOW = 5
Private Const SHADOW_OFFSET = 30

Private Const ABOUT_APPID = "About"

Private Declare Function FindExecutableA Lib "Shell32.dll" ( _
    ByVal lpFile As String, _
    ByVal lpDirectory As String, _
    ByVal lpResult As String) As Long

Private Declare Function ShellExecuteA Lib "Shell32.dll" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Public AboutExtra As String

Private Reg As Registry
Private CDInfo As String
Private cdLink As String
Private SubInfo As String
Private SubLink As String
Private HomeInfo As String
Private HomeLink As String
Private Sub btnOK_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Dim ver As String
    Dim s As String
    Dim i As Long

    txtExtra.Visible = False
    lblSepExtra.Visible = False

    CDInfo = Trim(lblCDInfo.Caption) & "  "
    cdLink = Trim(lblCDLink.Caption)
    SubInfo = Trim(lblSubInfo.Caption) & "  "
    SubLink = Trim(lblSubLink.Caption)
    HomeInfo = Trim(lblHomeInfo.Caption) & "  "
    HomeLink = Trim(lblHomeLink.Caption)

    lblTitle.Caption = App.FileDescription & " "

    ver = App.Major & "." & App.Minor
    If App.Revision <> 0 Then
        ver = ver & "." & App.Revision
    End If

    lblVersion.Caption = "Version " & ver
    lblVersion.Top = lblTitle.Top + lblTitle.Height

    s = Replace(App.LegalCopyright, "\n", vbCrLf, , , vbTextCompare)
    lblCopyright.Caption = Replace(s, "(c)", "©")
    lblCopyright.Top = lblVersion.Top + lblVersion.Height

    txtComments.Text = Replace(App.Comments, "\n", vbCrLf, , , vbTextCompare)

    Me.Caption = "About " & App.FileDescription

    ApiFormFont Me
End Sub
Private Sub Form_Resize()
    Static busy As Boolean

    If busy Then Exit Sub
    busy = True
    LayoutAbout
    busy = False
End Sub
Private Sub AddExtra()
    If Len(AboutExtra) > 0 Then
        If lblSepExtra.Visible Then Exit Sub
    Else
        Exit Sub
    End If

    Dim xySize As PT_TEXT_SIZE
    Dim Delta As Single

'    xySize = ApiTextSize(Me.hdc, AboutExtra)
    Delta = Me.TextHeight(AboutExtra)
    If Delta > 0 Then
        txtExtra.Height = Delta
        Delta = Delta + (lblCDInfo.Top - lblSep.Top)
        Me.Height = Me.Height + Delta
        btnOK.Top = btnOK.Top + Delta
'        txtSepExtra.Top = lblSep.Top + Delta
'        txtExtra.Top = lblSep.Top + (lblCDInfo.Top + lblSep.Top)
        txtExtra.Text = AboutExtra
        lblSepExtra.Visible = True
        txtExtra.Visible = True
    End If
End Sub
Private Sub LayoutAbout()
    Dim taInfo As PT_TEXT_ARRAY
    Dim xySize As PT_TEXT_SIZE
    Dim DiffVert As Single
    Dim DiffHorz As Single

    Dim NewCdInfoTop As Single
    Dim NewCdInfoHeight As Single
    Dim NewCdInfoWidth As Single
    Dim NewCdInfoCaption As String
    Dim NewCdLinkTop As Single
    Dim NewCdLinkLeft As Single
    Dim NewCdLinkHeight As Single
    Dim NewCdLinkWidth As Single

    Dim NewSubInfoTop As Single
    Dim NewSubInfoHeight As Single
    Dim NewSubInfoWidth As Single
    Dim NewSubInfoCaption As String
    Dim NewSubLinkTop As Single
    Dim NewSubLinkLeft As Single
    Dim NewSubLinkHeight As Single
    Dim NewSubLinkWidth As Single

    Dim NewHomeInfoTop As Single
    Dim NewHomeInfoHeight As Single
    Dim NewHomeInfoWidth As Single
    Dim NewHomeInfoCaption As String
    Dim NewHomeLinkTop As Single
    Dim NewHomeLinkLeft As Single
    Dim NewHomeLinkHeight As Single
    Dim NewHomeLinkWidth As Single

    Dim NewWidth As Single
    Dim BottomMargin As Single
    Dim RightMargin As Single
    Dim LeftMargin As Single
    Dim CurBase As Single

    AddExtra

    BottomMargin = Me.Height - (btnOK.Top + btnOK.Height)
    LeftMargin = txtComments.Left
    RightMargin = Me.Width - (btnOK.Left + btnOK.Width)
    NewWidth = Me.Width - (LeftMargin + RightMargin)
    NewWidth = txtComments.Width

    DiffVert = Me.Height - (btnOK.Top + btnOK.Height)
    DiffVert = DiffVert - BottomMargin

    DiffHorz = Me.Width - (btnOK.Left + btnOK.Width)
    DiffHorz = DiffHorz - RightMargin
    DoEvents

'    If (DiffVert = 0) And (DiffHorz = 0) Then Exit Sub
    CurBase = btnOK.Top + btnOK.Height

    NewHomeInfoWidth = NewWidth - (btnOK.Width + RightMargin)
    taInfo = ApiTextWrap(Me.hdc, NewHomeInfoWidth, HomeInfo)
    NewHomeInfoCaption = taInfo.WrappedText
    NewHomeInfoHeight = taInfo.Overall.HeightTwips
    NewHomeInfoTop = CurBase - NewHomeInfoHeight
    xySize = ApiTextSize(Me.hdc, HomeLink)
    NewHomeLinkHeight = xySize.HeightTwips
    NewHomeLinkWidth = xySize.WidthTwips
    If (taInfo.LastLine.WidthTwips + NewHomeLinkWidth) > NewHomeInfoWidth Then
        NewHomeInfoTop = NewHomeInfoTop - NewHomeLinkHeight
        NewHomeLinkTop = CurBase - NewHomeLinkHeight
        NewHomeLinkLeft = LeftMargin
    Else
        NewHomeLinkTop = (NewHomeInfoTop + NewHomeInfoHeight) - taInfo.LastLine.HeightTwips
        NewHomeLinkLeft = LeftMargin + taInfo.LastLine.WidthTwips
    End If
    DoEvents

    CurBase = NewHomeInfoTop - (NewHomeLinkHeight * 1#)

    NewSubInfoWidth = NewWidth
    taInfo = ApiTextWrap(Me.hdc, NewSubInfoWidth, SubInfo)
    NewSubInfoCaption = taInfo.WrappedText
    NewSubInfoHeight = taInfo.Overall.HeightTwips
    NewSubInfoTop = CurBase - NewSubInfoHeight
    xySize = ApiTextSize(Me.hdc, SubLink)
    NewSubLinkHeight = xySize.HeightTwips
    NewSubLinkWidth = xySize.WidthTwips
    If (taInfo.LastLine.WidthTwips + NewSubLinkWidth) > NewSubInfoWidth Then
        NewSubInfoTop = NewSubInfoTop - NewSubLinkHeight
        NewSubLinkTop = CurBase - NewSubLinkHeight
        NewSubLinkLeft = LeftMargin
    Else
        NewSubLinkTop = (NewSubInfoTop + NewSubInfoHeight) - taInfo.LastLine.HeightTwips
        NewSubLinkLeft = LeftMargin + taInfo.LastLine.WidthTwips
    End If
    DoEvents

    CurBase = NewSubInfoTop - (NewSubLinkHeight * 1#)

    NewCdInfoWidth = NewWidth
    taInfo = ApiTextWrap(Me.hdc, NewCdInfoWidth, CDInfo)
    NewCdInfoCaption = taInfo.WrappedText
    NewCdInfoHeight = taInfo.Overall.HeightTwips
    NewCdInfoTop = CurBase - NewCdInfoHeight
    xySize = ApiTextSize(Me.hdc, cdLink)
    NewCdLinkHeight = xySize.HeightTwips
    NewCdLinkWidth = xySize.WidthTwips
    If (taInfo.LastLine.WidthTwips + NewCdLinkWidth) > NewCdInfoWidth Then
        NewCdInfoTop = NewCdInfoTop - NewCdLinkHeight
        NewCdLinkTop = CurBase - NewCdLinkHeight
        NewCdLinkLeft = LeftMargin
    Else
        NewCdLinkTop = (NewCdInfoTop + NewCdInfoHeight) - taInfo.LastLine.HeightTwips
        NewCdLinkLeft = LeftMargin + taInfo.LastLine.WidthTwips
    End If
    DoEvents

    CurBase = NewCdInfoTop - (NewCdLinkHeight * 0.5)

    If lblSepExtra.Visible Then
        lblSepExtra.Top = CurBase - lblSepExtra.Height
        lblSepExtra.Width = NewWidth
        CurBase = lblSepExtra.Top - (NewCdLinkHeight * 0.5)
        txtExtra.Top = CurBase - txtExtra.Height
        CurBase = txtExtra.Top - (NewCdLinkHeight * 0.5)
    End If

    lblSep.Top = CurBase - lblSep.Height
    lblSep.Width = NewWidth

    CurBase = lblSep.Top - (NewCdLinkHeight * 0.5)

    txtComments.Height = CurBase - txtComments.Top
    txtComments.Width = NewWidth

    lblHomeInfo.Top = NewHomeInfoTop
    lblHomeInfo.Height = NewHomeInfoHeight
    lblHomeInfo.Width = NewHomeInfoWidth
    lblHomeInfo.Caption = NewHomeInfoCaption

    lblHomeLink.Top = NewHomeLinkTop
    lblHomeLink.Left = NewHomeLinkLeft
    lblHomeLink.Height = NewHomeLinkHeight
    lblHomeLink.Width = NewHomeLinkWidth
    lblHomeLink.ZOrder

    lblSubInfo.Top = NewSubInfoTop
    lblSubInfo.Height = NewSubInfoHeight
    lblSubInfo.Width = NewSubInfoWidth
    lblSubInfo.Caption = NewSubInfoCaption

    lblSubLink.Top = NewSubLinkTop
    lblSubLink.Left = NewSubLinkLeft
    lblSubLink.Height = NewSubLinkHeight
    lblSubLink.Width = NewSubLinkWidth
    lblSubLink.ZOrder

    lblCDInfo.Top = NewCdInfoTop
    lblCDInfo.Height = NewCdInfoHeight
    lblCDInfo.Width = NewCdInfoWidth
    lblCDInfo.Caption = NewCdInfoCaption

    lblCDLink.Top = NewCdLinkTop
    lblCDLink.Left = NewCdLinkLeft
    lblCDLink.Height = NewCdLinkHeight
    lblCDLink.Width = NewCdLinkWidth
    lblCDLink.ZOrder
End Sub
Private Sub lblCDLink_Click()
    ShellExecuteA Me.hwnd, "open", lblCDLink.Caption, "", "", SW_SHOW Or SW_NORMAL
End Sub
Private Sub lblHomeLink_Click()
    ShellExecuteA Me.hwnd, "open", lblHomeLink.Caption, "", "", SW_SHOW Or SW_NORMAL
End Sub
Private Sub lblSubLink_Click()
    ShellExecuteA Me.hwnd, "open", lblSubLink.Caption, "", "", SW_SHOW Or SW_NORMAL
End Sub
Private Sub txtExtra_DblClick()
    Dim s As String
    Dim i As Long

    i = InStr(3, txtExtra.Text, ":")
    If i > 0 Then
        s = Trim(Mid(txtExtra.Text, i + 1))
    Else
        s = txtExtra.Text
    End If
    
    On Error Resume Next
    ShellExecuteA Me.hwnd, "open", s, vbNullString, vbNullString, SW_SHOW Or SW_NORMAL
    Err.Clear
End Sub
