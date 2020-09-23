VERSION 5.00
Begin VB.Form frmBrowse 
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5520
   Icon            =   "Browse.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnHelp 
      Caption         =   "&Help ..."
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   4560
      Width           =   975
   End
   Begin VB.CheckBox chkSubFolders 
      Caption         =   "Add sub-folders too"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   2175
   End
   Begin Hasher.FolderBrowser brw 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   7646
      Caption         =   "Select a &Folder:"
      Speedy          =   -1  'True
      ShowHidden      =   -1  'True
      ShowSystem      =   -1  'True
      ShowReadOnly    =   -1  'True
      ShowLocal       =   -1  'True
      ShowNetwork     =   0   'False
      ShowNetworkVisible=   -1  'True
      ShowFilesVisible=   -1  'True
      NewFolderVisible=   0   'False
      DetailsVisible  =   -1  'True
      Object.Height          =   4335
      Object.Width           =   5295
      Border          =   0   'False
      BackColor       =   -2147483633
      Appearance      =   0
      Object.Visible         =   -1  'True
      Enabled         =   -1  'True
      ShowFiles       =   0   'False
      FileFilter      =   ""
      FolderFilter    =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   4560
      Width           =   975
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Copyright © 2005 Karen Kenworthy
' All Rights Reserved
' http://www.karenware.com/
' Version 1.0 7/11/2005

Public Cancelled As Boolean
Private mSelFolder As String

Private Const PAD = 120
Private Const PAD2 = PAD * 2

Private MinWidth As Single
Private MinHeight As Single
Public Property Get ShowFiles() As Boolean
    ShowFiles = brw.ShowFiles
End Property
Public Property Let ShowFiles(ByVal NewValue As Boolean)
    brw.ShowFiles = NewValue
End Property
Public Property Get ShowNetwork() As Boolean
    ShowNetwork = brw.ShowNetwork
End Property
Public Property Let ShowNetwork(ByVal NewValue As Boolean)
    brw.ShowNetwork = NewValue
End Property
Public Property Get AddSubfolders() As Boolean
    If chkSubFolders.Value = vbChecked Then
        AddSubfolders = True
    Else
        AddSubfolders = False
    End If
End Property
Public Property Let AddSubfolders(ByVal NewValue As Boolean)
    If NewValue Then
        chkSubFolders.Value = vbChecked
    Else
        chkSubFolders.Value = vbUnchecked
    End If
End Property
Public Property Get SelFolder() As String
    SelFolder = mSelFolder
End Property
Public Property Let SelFolder(ByVal NewFolder As String)
    brw.Folder = NewFolder
End Property
Private Sub brw_FolderSelection(Folder As String)
    If Len(Folder) > 0 Then
        btnOK.Enabled = True
    Else
        btnOK.Enabled = False
    End If

    mSelFolder = Folder
End Sub
Private Sub btnCancel_Click()
    Cancelled = True
    Me.Hide
End Sub
Private Sub btnHelp_Click()
    ApiHelpTopic 6000
End Sub
Private Sub btnOK_Click()
    Cancelled = False
    Me.Hide
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        If ApiHelpEnabled Then
            ApiHelpTopic 6000
        End If
    End If
End Sub
Private Sub Form_Load()
    Me.Caption = App.FileDescription & " - Select Folder"
    ApiFormFont Me

    MinWidth = Me.Width
    MinHeight = Me.Height

    mSelFolder = ""

    If Len(brw.Folder) > 0 Then
        brw_FolderSelection brw.Folder
    Else
        btnOK.Enabled = False
    End If

    btnHelp.Visible = ApiHelpEnabled()
End Sub
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub

    If Me.Width < MinWidth Then
        Me.Width = MinWidth
        Exit Sub
    End If
    If Me.Height < MinHeight Then
        Me.Height = MinHeight
        Exit Sub
    End If

    brw.Width = Me.ScaleWidth - brw.Left - PAD
    btnOK.Left = Me.ScaleWidth - btnOK.Width - PAD
    btnCancel.Left = btnOK.Left - btnCancel.Width - PAD
    btnHelp.Left = btnCancel.Left - btnHelp.Width - PAD

    btnOK.Top = Me.ScaleHeight - btnOK.Height - PAD
    btnCancel.Top = btnOK.Top
    btnHelp.Top = btnOK.Top
    chkSubFolders.Top = btnOK.Top

    brw.Height = btnOK.Top - brw.Top - PAD
End Sub

