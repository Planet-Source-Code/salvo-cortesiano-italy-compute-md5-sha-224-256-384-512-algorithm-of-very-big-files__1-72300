VERSION 5.00
Begin VB.Form frmYesNo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Subfolders?"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   Icon            =   "YesNo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnNo2All 
      Caption         =   "N&o To All"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton btnNo 
      Caption         =   "&No"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton btnYes2All 
      Caption         =   "Y&es To All"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton btnYes 
      Caption         =   "&Yes"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblQuestion 
      BackStyle       =   0  'Transparent
      Caption         =   "Do you want to add subfolders of folder ""%FOLDER%""?"
      Height          =   1095
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "YesNo.frx":212A
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmYesNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Copyright © 2005 Karen Kenworthy
' All Rights Reserved
' http://www.karenware.com/

Public Enum YORN_VALUE
    YORN_NONE = 0
    YORN_ALL = &H10
    YORN_YES = &H1
    YORN_YES2ALL = YORN_YES Or YORN_ALL
    YORN_NO = &H2
    YORN_NO2ALL = YORN_NO Or YORN_ALL
End Enum

Private mAnswer As YORN_VALUE
Private mTitle As String
Private mQuestion As String
Public Property Get Title() As String
    Title = mTitle
End Property
Public Property Let Title(ByVal NewTitle As String)
    mTitle = NewTitle
    Me.Caption = mTitle
End Property
Public Property Get Question() As String
    Question = mQuestion
End Property
Public Property Let Question(ByVal NewQuestion As String)
    mQuestion = NewQuestion
    lblQuestion.Caption = mQuestion
End Property
Public Property Get Answer() As YORN_VALUE
    Answer = mAnswer
End Property
Public Property Let Answer(ByVal DefaultAnswer As YORN_VALUE)
    mAnswer = DefaultAnswer
    Select Case DefaultAnswer
        Case YORN_YES
            btnYes.Default = True
        Case YORN_YES2ALL
            btnYes2All.Default = True
        Case YORN_NO
            btnNo.Default = True
        Case YORN_NO2ALL
            btnNo2All.Default = True
        Case Else
            btnYes.Default = False
            btnYes2All.Default = False
            btnNo.Default = False
            btnNo2All.Default = False
            mAnswer = YORN_NONE
    End Select
End Property
Private Sub btnNo_Click()
    mAnswer = YORN_NO
    Me.Hide
End Sub
Private Sub btnNo2All_Click()
    mAnswer = YORN_NO2ALL
    Me.Hide
End Sub
Private Sub btnYes_Click()
    mAnswer = YORN_YES
    Me.Hide
End Sub
Private Sub btnYes2All_Click()
    mAnswer = YORN_YES2ALL
    Me.Hide
End Sub
Private Sub Form_Load()
    ApiFormFont Me
    mQuestion = lblQuestion.Caption
    mTitle = Me.Caption
    Answer = YORN_NONE
End Sub

