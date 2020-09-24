VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Advanced Browse"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2085
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   2025
      ScaleWidth      =   1545
      TabIndex        =   0
      Top             =   120
      Width           =   1605
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Website: http://Software.TW-Chat.co.uk"
      Height          =   195
      Left            =   1920
      TabIndex        =   4
      Top             =   1320
      Width           =   2910
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Email: Wayne@TW-Chat.co.uk"
      Height          =   195
      Left            =   1920
      TabIndex        =   3
      Top             =   1080
      Width           =   2235
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Caption         =   "Version x.x.x"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Advanced Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   2565
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAbout = Nothing
End Sub
