VERSION 5.00
Begin VB.Form frmOpen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open URL"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "http://"
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public frmOwner As Form

Private Sub cmdOpen_Click()
    frmOwner.WebBrowser1.Navigate txtURL.Text
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmOwner = Nothing
End Sub

Private Sub txtURL_Change()
    Select Case LCase$(txtURL.Text)
        Case "http://", "https://", "ftp://", Empty:
            cmdOpen.Enabled = False
            
        Case Else:
            cmdOpen.Enabled = True
            
    End Select
End Sub

Private Sub txtURL_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) And (cmdOpen.Enabled = True) Then _
        Call cmdOpen_Click
End Sub
