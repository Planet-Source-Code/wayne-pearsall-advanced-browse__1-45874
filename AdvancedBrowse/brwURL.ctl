VERSION 5.00
Begin VB.UserControl brwURL 
   Alignable       =   -1  'True
   ClientHeight    =   1110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7125
   ScaleHeight     =   1110
   ScaleWidth      =   7125
   Begin VB.ComboBox cboAddress 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7095
   End
   Begin VB.Label lblCapt 
      AutoSize        =   -1  'True
      Caption         =   "Address:"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "brwURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Control Created By Wayne Pearsall
' With Code Found At VBForums.Com

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                         ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' used in the auto complete
Private Const CB_FINDSTRING = &H14C
Private Const CB_ERR = (-1)
Public blnDelete As Boolean

' Needed to show the drop down
'Private Const CB_SHOWDROPDOWN = &H14F
' -------------------------------------------------------------------
'
' -------------------------------------------------------------------
'
' -------------------------------------------------------------------

Public Event Click()
Public Event KeyPress(KeyAscii As Integer)

' -------------------------------------------------------------------
'
' -------------------------------------------------------------------
'
' -------------------------------------------------------------------

Public Sub AddItem(Item As String)

    With cboAddress
        .Text = Item
        
        Dim i As Integer
        For i = 0 To .ListCount - 1
            If .List(i) = Item Then
                .RemoveItem i
                Exit For
            End If
            DoEvents
        Next i
        
        .AddItem Item, 0
        .Text = Item
        .SelStart = Len(.Text)
    End With
    
End Sub

Public Sub Undo()
    Call SendMessage(cboAddress.hwnd, EM_UNDO, 0, 0&)
End Sub

' -------------------------------------------------------------------
'
' -------------------------------------------------------------------
'
' -------------------------------------------------------------------

Public Property Get Caption() As String
    Caption = lblCapt.Caption
End Property
Public Property Let Caption(Value As String)
    lblCapt.Caption = Value
End Property

Public Property Get LabelVisible() As Boolean
    LabelVisible = lblCapt.Visible
    Call UserControl_Resize
End Property
Public Property Let LabelVisible(Value As Boolean)
    lblCapt.Visible = Value
    Call UserControl_Resize
End Property

Public Property Get SelText() As String
    SelText = cboAddress.SelText
End Property
Public Property Let SelText(Value As String)
    cboAddress.SelText = Value
End Property

Public Property Get SelLength() As Long
    SelLength = cboAddress.SelLength
End Property
Public Property Let SelLength(Value As Long)
    cboAddress.SelLength = Value
End Property

Public Property Get SelStart() As Long
    SelStart = cboAddress.SelStart
End Property
Public Property Let SelStart(Value As Long)
    cboAddress.SelStart = Value
End Property

Public Property Get Text() As String
    Text = cboAddress.Text
End Property
Public Property Let Text(Value As String)
    cboAddress.Text = Value
End Property

' -------------------------------------------------------------------
'
' -------------------------------------------------------------------
'
' -------------------------------------------------------------------

Private Sub cboAddress_Change()
    
    If blnDelete = True Then
        Exit Sub
    End If
    
    Static blnBusy As Boolean
    Dim nSelStart As Long
    Dim nPos As Long
    
    If blnBusy = False Then
        blnBusy = True
        With cboAddress
            nSelStart = .SelStart
            If .Text <> Empty Then
                nPos = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal .Text)
                If nPos = CB_ERR Then
                    nPos = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal "http://" & .Text)
                    If nPos <> CB_ERR Then nSelStart = .SelStart + 7
                End If
            End If
            If (nPos <> CB_ERR) Then
                .Text = .List(nPos)
                ' Show the drop down
            '    Call SendMessage(.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
                .SelStart = nSelStart
                .SelLength = Len(.Text)
                
            ' Hide the Drop Down
            'Else
            '    Dim txt As String
            '    txt = .Text
            '    Call SendMessage(.hwnd, CB_SHOWDROPDOWN, 0, ByVal 0&)
            '    .Text = txt
            '    .SelStart = Len(.Text)
            
            End If
        End With
        blnBusy = False
    End If
    
End Sub

Private Sub cboAddress_Click()
    RaiseEvent Click
End Sub

Private Sub cboAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete, vbKeyBack
            blnDelete = True
        Case Else
            blnDelete = False
    End Select
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_Resize()
    Select Case lblCapt.Visible
        Case True:
            cboAddress.Move 0, 240, ScaleWidth
            Height = 555
            
        Case False:
            cboAddress.Move 0, 0, ScaleWidth
            Height = 315
            
    End Select
End Sub
