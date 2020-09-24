VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmBrowser 
   Caption         =   "- Advanced Browse"
   ClientHeight    =   6300
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin AdvancedBrowse.brwURL brwURL1 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   3
      Top             =   540
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   979
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   6045
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16828
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   2249
            MinWidth        =   1058
            Text            =   "Secure Website "
            TextSave        =   "Secure Website "
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4695
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   9735
      ExtentX         =   17171
      ExtentY         =   8281
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   953
      ButtonWidth     =   1323
      ButtonHeight    =   953
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forwards"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Home"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSavePageAs 
         Caption         =   "Save Page As..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Setup"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Preview"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditSelectNone 
         Caption         =   "Select None"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewAllowPopups 
         Caption         =   "Allow Popups"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSuppressScriptErrors 
         Caption         =   "Suppress Script Errors"
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewTextSize 
         Caption         =   "Text Size"
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "Smallest"
            Index           =   0
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "Small"
            Index           =   1
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "Medium"
            Index           =   2
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "Large"
            Index           =   3
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "Largest"
            Index           =   4
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpWSI 
         Caption         =   "Web-Software Inc."
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub brwURL1_Click()
' ignore errors
On Error Resume Next
    ' simulate a return press of the url
    Call brwURL1_KeyPress(vbKeyReturn)
End Sub

Private Sub brwURL1_KeyPress(KeyAscii As Integer)
' ignore errors
On Error Resume Next
    ' if its a return press -> send the browser to the url text
    If KeyAscii = vbKeyReturn Then WebBrowser1.Navigate brwURL1.Text
    ' then clear any errors
    If Err Then Err.Clear
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' This is to stop new windows opening in IE
    ' when the user presses Ctrl + N
    ' ------> YOU MUST SET THE FORM'S KEY-PREVIEW TO TRUE
    
    ' Shift = 2 [Ctrl] And Keycode = 78 [N]
    If ((Shift = 2) And (KeyCode = 78)) Then
        
        ' Click the new window menu
        Call mnuFileNew_Click
        
        ' then erase the key-presses
        Shift = 0
        KeyCode = 0
    End If
    
    DoEvents
    
End Sub

Private Sub Form_Load()
    ' set the forms caption to "Loading... - " & app.title
    Caption = "Loading... - " & App.Title
End Sub

Private Sub Form_Resize()
' this sub will cause alot of errors,
' no matter what
' so ignore all errors for now
On Error Resume Next

    ' now position the webbrowser [ it is the only loose control on this form ]
    WebBrowser1.Move 0, 1150, ScaleWidth, ScaleHeight - 1500
    
    ' now clear any errors
    If Err Then Err.Clear

End Sub

Private Sub mnuEdit_Click()
    mnuEditUndo.Enabled = IIf(ActiveControl.Name = "WebBrowser1", True, False)
End Sub

Private Sub mnuEditCopy_Click()
On Error Resume Next
    Select Case ActiveControl.Name
        Case "WebBrowser1":
            WebBrowser1.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT
        Case "brwURL1":
            If brwURL1.SelText <> Empty Then Clipboard.SetText brwURL1.SelText
    End Select
    If Err Then Err.Clear
End Sub

Private Sub mnuEditCut_Click()
On Error Resume Next
    Select Case ActiveControl.Name
        Case "WebBrowser1":
            WebBrowser1.ExecWB OLECMDID_CUT, OLECMDEXECOPT_DODEFAULT
        Case "brwURL1":
            With brwURL1
                If .SelText <> Empty Then Call Clipboard.SetText(.SelText)
                .blnDelete = True
                .SelText = Empty
            End With
    End Select
    If Err Then Err.Clear
End Sub

Private Sub mnuEditFind_Click()
On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_SHOWFIND, OLECMDEXECOPT_DODEFAULT
    If Err Then Err.Clear
End Sub

Private Sub mnuEditPaste_Click()
On Error Resume Next
    Select Case ActiveControl.Name
        Case "WebBrowser1":
            WebBrowser1.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DODEFAULT
        Case "brwURL1":
            brwURL1.SelText = Clipboard.GetText
    End Select
    If Err Then Err.Clear
End Sub

Private Sub mnuEditSelectAll_Click()
On Error Resume Next
    Select Case ActiveControl.Name
        Case "WebBrowser1":
            WebBrowser1.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT
        Case "brwURL1":
            With brwURL1
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
    End Select
    If Err Then Err.Clear
End Sub

Private Sub mnuEditSelectNone_Click()
On Error Resume Next
    Select Case ActiveControl.Name
        Case "WebBrowser1":
            WebBrowser1.ExecWB OLECMDID_CLEARSELECTION, OLECMDEXECOPT_DODEFAULT
        Case "brwURL1":
            With brwURL1
                .SelStart = Len(.Text)
            End With
    End Select
    If Err Then Err.Clear
End Sub

Private Sub mnuEditUndo_Click()
On Error Resume Next
    If ActiveControl.Name <> "WebBrowser1" Then Exit Sub
    WebBrowser1.ExecWB OLECMDID_UNDO, OLECMDEXECOPT_DODEFAULT
    If Err Then Err.Clear
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePrint_Click()
On Error Resume Next
    ' Print The Page
    Call WebBrowser1.ExecWB(OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT)
    ' Clear any errors
    If Err Then Err.Clear
End Sub

Private Sub mnuFilePrintPreview_Click()
On Error Resume Next
    ' print preview
    Call WebBrowser1.ExecWB(OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT, 0, 0)
    ' Clear Any Errors
    If Err Then Err.Clear
End Sub

Private Sub mnuFileProperties_Click()
On Error Resume Next
    ' Print The Page
    Call WebBrowser1.ExecWB(OLECMDID_PROPERTIES, OLECMDEXECOPT_DODEFAULT)
    ' Clear any errors
    If Err Then Err.Clear
End Sub

Private Sub mnuHelpAbout_Click()
    Load frmAbout
    frmAbout.Show , Me
End Sub

Private Sub mnuViewAllowPopups_Click()
    With mnuViewAllowPopups
        .Checked = Not .Checked
    End With
End Sub

Private Sub mnuFileNew_Click()
' dont allow errors to crash the app
On Error Resume Next

    ' Add the location url to a string
    ' [this way, if the browser hasnt got a url
    ' IE: Is still navigating to the first url
    ' IT wont set the value, etc]
    Dim sTargetURL As String
    sTargetURL = WebBrowser1.LocationURL
    
    ' Then call for a new window to be loaded
    Call LoadNewWindow(sTargetURL, WebBrowser1.Silent)
    
    ' then clear any errors
    If Err Then Err.Clear
    
    ' and finish processing the sub
    DoEvents

End Sub

Private Sub mnuFileSavePageAs_Click()
On Error Resume Next
    ' show the save page as form
    Call WebBrowser1.ExecWB(OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT, 0, 0)
End Sub

Private Sub mnuHelpWSI_Click()
    ' Since it will be your own site in the default url
    ' 9/10, I Dont need to add a URL to send em to
    Call LoadNewWindow
End Sub

Private Sub mnuOpen_Click()
    Dim frmNOpen As New frmOpen
    Load frmNOpen
    With frmNOpen
        Set .frmOwner = Me
        .Show , Me
    End With
    Set frmNOpen = Nothing
End Sub

Private Sub mnuView_Click()
    ' Update the current view zoom level check
    ' This way you dont need to move your mouse over the
    ' menu's to see the correct check mark
    Call mnuViewTextSize_Click
End Sub

Private Sub mnuViewFontSize_Click(Index As Integer)
On Error Resume Next
    ' change the zoom for the active webbrowser
    WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(Index)
    ' clear any errors
    If Err Then Err.Clear
    ' and process form events
    DoEvents
End Sub

Private Sub mnuViewSuppressScriptErrors_Click()
    With mnuViewSuppressScriptErrors
        .Checked = Not .Checked
        WebBrowser1.Silent = .Checked
    End With
End Sub

Private Sub mnuViewTextSize_Click()
On Error Resume Next

    
    ' get the current selected font size data
    Dim lgCurNum As Variant
    WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, , lgCurNum
    
    ' If an error was caused, exit the sub routine
    If Err Then Err.Clear: Exit Sub
    
    ' uncheck all the menu items
    Dim iMnuItem As Long
    For iMnuItem = 0 To 4
        mnuViewFontSize(iMnuItem).Checked = False
    Next iMnuItem
    
    ' and check the correct value
    mnuViewFontSize(CInt(lgCurNum)).Checked = True


End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next

    With WebBrowser1
        Select Case Button.Index
            Case 1: ' Back
                .GoBack
            
            Case 2: ' Forwards
                .GoForward
            
            
            Case 4: ' Stop
                .Stop
                
            Case 5: ' Refresh
                .Refresh
                
            
            Case 7: ' Home
                .GoHome
                
            Case 8: ' Search
                ' Send the person to your own url
                .Navigate "http://www.first-zone.com/?AID=7691"
                ' Or use the Browsers Default URL
                '.GoSearch
                
        End Select
    End With
    
    ' clear any possible errors
    If Err Then Err.Clear
    
    ' finish processing the sub
    DoEvents
    
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
' ignore the errors
On Error Resume Next
    ' add the url to the address combo
    brwURL1.AddItem WebBrowser1.LocationURL
    ' then clear the errors
    If Err Then Err.Clear
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
On Error Resume Next
        
    DoEvents
    
    ' If the user is not allowing popups
    If mnuViewAllowPopups.Checked = False Then
        ' Cancel the new window
        Cancel = True
        
    Else
        ' otherwise
        ' Load a new browser
        Dim frmB As New frmBrowser
        With frmB
            Set ppDisp = .WebBrowser1.object
            ' and register it as a new window
            .WebBrowser1.RegisterAsBrowser = True
            ' set the silent property of the browser
            .WebBrowser1.Silent = WebBrowser1.Silent
            ' Then show it
            .Show
        End With
        ' Then destroy the reference
        Set frmB = Nothing
        ' and say the new window is allowed
        Cancel = False
    
    End If

End Sub

Private Sub WebBrowser1_SetSecureLockIcon(ByVal SecureLockIcon As Long)
' Dont allow errors to crash the app
On Error Resume Next

    With StatusBar1.Panels(2)
        ' if the value is greater the 0,
        ' it is a secure site
        If SecureLockIcon > 0 Then
            ' so show the "Secure Site" Panel
            .Visible = True
        Else
            ' Otherwise, if it is 0, its a normal site
            ' so hide it
            .Visible = False
        End If
    End With
    
    ' then clear any errors
    If Err Then Err.Clear
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
' Dont allow errors here to crash the app
On Error Resume Next
    ' Remove excess spaces from the ends of the text
    Text = Trim$(Text)
    ' and then set the SB Panel 1 text to the text
    StatusBar1.Panels(1).Text = Text
    ' then if an error did happen, Clear it
    If Err Then Err.Clear
End Sub

Private Sub WebBrowser1_TitleChange(ByVal Text As String)
' Dont allow errors here to crash the app
On Error Resume Next
    ' Remove excess spaces from the ends of the text
    Text = Trim$(Text)
    
    ' Set the form caption to the Text & " - " & App Title
    Me.Caption = Text & " - " & App.Title
    
    ' then if an error did happen, Clear it
    If Err Then Err.Clear

End Sub

Private Sub WebBrowser1_WindowClosing(ByVal IsChildWindow As Boolean, Cancel As Boolean)
    If MsgBox("This window is attempting to close itself!" & vbCrLf & "Do you want to allow this?", vbQuestion + vbYesNo, "Allow Window Close?") = vbYes Then
        Cancel = False
        Unload Me
    Else
        Cancel = True
    End If
End Sub

