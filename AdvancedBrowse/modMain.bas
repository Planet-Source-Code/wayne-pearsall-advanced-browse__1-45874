Attribute VB_Name = "modMain"
Option Explicit

Sub Main()
    ' call for the webbrowser to load a new window
    Call LoadNewWindow("http://Software.TW-Chat.co.uk")
End Sub

Public Sub LoadNewWindow(Optional TargetURL As String = Empty, Optional bSilent As Boolean = False)
    ' if the url wasnt presented to the sub
    ' Use your own default URL
    If TargetURL = Empty Then TargetURL = "http://Software.TW-Chat.co.uk"
    
    ' Then load a new browser window
    Dim frmB As New frmBrowser
    Load frmB
    With frmB
        ' Set weather or not the browser displays errors
        .WebBrowser1.Silent = bSilent
        ' send the browser to the new url
        .WebBrowser1.Navigate TargetURL
        ' show the form
        .Show
    End With
    ' then destroy the form reference [No longer needed
    Set frmB = Nothing
    
End Sub
