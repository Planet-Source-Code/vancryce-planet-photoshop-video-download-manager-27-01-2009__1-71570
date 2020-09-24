Attribute VB_Name = "basGeral"
Function GetFileFromPath(vPath As String) As String
    Dim Items() As String
    Dim varTemp As String
    varTemp = Right$(Left$(vPath, Len(vPath)), Len(vPath) - 1)
    Items = Split(varTemp, "/")
    If UBound(Items) = -1 Then Exit Function
    GetFileFromPath = Items(UBound(Items))
    'GetFileFromPath = Left$(varTemp, Len(varTemp) - 0)
End Function

Function GetPathFromFile(vPath As String) As String
    Dim Items() As String
    Dim varTemp1 As String
    Dim varTemp2 As String
    Dim varTemp3 As String
    varTemp1 = Left$(vPath, Len(vPath) - 1)
    varTemp2 = Right$(varTemp1, Len(varTemp1) - 1)
    'MsgBox varTemp2, vbSystemModal
    Items = Split(varTemp2, "\")
    If UBound(Items) = -1 Then Exit Function
    varTemp3 = Items(UBound(Items))
    GetPathFromFile = Left$(varTemp2, Len(varTemp2) - Len(varTemp3))
    'MsgBox varTemp3, vbSystemModal
End Function

'**************************************
' Name: Display the Text of a file in a
'     texbox
' Description:This code will take the te
'     xt from a file and display it in a textb
'     ox. This is useful for notepad type stuf
'     f. The good thing about this code is tha
'     t
'you can take text from any file that has text; Not just .txt, .doc, etc. Useful If you created your own extension.
' By: Blake Galeas
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/vb/scripts/Sho
'     wCode.asp?txtCodeId=1933&lngWId=1'for details.'**************************************

'Setup: add a 1 textbox(Text1) and 1 but
'     ton(Command1)
'Set text1's "multiline" to true and set
'     text1's scrollbars to "vertical"
'add a module(module1)
'Start Module Code here


Function GetTextFromFile(txtFile, txtopen As TextBox)
    Dim sfile As String
    Dim nfile As Integer
    nfile = FreeFile
    sfile = txtFile
    Open sfile For Input As nfile
    txtopen = Input(LOF(nfile), nfile)
    Close nfile
End Function
