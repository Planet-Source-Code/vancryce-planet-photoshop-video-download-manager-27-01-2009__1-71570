Attribute VB_Name = "basClipboard"

'In a module
'These routines are explained in our subclassing tutorial.
'http://www.allapi.net/vbtutor/subclass.php
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetClipboardViewer Lib "user32" (ByVal hwnd As Long) As Long
Public Const WM_DRAWCLIPBOARD = &H308
Public Const GWL_WNDPROC = (-4)
Dim PrevProc As Long

Public Sub HookForm(F As Form)
    PrevProc = SetWindowLong(F.hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnHookForm(F As Form)
    SetWindowLong F.hwnd, GWL_WNDPROC, PrevProc
End Sub

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
    If uMsg = WM_DRAWCLIPBOARD Then
        If Left$(Clipboard.GetText, 30) = "http://www.planetphotoshop.com" Then
            frmCentral.List1.AddItem Clipboard.GetText
            frmCentral.txtLinks.Text = Val(frmCentral.txtLinks.Text) + 1
            If Val(frmCentral.txtLinks.Text) > 0 Then
                frmCentral.botHtms.Enabled = True
                frmCentral.botLimpar.Enabled = True
            End If
        End If
    End If
End Function
