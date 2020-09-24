VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmAbout.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private Declare Function SetCapture Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Const MOUSEEVENTF_LEFTUP As Long = &H4
Private Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dY As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function SetFocusAPI Lib "user32.dll" Alias "SetFocus" (ByVal hwnd As Long) As Long

' Open ULW_Readme.txt.  It is only a few paragraphs and may help understand what/why.

' Trying to offer a friend some advice on the UpdateLayeredWindow and SetLayeredWindowAttributes
' APIs, I found myself needing to understand it a bit more. Therefore, I whipped together
' a simple demo and thought it might be worth sharing.

' REQUIRES WINDOWS 2000, XP or VISTA


Private Const WS_EX_LAYERED As Long = &H80000
Private Const GWL_EXSTYLE As Long = -20
Private Const ULW_ALPHA As Long = &H2
Private Const ULW_COLORKEY As Long = &H1
Private Const WM_NCLBUTTONDOWN As Long = &HA1
Private Const HTCAPTION As Long = 2
Private Const AC_SRC_ALPHA As Long = &H1
Private Const AC_SRC_OVER As Long = &H0
Private Const GWL_STYLE As Long = -16
Private Const WS_BORDER As Long = &H800000

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type Size
    cx As Long
    cy As Long
End Type

Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, ByRef pptDst As Any, ByRef psize As Any, ByVal hdcSrc As Long, ByRef pptSrc As Any, ByVal crKey As Long, ByRef pblend As Long, ByVal dwFlags As Long) As Long
' modified above API parameters
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private mButton As Integer          ' see mouse_move & mouse_up
Private mMousePoints As POINTAPI
Private cComposite As Ac32bppDIB

'on top
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub OnTop()
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    Const Flags = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    
    If SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags) = True Then
        success% = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Unload Me
    Set frmAbout = Nothing
End Sub

Private Sub Form_Load()
    ' with this simple demo, the form must be borderless

    Dim cImage As Ac32bppDIB
    Dim lBlend As Long
    Dim srcPt As POINTAPI
    Dim srcSize As Size
    Dim lBlendFunc As Long


'------------------

    If ItemExist(ProgPath & "frmSplash.png") = False Then BuildFileFromResource ProgPath & "frmSplash.png", png000, "PNG"        'está certo

'-----------------


    ' this will be the class we hold the image in
    Set cComposite = New Ac32bppDIB
    cComposite.ManageOwnDC = True
    Set cImage = New Ac32bppDIB
    cComposite.LoadPicture_File ProgPath & "frmSplash.png"
    cImage.Render cComposite.LoadDIBinDC(True), 146, 19
    
    Set cImage = Nothing    ' not needed any longer
    
    srcSize.cx = cComposite.Width
    srcSize.cy = cComposite.Height
    
    ' apply the layered attribute
    SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    
    ' create a blend function. Change 180 below to whatever opacity you want
    'lBlendFunc = AC_SRC_OVER Or (180 * &H10000) Or (AC_SRC_ALPHA * &H1000000)
    lBlendFunc = AC_SRC_OVER Or (255 * &H10000) Or (AC_SRC_ALPHA * &H1000000)
    
    
    ' tell windows to draw our background form whenever it needs redrawing
    UpdateLayeredWindow Me.hwnd, 0&, ByVal 0&, srcSize, cComposite.LoadDIBinDC(True), srcPt, 0&, lBlendFunc, ULW_ALPHA
    
    Call OnTop

    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    'AnimateForm Me, aload, eCurtonHorizontal, 1, 33
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture
    Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Me
    Set frmAbout = Nothing
End Sub


