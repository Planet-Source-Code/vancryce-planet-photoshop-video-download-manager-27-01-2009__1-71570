VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmCentral 
   BorderStyle     =   0  'None
   Caption         =   "PPVDM"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCentral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Project1.jcbutton jcbutton2 
      Height          =   615
      Left            =   8760
      TabIndex        =   15
      Top             =   10800
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1085
      ButtonStyle     =   16
      BorderStyle     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   5264457
      Caption         =   ""
      MouseIcon       =   "frmCentral.frx":57E2
      Begin Project1.jcbutton botHtms 
         Height          =   375
         Left            =   2160
         TabIndex        =   16
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         ButtonStyle     =   7
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14935011
         Caption         =   "Get *.FLV Files"
         MousePointer    =   99
         MouseIcon       =   "frmCentral.frx":5944
      End
      Begin Project1.jcbutton botDownload 
         Height          =   375
         Left            =   4320
         TabIndex        =   17
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         ButtonStyle     =   7
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14935011
         Caption         =   "Download Files"
         MousePointer    =   99
         MouseIcon       =   "frmCentral.frx":5AA6
      End
      Begin Project1.jcbutton botLimpar 
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         ButtonStyle     =   7
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14935011
         Caption         =   "Clear All"
         MousePointer    =   99
         MouseIcon       =   "frmCentral.frx":5C08
      End
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00505449&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   8760
      TabIndex        =   12
      Top             =   8760
      Width           =   6495
   End
   Begin VB.ListBox List4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFD1AD&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   13920
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFD1AD&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   14640
      TabIndex        =   9
      Top             =   3840
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.TextBox txtHtml 
      Appearance      =   0  'Flat
      Height          =   1095
      Left            =   13680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   5520
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00505449&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   8760
      TabIndex        =   0
      Top             =   7680
      Width           =   6495
   End
   Begin Project1.ucTextbox txtPasta 
      Height          =   345
      Left            =   8760
      TabIndex        =   1
      Top             =   9840
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   609
      Caption         =   "Destiny Folder"
      CaptionBold     =   0   'False
      TextFormat      =   10
      TextBackColor   =   5264457
      Text            =   "C:\Downloads\"
   End
   Begin Project1.jcbutton botFechar 
      Height          =   375
      Left            =   13920
      TabIndex        =   5
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "Exit"
      MousePointer    =   99
      MouseIcon       =   "frmCentral.frx":5D6A
   End
   Begin Project1.jcbutton botCancelar 
      Height          =   375
      Left            =   14160
      TabIndex        =   7
      Top             =   2040
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   "Cancelar"
      MousePointer    =   99
      MouseIcon       =   "frmCentral.frx":5ECC
   End
   Begin Project1.jcbutton fraTopo 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   873
      ButtonStyle     =   16
      BorderStyle     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   5264457
      Caption         =   ""
      MouseIcon       =   "frmCentral.frx":602E
      Begin Project1.jcbutton botMinimizar 
         Height          =   375
         Left            =   12480
         TabIndex        =   4
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14935011
         Caption         =   "Minimize"
         MousePointer    =   99
         MouseIcon       =   "frmCentral.frx":6190
      End
      Begin Project1.jcbutton botAbout 
         Height          =   375
         Left            =   9840
         TabIndex        =   26
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14935011
         Caption         =   "About"
         MousePointer    =   99
         MouseIcon       =   "frmCentral.frx":62F2
      End
      Begin Project1.jcbutton botHelp 
         Height          =   375
         Left            =   11160
         TabIndex        =   30
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14935011
         Caption         =   "Help"
         MousePointer    =   99
         MouseIcon       =   "frmCentral.frx":6454
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Planet Photoshop - Video Download Manager [2009]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   6495
      End
   End
   Begin Project1.ucTextbox txtFile 
      Height          =   345
      Left            =   8760
      TabIndex        =   13
      Top             =   10320
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   609
      Caption         =   "Filename"
      CaptionBold     =   0   'False
      Enabled         =   0   'False
      TextFormat      =   10
      TextBackColor   =   5264457
   End
   Begin Project1.jcbutton jcbutton3 
      Height          =   3735
      Left            =   120
      TabIndex        =   18
      Top             =   7680
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   6588
      ButtonStyle     =   16
      BorderStyle     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   5264457
      Caption         =   ""
      MouseIcon       =   "frmCentral.frx":65B6
      Begin VB.TextBox txtLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   1575
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   2040
         Width           =   8295
      End
      Begin Project1.ucTextbox txtLinks 
         Height          =   345
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   609
         Caption         =   "Total Links Added"
         CaptionBold     =   0   'False
         Enabled         =   0   'False
         BorderColor     =   5264457
         TextBackColor   =   5264457
         TextForeColor   =   16777215
         CaptionColor    =   16777215
      End
      Begin Project1.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   661
         Value           =   0
         Theme           =   10
         TextStyle       =   3
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextForeColor   =   16777215
         TextAlignment   =   2
         Text            =   "ProgressBar1"
         TextEffectColor =   16777215
         PBSCustomeColor1=   5264457
      End
      Begin Project1.ucTextbox txtVideos 
         Height          =   345
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   609
         Caption         =   "Total Videos Found"
         CaptionBold     =   0   'False
         Enabled         =   0   'False
         BorderColor     =   5264457
         TextBackColor   =   5264457
         TextForeColor   =   16777215
         CaptionColor    =   16777215
      End
      Begin Project1.ucTextbox txtPerdidos 
         Height          =   345
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   609
         Caption         =   "Lost Videos"
         CaptionBold     =   0   'False
         Enabled         =   0   'False
         BorderColor     =   5264457
         TextBackColor   =   5264457
         TextForeColor   =   16777215
         CaptionColor    =   16777215
      End
      Begin Project1.ucTextbox txtDownload 
         Height          =   345
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   609
         Caption         =   "Current Download File"
         CaptionBold     =   0   'False
         Enabled         =   0   'False
         BorderColor     =   5264457
         TextBackColor   =   5264457
         TextForeColor   =   16777215
         CaptionColor    =   16777215
      End
   End
   Begin Project1.jcbutton fraBase 
      Height          =   11415
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   20135
      ButtonStyle     =   16
      BorderStyle     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Caption         =   ""
      MouseIcon       =   "frmCentral.frx":6718
      Begin Project1.jcbutton botBrowse 
         Height          =   345
         Left            =   14760
         TabIndex        =   29
         Top             =   9840
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14935011
         Caption         =   "..."
         MousePointer    =   99
         MouseIcon       =   "frmCentral.frx":687A
      End
      Begin Project1.jcbutton jcbutton6 
         Height          =   6975
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   12303
         ButtonStyle     =   16
         BorderStyle     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   5264457
         Caption         =   ""
         MouseIcon       =   "frmCentral.frx":69DC
         Begin Project1.jcbutton botConnect 
            Height          =   975
            Left            =   4920
            TabIndex        =   28
            Top             =   2520
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   1720
            ButtonStyle     =   7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   14935011
            Caption         =   "Connect to Planet Photoshop"
            MousePointer    =   99
            MouseIcon       =   "frmCentral.frx":6B3E
         End
         Begin SHDocVwCtl.WebBrowser WebBrowser1 
            Height          =   6735
            Left            =   120
            TabIndex        =   31
            Top             =   120
            Visible         =   0   'False
            Width           =   14895
            ExtentX         =   26273
            ExtentY         =   11880
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
      End
   End
   Begin VB.Label lblHtml 
      Caption         =   "..."
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   8640
      Width           =   3375
   End
   Begin VB.Label lblTotal 
      Height          =   255
      Left            =   9120
      TabIndex        =   11
      Top             =   10920
      Visible         =   0   'False
      Width           =   3375
   End
End
Attribute VB_Name = "frmCentral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents objDoc As MSHTML.HTMLDocument
Attribute objDoc.VB_VarHelpID = -1
Dim WithEvents objWind As MSHTML.HTMLWindow2
Attribute objWind.VB_VarHelpID = -1
Dim objEvent As CEventObj

'downloading code
Private WithEvents mydl As VicsDL
Attribute mydl.VB_VarHelpID = -1
Dim IsCancelRequested As Boolean

Public varStart, varEnd, varLen As Integer
Public varFile As String
Private TargetPosition As Integer
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


'Pre-selecting a Folder with SHBrowseForFolder (API)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2005 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
' Original Source :: http://vbnet.mvps.org/index.html?code/callback/browsecallback.htm
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Code Founder :: D.W.
'Code optimized by whoknows [http://pipiscrew.6x.to] -- [http://pipiscrew.tk]
Public sFolder As String
Private Type BrowseInfo
  hOwner As Long
  pIDLRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
'-----------------------------------

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
'on top
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub OnTop()
    Dim success%
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    Const Flags = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    
    If SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags) = True Then
        success% = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
    End If
End Sub

Private Sub botAbout_Click()
    Load frmAbout
    frmAbout.Show 1
End Sub

Private Sub botBrowse_Click()
    sFolder = txtPasta.Text
    txtPasta.Text = BrowseForFolderByPIDL(sFolder)
End Sub

Private Sub botConfig_Click()
    MsgBox "Option disabled!", vbSystemModal + vbCritical, "Config"
End Sub

Private Sub botConnect_Click()
    botConnect.Visible = False
    'WBControl1.URL = "http://www.planetphotoshop.com/category/tutorials"
    'WBControl1.Visible = True
    WebBrowser1.Navigate "http://www.planetphotoshop.com/category/tutorials"
    WebBrowser1.Visible = True
    'WBControl1.object
    'Set objDoc = WBControl1.object
    'Set objWind = objDoc.parentWindow
    'Set objDoc = WebBrowser1.Document
    'Set objWind = objDoc.parentWindow
End Sub

Private Sub botDownload_Click()
    'do a single file download with form waiting for response from function
    Dim FileList As String
    Dim Pos As Integer
    Dim target As String
    
    If List3.ListCount = 0 Then
        Exit Sub
    Else
        For v = 0 To List2.ListCount - 1
            List2.ListIndex = v
            If List2.ListIndex = 0 Then
                FileList = List2.Text & "," & txtPasta.Text & GetFileFromPath(List2.Text) & "," & "1" & ","
            Else
                FileList = FileList & List2.Text & "," & txtPasta.Text & GetFileFromPath(List2.Text) & "," & "1" & ","
            End If
        Next v
        botDownload.Enabled = False
        botHtms.Enabled = False
        botLimpar.Enabled = False
        botConnect.Caption = "Downloading Files..."
        botConnect.Enabled = False
        botConnect.Visible = True
        DoEvents
        'Call frmDnLoad.ShowDownLoad(FileList, Me, Me)
'-------------------------

        ProgressBar1.Visible = True
        
        'For x = 0 To List1.ListCount - 1
'------------------------
            'List1.ListIndex = x
            
            Set mydl = New VicsDL 'implement the class on this form
            'this would usually be in the form_load event... but I do not use that event in this project
            
            'be sure the focus is set on the calling form so download can be cancelled easier
            'CallingForm.SetFocus
            'IsCancelRequested = False
            'DoEvents
            'DoFormStuff 'draw the form
            'If IsMissing(Owner) = False Then
            '    Me.Show 'You're better off to show without owner form... otherwise the function will wait till you close the form before it does anything... :(
            'Else
            '    Me.Show vbModeless, Owner 'I leave this incase you want response.
            'End If
            'Me.Show
            'Me.Refresh 'force form to be displayed 1st before processing the code that follows
            'split files to download from FileList
            Dim i, X As Integer
            Dim File2DownLoad As String, File2Save As String, DeleteCache As Boolean, TopLimit As Integer, TempDelete As String, Offset As Integer
            i = Split(FileList, ",")
            TopLimit = (UBound(i) - 2) / 3 'filelist comes in as:File2DownLoad,File2Save,DeleteCache
            Offset = 0
            For X = 0 To TopLimit '<-- start the processing loop and check occasionally for a cancel
                
                '--> Before beginning, check to see if a cancel request was received
                'If IsCancelRequested Then Exit For 'if so, leave this loop - otherwise, more files could be downloaded
                
                '--> no cancel request exists... so, start processing files <--
                File2DownLoad = i(Offset)
                File2Save = i(Offset + 1)
                TempDelete = i(Offset + 2)
                If TempDelete = "1" Then
                    DeleteCache = True
                Else
                    DeleteCache = False
                End If
                Offset = Offset + 3 'increment the offset for next file
                ProgressBar1.Value = 0 'initialize the progress bar
                
                'inform the calling form of action for purposes of this demo
                'frmDemo.Text3.Text = frmDemo.Text3.Text & "Starting Download..." & vbCrLf
                'frmDemo.Text3.Text = frmDemo.Text3.Text & File2DownLoad & vbCrLf
                
                'You may want to download the file from IE's cache... if so... set DeleteCache = False
                'however, this may result in an old file being "downloaded" from the cache and not the internet
                'Note that the remote URL is passed since this is the name that the cached file is known by.
                'This does NOT delete the file from the remote server... ONLY the local machine copy
                'Deleting the cached copy (if it exists) forces a new copy to be downloaded from internet
                
                If DeleteCache Then
                    If mydl.DeleteVicCache(File2DownLoad) = 1 Then 'file was found and deleted
                        txtLog.Text = txtLog.Text & "Found Cached File and Deleted It..." & vbCrLf
                    Else
                        txtLog.Text = txtLog.Text & "Did Not Find Cached Copy Of Requested File" & vbCrLf 'no local copy existed

                    End If
                End If
                txtDownload.Text = File2DownLoad
                ProgressBar1.Text = File2DownLoad
                'proceed with the download part
                If mydl.StartTheStinkinDownLoad(File2DownLoad, File2Save) Then
                    txtLog.Text = txtLog.Text & File2DownLoad & " - Download Completed!" & vbCrLf
                    '-->you may want some other notification back to calling form here
                Else
                    txtLog.Text = txtLog.Text & "File Download Failed!" & vbCrLf
                    'ShowDownLoad = False
                    '-->you may want some other notification back to calling form here
                End If
            Next
BailingOut:
            Set mydl = Nothing 'free memory
            ProgressBar1.Value = 0
            ProgressBar1.Visible = False
            txtLog.Text = txtLog.Text & "Download Video Files Completed!" & vbCrLf
            botLimpar.Enabled = True
            botConnect.Caption = "Operation Completed!"
'-------------------------
    End If
End Sub

Private Sub botFechar_Click()
    Unload Me
    End
    Set frmCentral = Nothing
End Sub

Private Sub botHelp_Click()
    Load frmHelp
    frmHelp.Show 1
End Sub

Private Sub botHtms_Click()
    Dim FileList As String
    If txtPasta.Text = "" Then
        MsgBox "Destiny Folder not set!", vbSystemModal + vbCritical, "Folder Error"
        Exit Sub
    End If
    If List1.ListCount = 0 Then
        Exit Sub
    Else
        'WBControl1.Visible = False
        WebBrowser1.Visible = False
        botConnect.Caption = "Downloading Links..."
        botConnect.Enabled = False
        botConnect.Visible = True
        botHtms.Enabled = False
        botLimpar.Enabled = False
        List1.Enabled = False
        List2.Enabled = False
        List3.Enabled = False
        List4.Enabled = False
        For xx = 0 To List1.ListCount - 1
            List1.ListIndex = xx
            txtFile.Text = txtPasta.Text & GetFileFromPath(List1.Text)
            'do a single file download with form waiting for response from function
            List3.AddItem txtFile.Text
            If List1.ListIndex = 0 Then
                FileList = List1.Text & "," & txtFile.Text & "," & "1" & ","
            Else
                FileList = FileList & List1.Text & "," & txtFile.Text & "," & "1" & ","
            End If
        Next xx
        DoEvents
        
        ProgressBar1.Visible = True
        
'------------------------
            
            Set mydl = New VicsDL 'implement the class on this form
            Dim i, X As Integer
            Dim File2DownLoad As String, File2Save As String, DeleteCache As Boolean, TopLimit As Integer, TempDelete As String, Offset As Integer
            i = Split(FileList, ",")
            TopLimit = (UBound(i) - 2) / 3 'filelist comes in as:File2DownLoad,File2Save,DeleteCache
            Offset = 0
            For X = 0 To TopLimit '<-- start the processing loop and check occasionally for a cancel
                File2DownLoad = i(Offset)
                File2Save = i(Offset + 1)
                TempDelete = i(Offset + 2)
                If TempDelete = "1" Then
                    DeleteCache = True
                Else
                    DeleteCache = False
                End If
                Offset = Offset + 3 'increment the offset for next file
                ProgressBar1.Value = 0 'initialize the progress bar
                
                If DeleteCache Then
                    If mydl.DeleteVicCache(File2DownLoad) = 1 Then 'file was found and deleted
                        txtLog.Text = txtLog.Text & "Found Cached File and Deleted It..." & vbCrLf
                    Else
                        txtLog.Text = txtLog.Text & "Did Not Find Cached Copy Of Requested File" & vbCrLf 'no local copy existed
                    End If
                End If
                txtDownload.Text = File2DownLoad
                ProgressBar1.Text = File2DownLoad
                If mydl.StartTheStinkinDownLoad(File2DownLoad, File2Save) Then
                    txtLog.Text = txtLog.Text & File2DownLoad & " - Download Completed!" & vbCrLf
                Else
                    txtLog.Text = txtLog.Text & "File Download Failed!" & vbCrLf
                End If
            Next
BailingOut:
            Set mydl = Nothing 'free memory
            ProgressBar1.Value = 0
            ProgressBar1.Visible = False
            txtLog.Text = txtLog.Text & "Finished Downloading Links!" & vbCrLf
            
            botConnect.Caption = "Download of the Links Finished!"
'-------------------------

    End If
    For Y = 0 To List3.ListCount - 1
        List3.ListIndex = Y
        If ItemExist(List3.Text) Then
            Call GetTextFromFile(List3.Text, txtHtml)
        End If
    
        If txtHtml.Text = "" Then
            MsgBox "Can't get HTML text.", vbSystemModal + vbCritical
            Exit Sub
        Else
            target = "swfplayer.swf?video="
            Pos = InStr(1, txtHtml.Text, target)
            If Pos > 0 Then
                varStart = Str(Pos + 20)
                target = ".flv"
                Pos = InStr(1, txtHtml.Text, target)
                If Pos > 0 Then
                    varEnd = Str(Pos)
                    varLen = varEnd - varStart
                    varFile = Right$(Left$(txtHtml.Text, varStart + varLen + 3), varLen + 4)
                    List2.AddItem "http://www.planetphotoshop.com/videos/" & varFile
                    List4.AddItem txtPasta.Text & varFile
                Else
                    GoTo salto
                End If
            Else
                GoTo salto
            End If
        End If
salto:
    Next Y
    txtVideos.Text = List2.ListCount
    txtPerdidos.Text = Val(txtLinks.Text) - Val(txtVideos.Text)
    botHtms.Enabled = False
    botLimpar.Enabled = True
    If Val(txtVideos.Text) > 0 Then botDownload.Enabled = True
End Sub

Private Sub botLimpar_Click()
    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    txtLog.Text = ""
    txtFile.Text = ""
    txtLinks.Text = ""
    txtVideos.Text = ""
    txtDownload.Text = ""
    txtPerdidos.Text = ""
    ProgressBar1.Text = ""
    WebBrowser1.Navigate "http://www.planetphotoshop.com/category/tutorials"
    WebBrowser1.Visible = False
    'WBControl1.URL = "http://www.planetphotoshop.com/category/tutorials"
    'WBControl1.Visible = False
    botConnect.Caption = "Connect to Planet Photoshop"
    botConnect.Enabled = True
    botConnect.Visible = True
    botHtms.Enabled = False
    botDownload.Enabled = False
    botLimpar.Enabled = False
End Sub

Private Sub botMinimizar_Click()
    Me.WindowState = 1
End Sub

Private Sub Form_Load()
    Clipboard.Clear
    HookForm Me
    SetIcon Me.hwnd, "AAA"
    GetProgPath
    fraBase.Width = Me.Width
    fraBase.Height = Me.Height
    botCancelar.Enabled = False
    SetClipboardViewer Me.hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnHookForm Me
End Sub

Private Sub fraTopo_DblClick()
    Me.WindowState = 0
    Me.Move 0, 0
    Me.WindowState = 2
End Sub

Private Sub fraTopo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture
    Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture
    Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 46
            List1.RemoveItem List1.ListIndex
            txtLinks.Text = Val(txtLinks.Text) - 1
            List1.Refresh
    End Select
End Sub

Sub getFile()
    Dim Pos As Integer
    Dim target As String

    If Text2.Text = "" Then
        Exit Sub
    Else
        target = "swfplayer.swf?video="
        Pos = InStr(1, Text2.Text, target)
        If Pos > 0 Then
            varStart = Str(Pos + 20)
        End If

        target = ".flv"
        Pos = InStr(1, Text2.Text, target)
        If Pos > 0 Then
            varEnd = Str(Pos)
        End If

        varLen = varEnd - varStart
        varFile = Right$(Left$(Text2.Text, varStart + varLen + 3), varLen + 4)
        List2.AddItem "http://www.planetphotoshop.com/videos/" & varFile
    End If
End Sub

Private Sub mydl_VicDLProg(ByVal VicBytesIn As Long, ByVal VicTotalBytes As Long)
    On Error GoTo OhCrap
    '++ Raised In: IBindStatusCallback_OnProgress event
    'use this event's info to update your progress bar, etc.
    
    'VicBytesIn = # of BYTES downloaded so far = ulProgress
    'VicTotalBytes = Total # of BYTES to ultimately be downloaded = ulProgressMax
    '--> URLDownloadToFile sometimes freaks out here... so control the damage...
    '    and it will catch back up with itself.
    'Here are a few combinations that I have observed while debugging:
    'ulProgress = 0: ulProgressMax = 0 -> Set ProgressBar1.Max = 0 fires error
    'ulProgress > ulProgressMax -> Set ProgressBar1.Value > ProgressBar1.Max fires error
    'I've already trapped the ulProgressMax errors in IBindStatusCallback_OnProgress
    'so all that's left to guard against is:
    
    'handle the ulProgress error possibilities
    If VicBytesIn >= 0 And VicBytesIn <= VicTotalBytes Then
        DoEvents
        ProgressBar1.Max = VicTotalBytes ' set/re-set the progress bar's max value after it is known for sure
        '-->Be sure to set the max value before assigning the bar value!
        ProgressBar1.Value = VicBytesIn ' set the current level of progress
        DoEvents 'force a refresh... even though this slows things down
    End If
Exit Sub
OhCrap:
    'this shouldn't ever fire... but there's no sense in letting your progress bar screw
    'things up now!  To my way of thinking, it's better to have a slightly mis-informed
    'user than a bad download and crash.  An error here is caused by the Progress bar.
    'I guess, if your using this to download movies (or similar) your progress bar's Max
    'and Value limits could be exceeded... so if you need to, you can use this handler to
    'hide the progress bar and switch to a text only progress update?
    Resume Next
End Sub

Private Function BrowseForFolderByPIDL(sSelPath As String) As String

   Dim BI As BrowseInfo
   Dim pidl As Long
   Dim SPath As String * 260
     
   With BI
      .hOwner = Me.hwnd
      .pIDLRoot = 0
      .lpszTitle = "Pre-selecting a folder using the folder's pidl."
      .lpfn = FARPROC(AddressOf BrowseCallbackProc)
      .lParam = SHSimpleIDListFromPath(StrConv(sSelPath, vbUnicode))
   End With
  
   pidl = SHBrowseForFolder(BI)
  
   If pidl Then
      If SHGetPathFromIDList(pidl, SPath) Then
         BrowseForFolderByPIDL = Left$(SPath, InStr(SPath, vbNullChar) - 1)
      Else
         BrowseForFolderByPIDL = ""
      End If
     
     'free the pidl from SHBrowseForFolder call
      Call CoTaskMemFree(pidl)
   Else
      BrowseForFolderByPIDL = ""
   End If
  
 'free the pidl (lparam) from GetPIDLFromPath call
   Call CoTaskMemFree(BI.lParam)
  
End Function

Private Function FARPROC(pfn As Long) As Long
  'A dummy procedure that receives and returns
  'the value of the AddressOf operator.
 
  FARPROC = pfn
End Function

Private Sub objWind_onerror(ByVal description As String, ByVal Url As String, ByVal line As Long)
    Set objEvent = objWind.event
    objEvent.returnValue = True
    'MsgBox (description)
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, Url As Variant)
    Set objDoc = WebBrowser1.Document
    Set objWind = objDoc.parentWindow
End Sub

