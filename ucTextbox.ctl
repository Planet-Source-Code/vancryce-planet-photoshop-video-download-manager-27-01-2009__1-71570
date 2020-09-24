VERSION 5.00
Begin VB.UserControl ucTextbox 
   BackColor       =   &H00FF00FF&
   ClientHeight    =   1260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   1260
   ScaleWidth      =   5550
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1200
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   60
      Width           =   2940
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   240
      Picture         =   "ucTextbox.ctx":0000
      Top             =   840
      Width           =   240
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   30
      MouseIcon       =   "ucTextbox.ctx":014A
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   75
      Width           =   60
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   0
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ucTextbox"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   120
      MouseIcon       =   "ucTextbox.ctx":029C
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   915
   End
End
Attribute VB_Name = "ucTextbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const m_def_Caption = "ucTextbox"
Const m_def_CaptionBold = True
Const m_def_AutoSelect = True
Const m_def_Enabled = True
Const m_def_Locked = False
Const m_def_BorderColor = vbBlack
Const m_def_TextBackColor = vbWhite
Const m_def_TextForeColor = vbBlack
Const m_def_CaptionColor = vbBlack
Const m_def_Text = ""
Const m_def_MaxLength = 0
Const m_def_PassText = ""
Const m_def_TextFormat = 0

Dim m_Caption As String
Dim m_CaptionBold As Boolean
Dim m_AutoSelect As Boolean
Dim m_Enabled As Boolean
Dim m_Locked As Boolean
Dim m_BorderColor As OLE_COLOR
Dim m_TextBackColor As OLE_COLOR
Dim m_TextForeColor As OLE_COLOR
Dim m_CaptionColor As OLE_COLOR
Dim m_Text As String
Dim m_MaxLength As Integer
Dim m_PassText As String
Dim m_TextFormat As TextFormats

Public Enum TextFormats
    df_AllChars = 0
    df_AlphaOnly = 1
    df_NumOnly = 2
    df_NumAndChars = 3
    df_NumAndAlpha = 4
    df_NumAndAlphaChars = 5
    df_UCase = 6
    df_UCaseAlphaOnly = 7
    df_UCaseNumAndAlpha = 8
    df_UCaseNumAndAlphaChars = 9
    df_LCase = 10
    df_LCaseAlphaOnly = 11
    df_LCaseNumAndAlpha = 12
    df_LCaseNumAndAlphaChars = 13
    df_AlphaAndChars = 14
    df_UCaseAlphaAndChars = 15
    df_LCaseAlphaAndChars = 16
End Enum

Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=Text1,Text1,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=Text1,Text1,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=Text1,Text1,-1,KeyUp

Private Sub Label1_Click()
    Text1.SetFocus
End Sub

Private Sub Label2_Click()
    Text1.SetFocus
End Sub

Private Sub Text1_Change()
   Text = Text1.Text
End Sub

Private Sub Text1_GotFocus()
    Dim TxtLen As Integer
    'put carot at end of text
    TxtLen = Len(Text1.Text)
    Text1.SelStart = TxtLen
    If AutoSelect = True Then
        On Error Resume Next
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = InputCheck(KeyAscii)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_EnterFocus()
    'Text1.SetFocus
End Sub

Private Sub UserControl_GotFocus()
    Text1.SetFocus
End Sub

Private Sub UserControl_Initialize()
   m_Caption = m_def_Caption
   m_CaptionBold = m_def_CaptionBold
   m_AutoSelect = m_def_AutoSelect
   m_Enabled = m_def_Enabled
   m_Locked = m_def_Locked
   m_TextFormat = m_def_TextFormat
   m_BorderColor = m_def_BorderColor
   m_TextBackColor = m_def_TextBackColor
   m_CaptionColor = m_def_CaptionColor
   m_Text = m_def_Text
   m_MaxLength = m_def_MaxLength
End Sub

Private Sub UserControl_InitProperties()
   Caption = Extender.Name
   CaptionBold = False
   AutoSelect = True
   Enabled = True
   Locked = False
   TextFormat = df_AllChars
   BorderColor = m_BorderColor
   TextBackColor = m_TextBackColor
   CaptionColor = m_CaptionColor
   Text = m_Text
   MaxLength = m_MaxLength
End Sub

Private Sub UserControl_Resize()
    On Error GoTo erros_usercontrol_resize
    Label2.Caption = "  " & Caption & "   "  'presizes label1 width
    'position and size all the components
    Text1.Top = 70
    Label1.Left = 20
    Image1.Top = Label1.Top
    Label1.Width = Label2.Width
    Shape1.Width = UserControl.Width
    Label1.Caption = Label2.Caption
    Text1.Left = Label1.Width + 100
    Image1.Left = Label1.Width - 150
    Text1.Width = UserControl.Width - Label1.Width - 160
    Text1.Height = Shape1.Height - 100
    UserControl.Height = Shape1.Height
    Exit Sub
erros_usercontrol_resize:
    Select Case Err.Number
        Case 380
            'MsgBox Err.Number
            Resume Next
            
        Case Else
            MsgBox Err.Number & " : " & Err.Description, vbSystemModal + vbCritical, Err.Number
            Resume Next
    
    End Select
End Sub

Public Property Get Caption() As String
   Caption = m_Caption
End Property

Public Property Let Caption(NewCaption As String)
   m_Caption = NewCaption
   Label1.Caption = m_Caption
   PropertyChanged "Caption"
   UserControl_Resize
End Property

Public Property Get CaptionBold() As Boolean
   CaptionBold = m_CaptionBold
End Property

Public Property Let CaptionBold(NewCaption As Boolean)
   m_CaptionBold = NewCaption
   Label1.FontBold = m_CaptionBold
   Label2.FontBold = m_CaptionBold
   PropertyChanged "CaptionBold"
   UserControl_Resize
End Property

Public Property Get Enabled() As Boolean
   Enabled = m_Enabled
End Property

Public Property Let Enabled(NewEnabled As Boolean)
    m_Enabled = NewEnabled
    UserControl.Enabled = m_Enabled
    PropertyChanged "Enabled"
    UserControl_Resize
End Property

Public Property Get Locked() As Boolean
   Locked = m_Locked
End Property

Public Property Let Locked(NewLocked As Boolean)
    m_Locked = NewLocked
    Text1.Locked = m_Locked
    PropertyChanged "Locked"
    UserControl_Resize
End Property

Public Property Get AutoSelect() As Boolean
   AutoSelect = m_AutoSelect
End Property

Public Property Let AutoSelect(NewSelect As Boolean)
   m_AutoSelect = NewSelect
   PropertyChanged "AutoSelect"
   'UserControl_Resize
End Property

Public Property Get TextFormat() As TextFormats
    TextFormat = m_TextFormat
End Property

Public Property Let TextFormat(ByVal New_TextFormat As TextFormats)
    m_TextFormat = New_TextFormat
    PropertyChanged "TextFormat"
End Property

Public Property Get CaptionColor() As OLE_COLOR
   CaptionColor = m_CaptionColor
End Property

Public Property Let CaptionColor(NewCaptionColor As OLE_COLOR)
   m_CaptionColor = NewCaptionColor
   Label1.ForeColor = m_CaptionColor
   PropertyChanged "CaptionColor"
   UserControl_Resize
End Property

Public Property Get BorderColor() As OLE_COLOR
   BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(NewBorderColor As OLE_COLOR)
   m_BorderColor = NewBorderColor
   Shape1.BorderColor = BorderColor
   'Label1.ForeColor = BorderColor
   PropertyChanged "BorderColor"
   UserControl_Resize
End Property

Public Property Get Text() As String
   Text = m_Text
End Property

Public Property Let Text(NewText As String)
   m_Text = NewText
   Text1.Text = m_Text
   PropertyChanged "Text"
End Property

Public Property Get PassText() As String
   PassText = m_PassText
End Property

Public Property Let PassText(NewText As String)
    On Error Resume Next
    m_PassText = NewText
    Text1.PasswordChar = Left$(m_PassText, 1)
    PropertyChanged "Text"
End Property


Public Property Get TextBackColor() As OLE_COLOR
   TextBackColor = m_TextBackColor
End Property

Public Property Let TextBackColor(NewTextBackColor As OLE_COLOR)
   m_TextBackColor = NewTextBackColor
   Text1.BackColor = m_TextBackColor
   Shape1.FillColor = m_TextBackColor
   Label1.BackColor = m_TextBackColor
   PropertyChanged "TextBackColor"
   UserControl.BackColor = m_TextBackColor
   UserControl_Resize
End Property

Public Property Get TextForeColor() As OLE_COLOR
   TextForeColor = m_TextForeColor
End Property

Public Property Let TextForeColor(NewTextForeColor As OLE_COLOR)
   m_TextForeColor = NewTextForeColor
   Text1.ForeColor = m_TextForeColor
   PropertyChanged "TextForeColor"
   UserControl_Resize
End Property


Public Property Get MaxLength() As Integer
   MaxLength = m_MaxLength
End Property

Public Property Let MaxLength(NewMaxLength As Integer)
   m_MaxLength = NewMaxLength
   Text1.MaxLength = m_MaxLength
   PropertyChanged "MaxLength"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Caption = PropBag.ReadProperty("Caption", m_def_Caption)
   CaptionBold = PropBag.ReadProperty("CaptionBold", m_def_CaptionBold)
   AutoSelect = PropBag.ReadProperty("AutoSelect", m_def_AutoSelect)
   Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
   Locked = PropBag.ReadProperty("Locked", m_def_Locked)
   TextFormat = PropBag.ReadProperty("TextFormat", m_def_TextFormat)
   BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
   TextBackColor = PropBag.ReadProperty("TextBackColor", m_def_TextBackColor)
   TextForeColor = PropBag.ReadProperty("TextForeColor", m_def_TextForeColor)
   CaptionColor = PropBag.ReadProperty("CaptionColor", m_def_CaptionColor)
   Text = PropBag.ReadProperty("Text", m_def_Text)
   PassText = PropBag.ReadProperty("PassText", m_def_PassText)
   MaxLength = PropBag.ReadProperty("MaxLength", m_def_MaxLength)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Caption", m_Caption, m_def_Caption)
        Call .WriteProperty("CaptionBold", m_CaptionBold, m_def_CaptionBold)
        Call .WriteProperty("AutoSelect", m_AutoSelect, m_def_AutoSelect)
        Call .WriteProperty("Enabled", m_Enabled, m_def_Enabled)
        Call .WriteProperty("Locked", m_Locked, m_def_Locked)
        Call .WriteProperty("TextFormat", m_TextFormat, m_def_TextFormat)
        Call .WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
        Call .WriteProperty("TextBackColor", m_TextBackColor, m_def_TextBackColor)
        Call .WriteProperty("TextForeColor", m_TextForeColor, m_def_TextForeColor)
        Call .WriteProperty("CaptionColor", m_CaptionColor, m_def_CaptionColor)
        Call .WriteProperty("Text", m_Text, m_def_Text)
        Call .WriteProperty("PassText", m_PassText, m_def_PassText)
        Call .WriteProperty("MaxLength", m_MaxLength, m_def_MaxLength)
   End With
End Sub

Private Function InputCheck(strdfText As Integer) As Integer

    Dim tmpStr As String

    tmpStr = Chr(strdfText)

    
    Select Case TextFormat
    
    Case TextFormats.df_AllChars
        InputCheck = strdfText
        Exit Function
        
    Case TextFormats.df_AlphaOnly
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = strdfText
            Exit Function
        End Select
        
    Case TextFormats.df_LCase
        InputCheck = Asc(LCase(Chr(strdfText)))
        Exit Function
        
    Case TextFormats.df_LCaseAlphaOnly
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = Asc(LCase(Chr(strdfText)))
            Exit Function
        End Select
        
    Case TextFormats.df_LCaseNumAndAlpha
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), 0 To 9, "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = Asc(LCase(Chr(strdfText)))
            Exit Function
        End Select
        
    Case TextFormats.df_LCaseNumAndAlphaChars
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), 0 To 9, ".", ",", "-", "/", "*", "+", "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = Asc(LCase(Chr(strdfText)))
            Exit Function
        End Select
        
    Case TextFormats.df_NumAndAlpha
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), 0 To 9, "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = strdfText
            Exit Function
        End Select
        
    Case TextFormats.df_NumAndAlphaChars
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), 0 To 9, ".", ",", "-", "/", "*", "+", "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = strdfText
            Exit Function
        End Select
        
    Case TextFormats.df_NumAndChars
        Select Case tmpStr
        Case Chr(8), 0 To 9, ".", ",", "-", "/", "*", "+", "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = strdfText
            Exit Function
        End Select
        
    Case TextFormats.df_NumOnly
        Select Case tmpStr
        Case Chr(8), 0 To 9
            InputCheck = strdfText
            Exit Function
        End Select
        
    Case TextFormats.df_UCase
        InputCheck = Asc(UCase(Chr(strdfText)))
        Exit Function
    
    Case TextFormats.df_UCaseAlphaOnly
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = Asc(UCase(Chr(strdfText)))
            Exit Function
        End Select
    
    Case TextFormats.df_UCaseAlphaAndChars
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), ".", ",", "-", "/", "*", "+", "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = Asc(UCase(Chr(strdfText)))
            Exit Function
        End Select
    
    Case TextFormats.df_LCaseAlphaAndChars
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), ".", ",", "-", "/", "*", "+", "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = Asc(LCase(Chr(strdfText)))
            Exit Function
        End Select
    
    Case TextFormats.df_AlphaAndChars
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), ".", ",", "-", "/", "*", "+", "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = strdfText
            Exit Function
        End Select
        
    Case TextFormats.df_UCaseNumAndAlpha
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), 0 To 9, "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = Asc(UCase(Chr(strdfText)))
            Exit Function
        End Select
        
        
    Case TextFormats.df_UCaseNumAndAlphaChars
        Select Case tmpStr
        Case "A" To "Z", "a" To "z", Chr(8), 0 To 9, ".", ",", "-", "/", "*", "+", "ü", "Ü", "ö", "Ö", "ä", "Ä", "ß"
            InputCheck = Asc(UCase(Chr(strdfText)))
            Exit Function
        End Select
    
    End Select
    
    InputCheck = 0
    
End Function

