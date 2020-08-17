VERSION 5.00
Begin VB.Form Setting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4455
   Icon            =   "Setting.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Visible         =   0   'False
   Begin VB.TextBox ShowText 
      Height          =   270
      Left            =   960
      TabIndex        =   0
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      Caption         =   "位置和大小："
      Height          =   1060
      Left            =   1920
      TabIndex        =   17
      Top             =   930
      Width           =   2415
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   480
         TabIndex        =   2
         Top             =   630
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1560
         TabIndex        =   4
         Top             =   630
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "L："
         Height          =   180
         Left            =   225
         TabIndex        =   21
         Top             =   300
         Width           =   285
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "T："
         Height          =   180
         Left            =   225
         TabIndex        =   20
         Top             =   660
         Width           =   285
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "H："
         Height          =   180
         Left            =   1320
         TabIndex        =   19
         Top             =   660
         Width           =   285
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "W："
         Height          =   180
         Left            =   1320
         TabIndex        =   18
         Top             =   300
         Width           =   285
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "窗体背景透明"
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "数码雨内容："
      Height          =   2055
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton Option5 
         Caption         =   "小写字母"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "大写字母"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "十进制数字"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "二进制数字"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "全Ascii码"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "宋体,五号"
      Top             =   540
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "退出程序"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "显示文本："
      Height          =   180
      Left            =   120
      TabIndex        =   22
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "数码雨字体："
      Height          =   180
      Left            =   2040
      TabIndex        =   10
      Top             =   580
      Width           =   1095
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   200
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "数码雨颜色："
      Height          =   180
      Left            =   2040
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Private Type ChooseColor
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Declare Function ChooseColorAPI Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long

'透明背景
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long

Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const LOGPIXELSY = 90
Private Const LF_FACESIZE = 32
Private Enum CF_
    CF_FORCEFONTEXIST = &H10000
    CF_INITTOLOGFONTSTRUCT = &H40&
    CF_BOTH = (&H1 Or &H2)
    CF_EFFECTS = &H100&
    CF_LIMITSIZE = &H2000&
End Enum
Private Type CHOOSEFONT
        lStructSize As Long
        hwndOwner As Long
        hdc As Long
        lpLogFont As Long
        iPointSize As Long
        flags As CF_
        rgbColors As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
        hInstance As Long
        lpszStyle As String
        nFontType As Integer
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long
        nSizeMax As Long
End Type
Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * LF_FACESIZE
End Type

Private Sub Check1_Click()
    If Check1.Value = 0 Then
        Call SetWindowLong(RainGround.hWnd, -20, GetWindowLong(RainGround.hWnd, -20) Or &H80000)
        Call SetLayeredWindowAttributes(RainGround.hWnd, RainGround.BackColor, 255, 2)
        
        If RainGround.WindowState = 2 Then
            RainGround.Line1.Visible = False
            RainGround.Line2.Visible = False
        End If
    Else '透明
        Call SetWindowLong(RainGround.hWnd, -20, GetWindowLong(RainGround.hWnd, -20) Or &H80000)
        Call SetLayeredWindowAttributes(RainGround.hWnd, RainGround.BackColor, 255, 1)
        
        If RainGround.WindowState = 2 Then
            RainGround.Line1.Visible = True
            RainGround.Line2.Visible = True
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = -1
    Me.Hide
    If (Check1.Value = 0 And RainGround.WindowState = 2) Then ShowCursor False
End Sub

Private Sub Label4_Click()
  Dim MyColor As ChooseColor
  MyColor.lStructSize = Len(MyColor)
  MyColor.hInstance = App.hInstance
  MyColor.hwndOwner = Me.hWnd
  MyColor.flags = 0
  MyColor.lpCustColors = String$(16 * 4, 125)
  ChooseColorAPI MyColor
  
  If MyColor.rgbResult = vbBlack Then Exit Sub
  RainGround.ForeColor = MyColor.rgbResult
  Label4.BackColor = MyColor.rgbResult
End Sub


Private Sub Text1_Click()
    Dim CF As CHOOSEFONT, LF As LOGFONT
    With LF
        .lfFaceName = StrConv(RainGround.FontName, vbFromUnicode) & vbNullChar
        .lfItalic = RainGround.FontItalic
        .lfStrikeOut = RainGround.FontStrikethru
        .lfUnderline = RainGround.FontUnderline
        .lfWeight = RainGround.Font.Weight
        .lfCharSet = RainGround.Font.Charset
        .lfHeight = -MulDiv(RainGround.FontSize, GetDeviceCaps(hdc, LOGPIXELSY), 72)
    End With
    With CF
        .rgbColors = RainGround.ForeColor
        .lStructSize = Len(CF)
        .hwndOwner = hWnd
        .hInstance = App.hInstance
        .flags = CF_BOTH Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_EFFECTS Or CF_LIMITSIZE
        .lpLogFont = VarPtr(LF)
        .nSizeMin = 8
        .nSizeMax = 72
    End With
    If CHOOSEFONT(CF) = 0 Then Exit Sub
    With RainGround
        .FontName = StrConv(LF.lfFaceName, vbUnicode)
        .FontItalic = LF.lfItalic
        .FontStrikethru = LF.lfStrikeOut
        .FontUnderline = LF.lfUnderline
        .Font.Weight = LF.lfWeight
        .Font.Charset = LF.lfCharSet
        .FontSize = -LF.lfHeight - ((-LF.lfHeight) / 4) - IIf(-LF.lfHeight Mod 4 > 1, 1, 0)
    End With
    RainGround.Form_Resize
    Text1.Text = RainGround.FontName & "," & RainGround.FontSize
End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub Text2_Change()
    If Text2.Text <> "" Then
        If RainGround.WindowState = 2 Then Exit Sub
        
        If Val(Text2.Text) <= 6 Then
            RainGround.Width = 90
        ElseIf Val(Text2.Text) >= (Screen.Width - RainGround.Left) / 15 Then
            RainGround.Width = Screen.Width - RainGround.Left
        Else
            RainGround.Width = Val(Text2.Text) * 15
        End If
    End If
End Sub

Private Sub Text2_LostFocus()
    Text2.Text = RainGround.Width / 15
End Sub

Private Sub Text3_Change()
    If Text3.Text <> "" Then
        If RainGround.WindowState = 2 Then Exit Sub
        
        If Val(Text3.Text) <= 36 Then
            RainGround.Height = 540
        ElseIf Val(Text3.Text) >= (Screen.Height - RainGround.Top) / 15 Then
            RainGround.Height = Screen.Height - RainGround.Top
        Else
            RainGround.Height = Val(Text3.Text) * 15
        End If
    End If
End Sub

Private Sub Text3_LostFocus()
    Text3.Text = RainGround.Height / 15
End Sub

Private Sub Text4_Change()
    If Text4.Text <> "" Then
        If RainGround.WindowState = 2 Then Exit Sub
        
        If Val(Text4.Text) <= 0 Then
            RainGround.Left = 0
        ElseIf Val(Text4.Text) >= (Screen.Width - RainGround.Width) / 15 Then
            RainGround.Left = Screen.Width - RainGround.Width
        Else
            RainGround.Left = Val(Text4.Text) * 15
        End If
    End If
End Sub

Private Sub Text4_LostFocus()
    Text4.Text = RainGround.Left / 15
End Sub

Private Sub Text5_Change()
    If Text5.Text <> "" Then
        If RainGround.WindowState = 2 Then Exit Sub
        
        If Val(Text5.Text) <= 0 Then
            RainGround.Top = 0
        ElseIf Val(Text5.Text) >= (Screen.Height - RainGround.Height) / 15 Then
            RainGround.Top = Screen.Height - RainGround.Top
        Else
            RainGround.Top = Val(Text5.Text) * 15
        End If
    End If
End Sub

Private Sub Text5_LostFocus()
    Text5.Text = RainGround.Top / 15
End Sub
