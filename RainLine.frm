VERSION 5.00
Begin VB.Form RainGround 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "RainLine"
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8280
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FF00&
   Icon            =   "RainLine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   5  'Size
   ScaleHeight     =   5280
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   6600
      Top             =   600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   0
      X2              =   8280
      Y1              =   5265
      Y2              =   5265
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      X1              =   0
      X2              =   8280
      Y1              =   15
      Y2              =   15
   End
End
Attribute VB_Name = "RainGround"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
'����϶������ƶ�
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'��ǰ
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Dim RainLine() As String


Private Sub Form_Load()
    If UCase(Right(Command, 10)) = "FULLSCREEN" Then
        Me.Show
        Call Form_MouseUp(1, 0, 0#, 0#)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Me.WindowState = 0 Then
        ReleaseCapture
        SendMessageA Me.hWnd, &HA1, 2, 0&    '����϶������ƶ�
        If Me.WindowState = 0 Then
            Setting.Text4.Text = Me.Left / 15
            Setting.Text5.Text = Me.Top / 15
        End If
    End If
End Sub

Public Sub Form_Resize()
    '������غʹ�С�ı� ��������������
    ReDim Preserve RainLine(Int(Me.ScaleHeight / Me.TextHeight("A")))

    With Line1
        .X1 = 0
        .X2 = Me.ScaleWidth
        .Y1 = 15
        .Y2 = 15
    End With
    With Line2
        .X1 = 0
        .X2 = Me.ScaleWidth
        .Y1 = Me.ScaleHeight - 15
        .Y2 = Me.ScaleHeight - 15
    End With
    
    Setting.Text2.Text = Me.Width / 15
    Setting.Text3.Text = Me.Height / 15
    Setting.Text4.Text = Me.Left / 15
    Setting.Text5.Text = Me.Top / 15
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then   '��������
        If Me.WindowState = 2 Then
            Me.WindowState = 0
            If (Setting.Visible = False And Setting.Check1.Value = 0) Then ShowCursor True '��ʾ���
            Line1.Visible = True
            Line2.Visible = True
        ElseIf Me.WindowState = 0 Then   'ȫ��
            Me.WindowState = 2
            FullScreen
            If Setting.Check1.Value = 0 Then   '�������û��͸�� ���������ͱ�����
                If (Setting.Visible = False And Setting.Check1.Value = 0) Then ShowCursor False
                Line1.Visible = False
                Line2.Visible = False
            End If
        End If
    Else     '���ô���
        Setting.Text2.Text = Me.Width / 15
        Setting.Text3.Text = Me.Height / 15
        Setting.Text4.Text = Me.Left / 15
        Setting.Text5.Text = Me.Top / 15
        If (Setting.Check1.Value = 0 And Me.WindowState = 2) Then ShowCursor True
        Setting.Show , Me
    End If
    
    If Setting.Visible = True Then Setting.SetFocus
End Sub

Private Sub Timer1_Timer()  'ÿ10����ִ��һ��
    Dim LineN As Long, LineNew As String, CharTemp As String

    SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1 '��ǰ
    
    Me.Cls '��մ���
    
    '�������ÿ��Ԫ�ؼ̳�ǰһ��Ԫ�ص��ַ������ݣ�����һ��Ԫ���⣩
    For LineN = UBound(RainLine) To LBound(RainLine) + 1 Step -1
        RainLine(LineN) = RainLine(LineN - 1)
    Next
    
    '�������һ���µ�������
    Do While Me.TextWidth(LineNew) <= Me.ScaleWidth
        Randomize
        
        '������������������������������������������������������
        If Setting.Option1.Value = True Then
            CharTemp = Chr(126 * Rnd())          '����ȫAsci������
            'CharTemp = IIf(Rnd() > 0.7, "��", "  ")
        ElseIf Setting.Option2.Value = True Then
            CharTemp = Chr(Int(2 * Rnd() + 48))  '����������������
        ElseIf Setting.Option3.Value = True Then
            CharTemp = Chr(Int(10 * Rnd() + 48)) '����ʮ����������
        ElseIf Setting.Option4.Value = True Then
            CharTemp = Chr(Int(26 * Rnd() + 65)) '��������ĸ������
        ElseIf Setting.Option5.Value = True Then
            CharTemp = Chr(Int(26 * Rnd() + 97)) '����С��ĸ������
        End If
        '������������������������������������������������������
        
        '����µ�����ַ������Ʊ����س������з����Ͱ�����ַ����ӽ��µ�һ��������
        If (CharTemp <> Chr(9) And CharTemp <> Chr(13) And CharTemp <> Chr(10)) Then _
            LineNew = LineNew & CharTemp & " "
    Loop
    
    '�µ�һ�������긳�������һ��Ԫ��
    RainLine(0) = LineNew
    
    '�������������
    For LineN = LBound(RainLine) To UBound(RainLine)
        '�����м�һ��
        'If LineN = Int(UBound(RainLine) / 2) Then
        '    If Setting.ShowText.Text <> "" Then
        '        Me.ForeColor = vbRed
        '        Print String(Int((Me.Width / Me.TextWidth("A") - Len(Setting.ShowText.Text)) / 2), " ") & Setting.ShowText.Text
        '        Me.ForeColor = Setting.Label4.BackColor
        '    End If
        'Else
            Print RainLine(LineN)
        'End If
    Next
End Sub

Private Sub FullScreen()
    Dim LineN As Long, CharTemp As String, Index As Long
    ReDim Preserve RainLine(Int(Me.ScaleHeight / Me.TextHeight("A")))
    For LineN = LBound(RainLine) To UBound(RainLine)
        Do While Me.TextWidth(RainLine(LineN)) <= Me.ScaleWidth
            Randomize
            '������������������������������������������������������
            If Setting.Option1.Value = True Then
                CharTemp = Chr(126 * Rnd())          '����ȫAsci������
            ElseIf Setting.Option2.Value = True Then
                CharTemp = Chr(Int(2 * Rnd() + 48))  '����������������
            ElseIf Setting.Option3.Value = True Then
                CharTemp = Chr(Int(10 * Rnd() + 48)) '����ʮ����������
            ElseIf Setting.Option4.Value = True Then
                CharTemp = Chr(Int(26 * Rnd() + 65)) '��������ĸ������
            ElseIf Setting.Option5.Value = True Then
                CharTemp = Chr(Int(26 * Rnd() + 97)) '����С��ĸ������
            End If
            '������������������������������������������������������
            If (CharTemp <> Chr(9) And CharTemp <> Chr(13) And CharTemp <> Chr(10)) Then _
                RainLine(LineN) = RainLine(LineN) & CharTemp & " "
        Loop

        Print RainLine(LineN)
    Next
End Sub
