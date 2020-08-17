Attribute VB_Name = "SetFontColorOthers"
Option Explicit
Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const LOGPIXELSY = 90        ' Logical pixels/inch in Y
Private Const LF_FACESIZE = 32
Private Enum CF_
    CF_APPLY = &H200&
    CF_ANSIONLY = &H400&
    CF_TTONLY = &H40000
    CF_ENABLEHOOK = &H8&
    CF_ENABLETEMPLATE = &H10&
    CF_ENABLETEMPLATEHANDLE = &H20&
    CF_FIXEDPITCHONLY = &H4000&
    CF_NOVECTORFONTS = &H800&
    CF_NOOEMFONTS = CF_NOVECTORFONTS
    CF_NOFACESEL = &H80000
    CF_NOSCRIPTSEL = &H800000
    CF_NOSTYLESEL = &H100000
    CF_NOSIZESEL = &H200000
    CF_NOSIMULATIONS = &H1000&
    CF_NOVERTFONTS = &H1000000
    CF_SCALABLEONLY = &H20000
    CF_SCRIPTSONLY = CF_ANSIONLY
    CF_SELECTSCRIPT = &H400000
    CF_SHOWHELP = &H4&
    CF_USESTYLE = &H80&
    CF_WYSIWYG = &H8000 ' must also have CF_SCREENFONTS CF_PRINTERFONTS
    CF_FORCEFONTEXIST = &H10000
    CF_INITTOLOGFONTSTRUCT = &H40&
    CF_SCREENFONTS = &H1 '��ʾ��Ļ����
    CF_PRINTERFONTS = &H2 '��ʾ��ӡ������
    CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS) '���߶���ʾ
    CF_EFFECTS = &H100& '�������Ч��
    CF_LIMITSIZE = &H2000& '���������С����
End Enum
Private Type CHOOSEFONT
        lStructSize As Long
        hwndOwner As Long          ' caller's window handle
        hdc As Long                ' printer DC/IC or NULL
        lpLogFont As Long 'LogFont�ṹ��ַ
        iPointSize As Long         ' 10 * size in points of selected font
        flags As CF_              ' enum. type flags
        rgbColors As Long          ' returned text color
        lCustData As Long          ' data passed to hook fn.
        lpfnHook As Long           ' ptr. to hook function
        lpTemplateName As String     ' custom template name
        hInstance As Long          ' instance handle of.EXE that
                                       '    contains cust. dlg. template
        lpszStyle As String          ' return the style field here
                                       ' must be LF_FACESIZE or bigger
        nFontType As Integer          ' same value reported to the EnumFonts
                                       '    call back with the extra FONTTYPE_
                                       '    bits added
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long           ' minimum pt size allowed &
        nSizeMax As Long           ' max pt size allowed if
                                       '    CF_LIMITSIZE is used
End Type
Private Type LOGFONT
        lfHeight As Long '�����С
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long '�Ƿ����
        lfItalic As Byte '�Ƿ�б��
        lfUnderline As Byte '�Ƿ��»���
        lfStrikeOut As Byte '�Ƿ�ɾ����
        lfCharSet As Byte '�ַ���
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * LF_FACESIZE '��������
End Type
Private Sub Text1_Click()
Dim CF As CHOOSEFONT, LF As LOGFONT
With LF
    .lfFaceName = StrConv(Text1.FontName, vbFromUnicode) & vbNullChar '��ʼ���������ƣ���Ҫ��Unicodeת�������Կ��ַ���β
    .lfItalic = Text1.FontItalic '��ʼ���Ƿ���б��
    .lfStrikeOut = Text1.FontStrikethru '��ʼ���Ƿ���ɾ����
    .lfUnderline = Text1.FontUnderline '��ʼ���Ƿ����»���
    .lfWeight = Text1.Font.Weight '��ʼ�������С
    .lfCharSet = Text1.Font.Charset '��ʼ���ַ���
    .lfHeight = -MulDiv(Text1.FontSize, GetDeviceCaps(hdc, LOGPIXELSY), 72) '������ת��ΪlfHeight���õ���ʽ
End With
With CF
    .rgbColors = Text1.ForeColor '��ʼ��������ɫ
    .lStructSize = Len(CF)
    .hwndOwner = hwnd
    .hInstance = App.hInstance
    .flags = CF_BOTH Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_EFFECTS Or CF_LIMITSIZE
    .lpLogFont = VarPtr(LF) '����Ϊ����õ�LogFont�ṹ��ַ
    .nSizeMin = 8 '��С�����С
    .nSizeMax = 72 '��������С
End With
If CHOOSEFONT(CF) = 0 Then Exit Sub '�������ȡ�������˳�����
With Text1
    .FontName = StrConv(LF.lfFaceName, vbUnicode) '������������
    .FontItalic = LF.lfItalic '�����Ƿ�б��
    .FontStrikethru = LF.lfStrikeOut '�����Ƿ�ɾ����
    .FontUnderline = LF.lfUnderline '�����Ƿ��»���
    .Font.Weight = LF.lfWeight '�����Ƿ����
    .Font.Charset = LF.lfCharSet '�����ַ���
    .FontSize = -LF.lfHeight - ((-LF.lfHeight) / 4) - IIf(-LF.lfHeight Mod 4 > 1, 1, 0) '���������С��lfHeight���ֺŵ�ת����Ҫ�õ���ʽ
    .ForeColor = CF.rgbColors '����������ɫ
End With
End Sub
