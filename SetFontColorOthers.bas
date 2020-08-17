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
    CF_SCREENFONTS = &H1 '显示屏幕字体
    CF_PRINTERFONTS = &H2 '显示打印机字体
    CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS) '两者都显示
    CF_EFFECTS = &H100& '添加字体效果
    CF_LIMITSIZE = &H2000& '设置字体大小限制
End Enum
Private Type CHOOSEFONT
        lStructSize As Long
        hwndOwner As Long          ' caller's window handle
        hdc As Long                ' printer DC/IC or NULL
        lpLogFont As Long 'LogFont结构地址
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
        lfHeight As Long '字体大小
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long '是否粗体
        lfItalic As Byte '是否斜体
        lfUnderline As Byte '是否下划线
        lfStrikeOut As Byte '是否删除线
        lfCharSet As Byte '字符集
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * LF_FACESIZE '字体名称
End Type
Private Sub Text1_Click()
Dim CF As CHOOSEFONT, LF As LOGFONT
With LF
    .lfFaceName = StrConv(Text1.FontName, vbFromUnicode) & vbNullChar '初始化字体名称，需要从Unicode转换，须以空字符结尾
    .lfItalic = Text1.FontItalic '初始化是否有斜体
    .lfStrikeOut = Text1.FontStrikethru '初始化是否有删除线
    .lfUnderline = Text1.FontUnderline '初始化是否有下划线
    .lfWeight = Text1.Font.Weight '初始化字体大小
    .lfCharSet = Text1.Font.Charset '初始化字符集
    .lfHeight = -MulDiv(Text1.FontSize, GetDeviceCaps(hdc, LOGPIXELSY), 72) '把字体转换为lfHeight，用到公式
End With
With CF
    .rgbColors = Text1.ForeColor '初始化字体颜色
    .lStructSize = Len(CF)
    .hwndOwner = hwnd
    .hInstance = App.hInstance
    .flags = CF_BOTH Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_EFFECTS Or CF_LIMITSIZE
    .lpLogFont = VarPtr(LF) '设置为定义好的LogFont结构地址
    .nSizeMin = 8 '最小字体大小
    .nSizeMax = 72 '最大字体大小
End With
If CHOOSEFONT(CF) = 0 Then Exit Sub '如果按“取消”则退出过程
With Text1
    .FontName = StrConv(LF.lfFaceName, vbUnicode) '设置字体名称
    .FontItalic = LF.lfItalic '设置是否斜体
    .FontStrikethru = LF.lfStrikeOut '设置是否删除线
    .FontUnderline = LF.lfUnderline '设置是否下划线
    .Font.Weight = LF.lfWeight '设置是否粗体
    .Font.Charset = LF.lfCharSet '设置字符集
    .FontSize = -LF.lfHeight - ((-LF.lfHeight) / 4) - IIf(-LF.lfHeight Mod 4 > 1, 1, 0) '设置字体大小，lfHeight与字号得转换需要用到公式
    .ForeColor = CF.rgbColors '设置字体颜色
End With
End Sub
