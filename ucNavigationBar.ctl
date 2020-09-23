VERSION 5.00
Begin VB.UserControl ucNavigationBar 
   Alignable       =   -1  'True
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   60
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   201
   Begin VB.Menu mnuPopup 
      Caption         =   "[Popup]"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupItem 
         Caption         =   "[PopupItem]"
         Index           =   0
      End
   End
End
Attribute VB_Name = "ucNavigationBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/**************************************************************************
' * Created by Fernando "Poorlyte" Girotto
' * Version: 1.0.0
' *
' * This is control is free. Use and modify as you want.
' *
' * Some code was based on CommandBar from VbAccelerator
' * SubClassing code by Paul Caton
' **************************************************************************/
Option Explicit

Public Enum BarStylesConstants
    [Office97]
    [OfficeXP]
    [Office2003]
    [Office2007]
End Enum

Private Enum BorderColorPositionConstants
    bcLeft
    bcTop
    bcRight
    bcBottom
End Enum
Private Enum WindowsXPThemesConstants
    xpCustom = -1
    xpBlue
    xpOlive
    xpSilver
End Enum
Private Enum WindowsVersionConstants
    wvWindows9X
    wvWindowsNT
    wvWindows200
    wvWindowsXP
End Enum

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Type ITEM_INFO
    Key As String
    Text As String
    ToolTip As String
    parent As Long
    Children As Boolean
    lastChildren As Boolean
End Type
Private Type ITEM_BUTTON_INFO
    ref As Long
    pos As RECT
    mouseDown As Boolean
    mouseOver As Boolean
    mouseDropDown As Boolean
End Type

Private m_hWnd As Long
Private m_hDC As Long
Private m_bDesignTime As Boolean
Private m_bResizeInterlock As Boolean

Private m_bRedraw As Boolean
Private m_bEnabled As Boolean
Private m_eStyle As BarStylesConstants
Private m_bCustomBackground As Boolean
Private m_lCustomBackgroundColor As OLE_COLOR
Private m_Picture As StdPicture
Private m_lMaskColor As OLE_COLOR
Private m_bRootSelection As Boolean

Private m_eWinVer As WindowsVersionConstants
Private m_eTheme As WindowsXPThemesConstants
Private m_bIsXP As Boolean
Private m_bHasGradientAndTransparency As Boolean
Private m_bTrueColor As Boolean
Private m_sLastTooltip As String

Private m_lBorderSize As Long ' = 2
Private m_lPaddingSize As Long ' = 2
Private m_lButtonGlyphWidth As Long ' = 24
Private Const ROOT_KEY As String = "root"

Private m_lSelected As Long
Private m_uItems() As ITEM_INFO
Private m_uButtons() As ITEM_BUTTON_INFO

Private WithEvents m_Font As StdFont
Attribute m_Font.VB_VarHelpID = -1

Public Event ButtonClick(ByVal ButtonKey As String, ByVal ButtonText As String, ByRef Cancel As Boolean)
Public Event Resize()

'/****************************************************************************
' * API DECLARATIONS
' ****************************************************************************/

'-- OS Version
Private Const VER_PLATFORM_WIN32_NT = 2
Private Type OSVERSIONINFO
    dwVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(0 To 127) As Byte
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long
'-- window functions
Private Type POINTAPI
   X As Long
   Y As Long
End Type
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function EqualRect Lib "user32" (lpRect1 As RECT, lpRect2 As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptX As Long, ByVal ptY As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
'-- drawtext declares
Private Const DT_LEFT = &H0&
Private Const DT_TOP = &H0&
Private Const DT_CENTER = &H1&
Private Const DT_RIGHT = &H2&
Private Const DT_VCENTER = &H4&
Private Const DT_BOTTOM = &H8&
Private Const DT_WORDBREAK = &H10&
Private Const DT_SINGLELINE = &H20&
Private Const DT_EXPANDTABS = &H40&
Private Const DT_TABSTOP = &H80&
Private Const DT_NOCLIP = &H100&
Private Const DT_EXTERNALLEADING = &H200&
Private Const DT_CALCRECT = &H400&
Private Const DT_NOPREFIX = &H800
Private Const DT_INTERNAL = &H1000&
Private Const DT_WORD_ELLIPSIS = &H40000
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'-- gradient functions
Private Type TRIVERTEX
   X As Long
   Y As Long
   Red As Integer
   Green As Integer
   Blue As Integer
   Alpha As Integer
End Type
Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type
Private Type GRADIENT_TRIANGLE
    Vertex1 As Long
    Vertex2 As Long
    Vertex3 As Long
End Type
Private Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_TRIANGLE, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Const GRADIENT_FILL_TRIANGLE = &H2&
Private Enum GradientFillRectType
   GRADIENT_FILL_RECT_H = 0
   GRADIENT_FILL_RECT_V = 1
End Enum
'-- drawing functions
Private Const PS_SOLID = &H0
Private Const CLR_INVALID = -1
Private Const TRANSPARENT = &H1
Private Const OPAQUE = &H2
Private Const BITSPIXEL = 12         '  Number of bits per pixel
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
'-- theme declare functions
Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, ByVal dwMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function GetThemeFilename Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal iPropId As Long, pszThemeFileName As Long, ByVal cchMaxBuffChars As Long) As Long
'-- mouse tracking declares
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum
Private Type TRACKMOUSEEVENT_STRUCT
    cbSize As Long
    dwFlags As TRACKMOUSEEVENT_FLAGS
    hwndTrack As Long
    dwHoverTime As Long
End Type
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

'/****************************************************************************
' * PAINT GRAYSCALE API DECLARATIONS - BY JIM JOSE
' ****************************************************************************/

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type
Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type
Private Type BITMAP
   bmType       As Long
   bmWidth      As Long
   bmHeight     As Long
   bmWidthBytes As Long
   bmPlanes     As Integer
   bmBitsPixel  As Integer
   bmBits       As Long
End Type
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

'/****************************************************************************
' * SUBCLASSER API DECLARATIONS - BY PAUL CATON
' ****************************************************************************/

Private Enum eMsgWhen
    MSG_AFTER = 1                                                       'Message calls back after the original (previous) WndProc
    MSG_BEFORE = 2                                                      'Message calls back before the original (previous) WndProc
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                      'Message calls back before and after the original (previous) WndProc
End Enum
Private Type tSubData                                                   'Subclass data type
    hWnd                               As Long                            'Handle of the window being subclassed
    nAddrSub                           As Long                            'The address of our new WndProc (allocated memory).
    nAddrOrig                          As Long                            'The address of the pre-existing WndProc
    nMsgCntA                           As Long                            'Msg after table entry count
    nMsgCntB                           As Long                            'Msg before table entry count
    aMsgTblA()                         As Long                            'Msg after table array
    aMsgTblB()                         As Long                            'Msg Before table array
End Type
Private sc_aSubData()                As tSubData                        'Subclass data array
Private Const ALL_MESSAGES           As Long = -1                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                      'Table A (after) entry count patch offset
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'/****************************************************************************
' * User Control and Menu
' ****************************************************************************/

Private Sub UserControl_Initialize()
    m_bRedraw = False
    m_lBorderSize = 2
    m_lPaddingSize = 4
    m_lButtonGlyphWidth = 13 '14
    Set m_Font = UserControl.Font
End Sub

Private Sub UserControl_InitProperties()
    Call ControlInitialize
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    UserControl.Refresh
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call TrackMouseDown(Button, Shift)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call TrackMouseMove(Button, Shift)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call TrackMouseUp(Button, Shift)
End Sub

Private Sub UserControl_Paint()
    Debug.Print "Paint    "
    Call Draw
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Call ControlInitialize
    Dim defFont As New StdFont
    defFont.Name = "Tahoma"
    defFont.Size = 8.25
    Set Font = PropBag.ReadProperty("Font", defFont)
    Me.Enabled = PropBag.ReadProperty("Enabled", True)
    Me.Style = PropBag.ReadProperty("Style", BarStylesConstants.Office97)
    Me.CustomBackground = PropBag.ReadProperty("CustomBackground", False)
    Me.CustomBackgroundColor = PropBag.ReadProperty("CustomBackgroundColor", SystemColorConstants.vb3DFace)
    Set Me.Picture = PropBag.ReadProperty("Picture", Nothing)
    Me.PictureMaskColor = PropBag.ReadProperty("PictureMaskColor", UserControl.MaskColor)
    Me.RootSelection = PropBag.ReadProperty("RootSelection", True)
    Me.Redraw = True
End Sub

Private Sub UserControl_Resize()
    If Not (m_bResizeInterlock) Then
        m_bResizeInterlock = True
        Call Me.Resize
        RaiseEvent Resize
        m_bResizeInterlock = False
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Font", Me.Font
    PropBag.WriteProperty "Enabled", Me.Enabled, True
    PropBag.WriteProperty "Style", Me.Style, BarStylesConstants.Office97
    PropBag.WriteProperty "CustomBackground", Me.CustomBackground, False
    PropBag.WriteProperty "CustomBackgroundColor", Me.CustomBackgroundColor, SystemColorConstants.vb3DFace
    PropBag.WriteProperty "Picture", Me.Picture, Nothing
    PropBag.WriteProperty "PictureMaskColor", Me.PictureMaskColor, UserControl.MaskColor
    PropBag.WriteProperty "RootSelection", m_bRootSelection, True
End Sub

Private Sub UserControl_Terminate()
    Call ControlTerminate
End Sub

Private Sub mnuPopupItem_Click(Index As Integer)
    Dim buttonIndex As Long
    buttonIndex = CLng(mnuPopupItem(Index).Tag)
    Call ButtonClick(buttonIndex)
End Sub

'/****************************************************************************
' * SUBCLASS HANDLER
' * MUST be the first Public routine in this file. That includes public
' * properties also.
' ****************************************************************************/

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
    Select Case uMsg
        Case WM_MOUSELEAVE: Call ButtonTrack(0, 0)
    End Select
End Sub

'/****************************************************************************
' * Public Properties and Methods for UserControl
' ****************************************************************************/

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property
Public Property Let Font(stdfnt As StdFont)
    Set m_Font = stdfnt
    Call m_Font_FontChanged("")
    Call PropertyChanged("Font")
End Property
Public Property Set Font(stdfnt As IFont)
    Set m_Font = stdfnt
    Call m_Font_FontChanged("")
    Call PropertyChanged("Font")
End Property
Private Sub m_Font_FontChanged(ByVal PropertyName As String)
    Set UserControl.Font = m_Font
    Call Me.Refresh
    Call PropertyChanged("Font")
End Sub
Public Property Get Enabled() As Boolean
    Enabled = m_bEnabled
End Property
Public Property Let Enabled(ByVal bEnabled As Boolean)
    m_bEnabled = bEnabled
    Call UserControl.Refresh
    Call PropertyChanged("Enabled")
End Property
Public Property Get Style() As BarStylesConstants
    Style = m_eStyle
End Property
Public Property Let Style(ByVal NewValue As BarStylesConstants)
    If Not (m_eStyle = NewValue) Then
        m_eStyle = NewValue
        Call ThemeInitialize(GetDesktopWindow())
        Call UserControl.Refresh
        Call PropertyChanged("Style")
    End If
End Property
Public Property Get CustomBackground() As Boolean
    CustomBackground = m_bCustomBackground
End Property
Public Property Let CustomBackground(ByVal NewValue As Boolean)
    m_bCustomBackground = NewValue
    Call UserControl.Refresh
    Call PropertyChanged("CustomBackground")
End Property
Public Property Get CustomBackgroundColor() As OLE_COLOR
    CustomBackgroundColor = m_lCustomBackgroundColor
End Property
Public Property Let CustomBackgroundColor(ByVal NewValue As OLE_COLOR)
    m_lCustomBackgroundColor = NewValue
    Call UserControl.Refresh
    Call PropertyChanged("CustomBackgroundColor")
End Property
Public Property Get Redraw() As Boolean
    Redraw = m_bRedraw
End Property
Public Property Let Redraw(ByVal NewValue As Boolean)
    m_bRedraw = NewValue
    If (m_bRedraw) Then
        Call Me.Refresh
    End If
End Property
Public Property Get Picture() As StdPicture
    Set Picture = m_Picture
End Property
Public Property Set Picture(ByVal NewValue As StdPicture)
    If Not (NewValue Is Nothing) Then
        If Not ((NewValue.Type = vbPicTypeBitmap) Or (NewValue.Type = vbPicTypeIcon)) Then
            Call Err.Raise(1, "Invalid picture type. Only BMP and ICO are supported.")
            Exit Property
        End If
    End If
    Set m_Picture = NewValue
    Call UserControl.Refresh
    Call PropertyChanged("Picture")
End Property
Public Property Get PictureMaskColor() As OLE_COLOR
    PictureMaskColor = m_lMaskColor
End Property
Public Property Let PictureMaskColor(ByVal NewValue As OLE_COLOR)
    m_lMaskColor = NewValue
    Call UserControl.Refresh
    Call PropertyChanged("PictureMaskColor")
End Property
Public Property Get RootSelection() As Boolean
    RootSelection = m_bRootSelection
End Property
Public Property Let RootSelection(ByVal NewValue As Boolean)
    m_bRootSelection = NewValue
    Call PropertyChanged("RootSelection")
End Property

Friend Sub Refresh()
    Call PrepareButtonList    ' Rearrange buttons
    Call Draw                 ' Update control UI
End Sub
Friend Sub Resize()
    Debug.Print "Resize"
    Call Me.Refresh
End Sub
'/****************************************************************************
' * Instance Methods
' ****************************************************************************/

Private Sub ControlInitialize()

    ' Properties initialization
    m_bEnabled = True
    m_bCustomBackground = False
    m_lCustomBackgroundColor = UserControl.BackColor
    Set m_Picture = Nothing
    m_lMaskColor = UserControl.MaskColor
    m_bRootSelection = True
    m_eStyle = Office97

    ' Check if is on usermode
    m_bDesignTime = Not (UserControl.Ambient.UserMode)
    If (m_bDesignTime) Then Exit Sub

    ' Initialize internal arrays
    ReDim m_uItems(0) As ITEM_INFO
    ReDim m_uButtons(0) As ITEM_BUTTON_INFO
    ' Set default item (root). This item is always visible
    m_uItems(0).Key = "root"
    m_uItems(0).Text = "Default"
    m_uItems(0).parent = -1
    m_uItems(0).lastChildren = -1
    m_uItems(0).Children = True
    ' Set default selection
    m_lSelected = 0
    ' Initialize subclassing for mouse message
    If (Ambient.UserMode) Then
        Call Subclass_Start(UserControl.hWnd)
        Call Subclass_AddMsg(UserControl.hWnd, WM_MOUSELEAVE, MSG_AFTER)
    End If
    m_eTheme = xpCustom
    m_eWinVer = wvWindows9X

    ' Store control info for drawing functions
    m_hWnd = UserControl.hWnd
    m_hDC = UserControl.hdc

End Sub

Private Sub ControlTerminate()
    On Error Resume Next
    If Not (m_hWnd = 0) Then
        Set m_Font = Nothing
        If (Ambient.UserMode) Then
            Call Subclass_Stop(m_hWnd)
        End If
    End If
    m_hWnd = 0
End Sub

'/****************************************************************************
' * Mouse Handlers
' ****************************************************************************/

Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
    Dim tme As TRACKMOUSEEVENT_STRUCT
    On Error GoTo Errs
    With tme
        .cbSize = Len(tme)
        .dwFlags = TME_LEAVE
        .hwndTrack = lng_hWnd
    End With
    Call TrackMouseEvent(tme)  ' Track the mouse leaving the indicated window via subclassing
Errs:
End Sub

Private Sub TrackMouseDown(ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants)
    If (Not m_bEnabled) Then Exit Sub

    Dim tP As POINTAPI
    Dim Index As Long
    Dim bDropDown As Boolean
    Dim bRootSel As Boolean

    Debug.Print "TrackMouseDown"

    Call GetCursorPos(tP)
    Call ScreenToClient(m_hWnd, tP)
    Index = ButtonHitTest(tP.X, tP.Y, bDropDown)
    If (Button = vbLeftButton) Then
        If (Index > 0) Then
            If (m_uItems(m_uButtons(Index).ref).Key = ROOT_KEY) And (Not m_bRootSelection) Then
                bDropDown = True
                bRootSel = True
            End If
        End If
        Call ButtonTrack(Button, Index, True, bDropDown)
        If (m_bEnabled) And (bDropDown) Then
            Call ButtonPopupMenu(Index, Not bRootSel)
        End If
    Else
        Call ButtonTrack(Button, Index, False)
    End If
End Sub

Private Sub TrackMouseMove(ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants)
    If (Not m_bEnabled) Then Exit Sub

    Dim tP As POINTAPI
    Dim Index As Long
    Call GetCursorPos(tP)
    Call ScreenToClient(m_hWnd, tP)
    Index = ButtonHitTest(tP.X, tP.Y)
    Call ButtonTrack(Button, Index)
End Sub

Private Sub TrackMouseUp(ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants)
    If (Not m_bEnabled) Then Exit Sub
    
    Dim tP As POINTAPI
    Dim Index As Long
    Dim i As Long
    
    Debug.Print "TrackMouseUp"
    
    Call GetCursorPos(tP)
    Call ScreenToClient(m_hWnd, tP)
    Index = ButtonHitTest(tP.X, tP.Y)
    If (Button = vbLeftButton) Then
        If (Index > 0) Then
            If (m_uButtons(Index).mouseOver) And (m_uButtons(Index).mouseDown) Then
                Call ButtonClick(m_uButtons(Index).ref)
                Index = ButtonHitTest(tP.X, tP.Y)
            End If
        End If
        For i = UBound(m_uButtons) To 1 Step -1
            If (m_uButtons(i).mouseDown Or m_uButtons(i).mouseDropDown) Then
                m_uButtons(i).mouseDown = False
                m_uButtons(i).mouseDropDown = False
                Call DrawOneButton(i)
            End If
        Next
    End If

End Sub

Private Function ButtonHitTest(ByVal X As Long, ByVal Y As Long, Optional ByRef dropDown As Boolean = False) As Long
    Dim btn As Long
    For btn = 1 To UBound(m_uButtons)
        With m_uButtons(btn)
            ' O ponto está sobre um botão visível?
            If (PtInRect(.pos, X, Y)) Then
                ' Está sobre a setinha se tiver filhos?
                If (X > (.pos.Right - m_lButtonGlyphWidth)) And (X < .pos.Right) Then
                    dropDown = m_uItems(.ref).Children
                End If
                ButtonHitTest = btn
                Exit Function
            End If
        End With
    Next
    dropDown = False
    ButtonHitTest = 0  ' o índice zero é o item vazio dos botões em exibição
End Function

Private Function ButtonTrack(ByVal Button As MouseButtonConstants, ByVal Index As Long, Optional ByVal mouseDown As Boolean, Optional ByVal dropDown As Boolean) As Long
    Dim i As Long
    Dim changeCount As Long
    Dim changeIndex() As Long
    Dim addChange As Boolean
    Dim track As Boolean
    Dim sToolTip As String

    For i = UBound(m_uButtons) To 1 Step -1
        addChange = False

        If (i = Index) Then
            sToolTip = m_uItems(m_uButtons(i).ref).ToolTip

            If (m_uButtons(i).mouseDropDown) Then
                If Not (dropDown) Then
                    m_uButtons(i).mouseDown = False
                    m_uButtons(i).mouseDropDown = False
                    addChange = True
                End If
            ElseIf Not (m_uButtons(i).mouseOver) Then
                If (Button = vbLeftButton) Then
                    m_uButtons(i).mouseOver = m_uButtons(i).mouseDown
                Else
                    m_uButtons(i).mouseOver = True
                End If
                addChange = True
            End If

            If (addChange) Then
                track = True
            End If

            If (mouseDown) Then
                track = False
                If Not (m_uButtons(i).mouseDown) Then
                    m_uButtons(i).mouseDown = True
                    If (dropDown) Then
                        m_uButtons(i).mouseDropDown = True
                    End If
                    addChange = True
                End If
            End If

            If (addChange) Then
                changeCount = changeCount + 1
                ReDim Preserve changeIndex(1 To changeCount) As Long
                changeIndex(changeCount) = Index
            End If
            '
        Else
            '
            If (m_uButtons(i).mouseOver) Then
                m_uButtons(i).mouseOver = False
                changeCount = changeCount + 1
                ReDim Preserve changeIndex(1 To changeCount) As Long
                changeIndex(changeCount) = i
            End If
            '
        End If  ' If (i = iIndex) Then

    Next  ' For i = UBound(m_uButtons) To 1 Step -1

    If (changeCount > 0) Then
        For i = 1 To changeCount
            If Not (m_bEnabled) Then
                m_uButtons(changeIndex(i)).mouseDown = False
                m_uButtons(changeIndex(i)).mouseDropDown = False
                m_uButtons(changeIndex(i)).mouseOver = False
            End If
            Call DrawOneButton(changeIndex(i))
        Next
    End If

    If (track) Then
        Call TrackMouseLeave(m_hWnd)
    End If

    If (StrComp(sToolTip, m_sLastTooltip) <> 0) Then
        On Error Resume Next
        UserControl.Extender.ToolTipText = sToolTip
        m_sLastTooltip = sToolTip
    End If

End Function

'/****************************************************************************
' * Drawing Functions
' ****************************************************************************/

Private Sub Draw()
    If (m_hWnd = 0) Or (Not m_bRedraw) Then Exit Sub
    'draw background
    Dim tR As RECT
    Call GetClientRect(m_hWnd, tR)
    Call DrawBackground(Me.ThemeGradientColorStart, Me.ThemeGradientColorEnd, _
        tR.Left, tR.Top, tR.Right, tR.Bottom)
    'draw buttons
    Dim i As Long
    For i = UBound(m_uButtons) To 1 Step -1
        Call DrawOneButton(i)                                        ' Paint each button visible
    Next
End Sub

Private Sub DrawOneButton(ByVal Index As Long)

    If (m_hWnd = 0) Or (Not m_bRedraw) Then Exit Sub

    Debug.Print Format$(Index, "00") & " -> DrawOneButton " & " [" & Rnd & "]"

    With m_uButtons(Index)

        ' Redesenha apenas o fundo do botão
        Call DrawBackground(Me.ThemeGradientColorStart, Me.ThemeGradientColorEnd, _
            .pos.Left, .pos.Top, .pos.Right, .pos.Bottom)

        ' Desenha a borda e o fundo específico para o estado
        If (.mouseDropDown) Then
            Call DrawBackground(Me.ThemeBackgroundCheckedColorStart, Me.ThemeBackgroundCheckedColorEnd, _
                .pos.Left, .pos.Top, .pos.Right, .pos.Bottom)
            Call DrawBorderRectangle(Me.ThemeInnerBorderCheckedColor(bcLeft), Me.ThemeInnerBorderCheckedColor(bcTop), _
                Me.ThemeInnerBorderCheckedColor(bcRight), Me.ThemeInnerBorderCheckedColor(bcBottom), _
                .pos.Left + 1, .pos.Top + 1, .pos.Right - 1, .pos.Bottom - 1)
            Call DrawBorderRectangle(Me.ThemeBorderCheckedColor(bcLeft), Me.ThemeBorderCheckedColor(bcTop), _
                Me.ThemeBorderCheckedColor(bcRight), Me.ThemeBorderCheckedColor(bcBottom), _
                .pos.Left, .pos.Top, .pos.Right, .pos.Bottom)
        ElseIf (.mouseDown) Then
            Call DrawBackground(Me.ThemeBackgroundCheckedHotColorStart, Me.ThemeBackgroundCheckedHotColorEnd, _
                .pos.Left, .pos.Top, .pos.Right, .pos.Bottom)
            Call DrawBorderRectangle(Me.ThemeInnerBorderCheckedHotColor(bcLeft), Me.ThemeInnerBorderCheckedHotColor(bcTop), _
                Me.ThemeInnerBorderCheckedHotColor(bcRight), Me.ThemeInnerBorderCheckedHotColor(bcBottom), _
                .pos.Left + 1, .pos.Top + 1, .pos.Right - 1, .pos.Bottom - 1)
            Call DrawBorderRectangle(Me.ThemeBorderCheckedHotColor(bcLeft), Me.ThemeBorderCheckedHotColor(bcTop), _
                Me.ThemeBorderCheckedHotColor(bcRight), Me.ThemeBorderCheckedHotColor(bcBottom), _
                .pos.Left, .pos.Top, .pos.Right, .pos.Bottom)
        ElseIf (.mouseOver) Then
            Call DrawBackground(Me.ThemeBackgroundHotColorStart, Me.ThemeBackgroundHotColorEnd, _
                .pos.Left, .pos.Top, .pos.Right, .pos.Bottom)
            Call DrawBorderRectangle(Me.ThemeInnerBorderHotColor(bcLeft), Me.ThemeInnerBorderHotColor(bcTop), _
                Me.ThemeInnerBorderHotColor(bcRight), Me.ThemeInnerBorderHotColor(bcBottom), _
                .pos.Left + 1, .pos.Top + 1, .pos.Right - 1, .pos.Bottom - 1)
            Call DrawBorderRectangle(Me.ThemeBorderHotColor(bcLeft), Me.ThemeBorderHotColor(bcTop), _
                Me.ThemeBorderHotColor(bcRight), Me.ThemeBorderHotColor(bcBottom), _
                .pos.Left, .pos.Top, .pos.Right, .pos.Bottom)
        End If

        ' Se tiver subitens desenha a setinha
        If (m_uItems(.ref).Children) Then
            Dim tR As RECT
            Dim tGR As RECT
            Call CopyRect(tR, .pos)        ' Dimensões para o botão dropdown
            Call CopyRect(tGR, .pos)       ' Dimensões para a setinha

            Dim lSeparatorSize As Long
            lSeparatorSize = IIf((Me.Style = Office97), 2, 1)

            If (.mouseDown) Then
                If (ThemeTextDownEffect) Then
                    Call OffsetRect(tGR, 1, 1)
                End If
                If Not (.mouseDropDown) Then
                    Call DrawBackground(Me.ThemeBackgroundHotColorStart, Me.ThemeBackgroundHotColorEnd, _
                        (tR.Right - m_lButtonGlyphWidth) + 1, tR.Top + 1, tR.Right - 1, tR.Bottom - 1)
                    tR.Left = tR.Right - m_lButtonGlyphWidth
                    tR.Top = tR.Top + lSeparatorSize
                    tR.Right = tR.Left + lSeparatorSize
                    tR.Bottom = tR.Bottom - lSeparatorSize
                    Call DrawBorderRectangle(Me.ThemeBorderHotColor(bcLeft), Me.ThemeBorderHotColor(bcTop), _
                        Me.ThemeBorderHotColor(bcRight), Me.ThemeBorderHotColor(bcBottom), _
                        tR.Left, tR.Top, tR.Right, tR.Bottom)
                End If
            ElseIf (.mouseOver) Then
                tR.Left = tR.Right - m_lButtonGlyphWidth
                tR.Top = tR.Top + lSeparatorSize
                tR.Right = tR.Left + lSeparatorSize
                tR.Bottom = tR.Bottom - lSeparatorSize
                Call DrawBorderRectangle(Me.ThemeBorderCheckedColor(bcLeft), Me.ThemeBorderCheckedColor(bcTop), _
                    Me.ThemeBorderCheckedColor(bcRight), Me.ThemeBorderCheckedColor(bcBottom), _
                    tR.Left, tR.Top, tR.Right, tR.Bottom)
            End If

            Call DrawSubItemGlyph(tGR.Right - m_lButtonGlyphWidth, tGR.Top, m_lButtonGlyphWidth, tGR.Bottom, _
                IIf((.mouseOver), Me.ThemeTextHotColor, Me.ThemeTextColor), .mouseDropDown)
        End If

        ' Desenha a imagem caso seja o botão root
        If (m_uItems(.ref).Key = ROOT_KEY) Then
            Dim lPicWidth As Long
            Dim lPicHeight As Long
            Call GetButtonPictureSize(lPicWidth, lPicHeight)
            If (lPicWidth > 0) Then
                Dim lLeft As Long
                Dim lTop As Long
                lLeft = .pos.Left + m_lPaddingSize
                lTop = .pos.Top + (((.pos.Bottom - .pos.Top) - lPicHeight) \ 2)
                If (Me.ThemeTextDownEffect) And (.mouseDown) Then
                    lLeft = lLeft + 1
                    lTop = lTop + 1
                End If
                Call PaintPictureEx(m_Picture, lLeft, lTop, , , Not m_bEnabled)
                lPicWidth = lPicWidth + CLng(IIf((m_uItems(.ref).Text = ""), 0, m_lPaddingSize))
            End If
        End If

        ' Desenha o texto
        Dim tTR As RECT
        Call CopyRect(tTR, .pos)
        Call InflateRect(tTR, -m_lPaddingSize, 0)
        Call OffsetRect(tTR, lPicWidth, 0)
        If (Me.ThemeTextDownEffect) And (.mouseDown) Then
            Call OffsetRect(tTR, 1, 1)
        End If
        Call SetBkMode(m_hDC, TRANSPARENT)
        Call SetTextColor(m_hDC, IIf((.mouseOver), Me.ThemeTextHotColor, Me.ThemeTextColor))
        Call DrawText(m_hDC, m_uItems(.ref).Text, Len(m_uItems(.ref).Text), tTR, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER)

    End With

End Sub

Private Sub PaintPictureEx(ByVal hPicture As Long, _
                           ByVal lLeft As Long, _
                           ByVal lTop As Long, _
                           Optional ByVal lWidth As Long = -1, _
                           Optional ByVal lHeight As Long = -1, _
                           Optional ByVal bGrayscaled As Boolean = False)

    Dim BMP        As BITMAP
    Dim BMPiH      As BITMAPINFOHEADER
    Dim lBits()    As Byte 'Packed DIB
    Dim lTrans()   As Byte 'Packed DIB
    Dim TmpDC      As Long
    Dim X          As Long
    Dim xMax       As Long
    Dim TmpCol     As Long
    Dim R1         As Long
    Dim G1         As Long
    Dim B1         As Long
    Dim lMaskColor As Long
    Dim RM1        As Long ' red mask color
    Dim GM1        As Long ' green mask color
    Dim BM1        As Long ' blue mask color
    Dim bIsIcon    As Boolean

    'Get the Image format
    If (GetObjectType(hPicture) = 0) Then
        Dim mIcon As ICONINFO
        bIsIcon = True
        Call GetIconInfo(hPicture, mIcon)
        hPicture = mIcon.hbmColor
    End If

    'Get image info
    Call GetObject(hPicture, Len(BMP), BMP)

    'Prepare DIB header and redim. lBits() array
    With BMPiH
       .biSize = Len(BMPiH) '40
       .biPlanes = 1
       .biBitCount = 24
       .biWidth = BMP.bmWidth
       .biHeight = BMP.bmHeight
       .biSizeImage = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight
        If lWidth = -1 Then lWidth = .biWidth
        If lHeight = -1 Then lHeight = .biHeight
    End With
    ReDim lBits(Len(BMPiH) + BMPiH.biSizeImage)   '[Header + Bits]

    'Create TemDC and Get the image bits
    TmpDC = CreateCompatibleDC(m_hDC)
    Call GetDIBits(TmpDC, hPicture, 0, BMP.bmHeight, lBits(0), BMPiH, 0)

    ' Prepare mask dib array
    ReDim lTrans(Len(BMPiH) + BMPiH.biSizeImage)

    ' Get mask color used on bitmaps
    lMaskColor = ColorTranslate(m_lMaskColor)
    RM1 = (lMaskColor And &HFF)
    GM1 = (lMaskColor And &HFF00&) \ &H100&
    BM1 = (lMaskColor And &HFF0000) \ &H10000

    'Loop through the array... (grayscale - average!!)
    xMax = BMPiH.biSizeImage ' - 1
    For X = 0 To xMax - 3 Step 3
        R1 = lBits(X)
        G1 = lBits(X + 1)
        B1 = lBits(X + 2)

        ' Prepare mask colors to bitmaps
        If ((R1 = RM1) And (G1 = GM1) And (B1 = BM1)) And Not (bIsIcon) Then
            lTrans(X) = 255
            lTrans(X + 1) = 255
            lTrans(X + 2) = 255
            lBits(X) = 0
            lBits(X + 1) = 0
            lBits(X + 2) = 0

        ' Conver colors to grayscale
        ElseIf (bGrayscaled) Then
            TmpCol = (R1 + G1 + B1) \ 3
            If (TmpCol > 0) Then
                ' Turn color more lighten
                TmpCol = IIf((TmpCol + 30) > 255, 255, TmpCol + 30)
            End If
            lBits(X) = TmpCol
            lBits(X + 1) = TmpCol
            lBits(X + 2) = TmpCol
        End If

    Next

    ' Paint it!
    If bIsIcon Then
        Call GetDIBits(TmpDC, mIcon.hbmMask, 0, BMP.bmHeight, lTrans(0), BMPiH, 0)        ' Get the icon mask
        Call StretchDIBits(m_hDC, lLeft, lTop, lWidth, lHeight, 0, 0, BMP.bmWidth, BMP.bmHeight, lTrans(0), BMPiH, 0, vbSrcAnd)    ' Draw the mask
        Call StretchDIBits(m_hDC, lLeft, lTop, lWidth, lHeight, 0, 0, BMP.bmWidth, BMP.bmHeight, lBits(0), BMPiH, 0, vbSrcPaint)   'Draw the gray
        Call DeleteObject(mIcon.hbmMask)   'Delete the extracted images
        Call DeleteObject(mIcon.hbmColor)
    Else
        Call StretchDIBits(m_hDC, lLeft, lTop, lWidth, lHeight, 0, 0, BMP.bmWidth, BMP.bmHeight, lTrans(0), BMPiH, 0, vbSrcAnd)    ' Draw the mask
        Call StretchDIBits(m_hDC, lLeft, lTop, lWidth, lHeight, 0, 0, BMP.bmWidth, BMP.bmHeight, lBits(0), BMPiH, 0, vbSrcPaint)  ' vbSrcCopy
    End If

    'Clear memory
    Call DeleteDC(TmpDC)

End Sub

Private Sub DrawSubItemGlyph(ByVal lLeft As Long, ByVal lTop As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal color As Long, Optional ByVal bVerticalOrientation As Boolean = False)
    Dim lCentreY As Long
    Dim lCentreX As Long
    Dim tJ As POINTAPI
    Dim hPen As Long
    Dim hPenOld As Long

    lCentreX = (lLeft + lWidth \ 2)
    lCentreY = (lTop + lHeight \ 2)

    hPen = CreatePen(PS_SOLID, 1, color)
    hPenOld = SelectObject(m_hDC, hPen)

    If (bVerticalOrientation) Then
        lCentreY = lCentreY - 2
        Call MoveToEx(m_hDC, lCentreX - 3, lCentreY, tJ)
        Call LineTo(m_hDC, lCentreX + 2, lCentreY)
        Call MoveToEx(m_hDC, lCentreX - 2, lCentreY + 1, tJ)
        Call LineTo(m_hDC, lCentreX + 1, lCentreY + 1)
        Call SetPixel(m_hDC, lCentreX - 1, lCentreY + 2, color)
    Else
        lCentreY = lCentreY - 1
        Call MoveToEx(m_hDC, lCentreX - 1, lCentreY - 2, tJ)
        Call LineTo(m_hDC, lCentreX - 1, lCentreY + 3)
        Call MoveToEx(m_hDC, lCentreX, lCentreY - 1, tJ)
        Call LineTo(m_hDC, lCentreX, lCentreY + 2)
        Call SetPixel(m_hDC, lCentreX + 1, lCentreY, color)
    End If

    Call SelectObject(m_hDC, hPenOld)
    Call DeleteObject(hPen)
End Sub

Private Sub DrawBackground( _
        ByVal colorStart As Long, _
        ByVal colorEnd As Long, _
        ByVal Left As Long, _
        ByVal Top As Long, _
        ByVal Right As Long, _
        ByVal Bottom As Long, _
        Optional ByVal horizontal As Boolean = False _
    )
    If (colorStart = TRANSPARENT) Or (colorEnd = TRANSPARENT) Then
        ' do nothing
    Else
        Dim tR As RECT
        Call SetRect(tR, Left, Top, Right, Bottom)
        If (colorStart = colorEnd) Then
            ' solid fill:
            Dim hBr As Long
            hBr = CreateSolidBrush(colorStart)
            Call FillRect(m_hDC, tR, hBr)
            Call DeleteObject(hBr)
        Else
            ' gradient fill vertical:
            Call GradientFillRect(hdc, tR, _
                colorStart, colorEnd, _
                IIf(horizontal, GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V))
        End If
    End If
End Sub

Private Sub DrawBorderRectangle( _
      ByVal colorLeft As Long, _
      ByVal colorTop As Long, _
      ByVal colorRight As Long, _
      ByVal colorBottom As Long, _
      ByVal Left As Long, _
      ByVal Top As Long, _
      ByVal Right As Long, _
      ByVal Bottom As Long _
   )
Dim tJ As POINTAPI
Dim hPenOld As Long
Dim hPen As Long

    If Not (colorLeft = TRANSPARENT) Then
        hPen = CreatePen(PS_SOLID, 1, colorLeft)
        hPenOld = SelectObject(m_hDC, hPen)
        Call MoveToEx(m_hDC, Left, Bottom - 2, tJ)
        Call LineTo(m_hDC, Left, Top)  'left line
        Call SelectObject(m_hDC, hPenOld)
        Call DeleteObject(hPen)
    End If
    If Not (colorTop = TRANSPARENT) Then
        hPen = CreatePen(PS_SOLID, 1, colorTop)
        hPenOld = SelectObject(m_hDC, hPen)
        Call MoveToEx(m_hDC, Left, Top, tJ)
        Call LineTo(m_hDC, Right - 1, Top)  'top line
        Call SelectObject(m_hDC, hPenOld)
        Call DeleteObject(hPen)
    End If
    If Not (colorRight = TRANSPARENT) Then
        hPen = CreatePen(PS_SOLID, 1, colorRight)
        hPenOld = SelectObject(m_hDC, hPen)
        Call MoveToEx(m_hDC, Right - 1, Top, tJ)
        Call LineTo(m_hDC, Right - 1, Bottom - 1)  'right line
        Call SelectObject(m_hDC, hPenOld)
        Call DeleteObject(hPen)
    End If
    If Not (colorBottom = TRANSPARENT) Then
        hPen = CreatePen(PS_SOLID, 1, colorBottom)
        hPenOld = SelectObject(m_hDC, hPen)
        Call MoveToEx(m_hDC, Right - 1, Bottom - 1, tJ)
        Call LineTo(m_hDC, Left - 1, Bottom - 1)  'bottom line
        Call SelectObject(m_hDC, hPenOld)
        Call DeleteObject(hPen)
    End If
End Sub

Private Sub GradientFillRect( _
      ByVal lHDC As Long, _
      tR As RECT, _
      ByVal oStartColor As OLE_COLOR, _
      ByVal oEndColor As OLE_COLOR, _
      ByVal eDir As GradientFillRectType _
   )
Dim hBrush As Long
Dim lStartColor As Long
Dim lEndColor As Long
Dim lR As Long

   ' Use GradientFill:
   If (m_bHasGradientAndTransparency) Then
      lStartColor = ColorTranslate(oStartColor)
      lEndColor = ColorTranslate(oEndColor)

      Dim tTV(0 To 1) As TRIVERTEX
      Dim tGR As GRADIENT_RECT

      Call ColorSetTriVertex(tTV(0), lStartColor)
      tTV(0).X = tR.Left
      tTV(0).Y = tR.Top
      Call ColorSetTriVertex(tTV(1), lEndColor)
      tTV(1).X = tR.Right
      tTV(1).Y = tR.Bottom

      tGR.UpperLeft = 0
      tGR.LowerRight = 1

      Call GradientFill(lHDC, tTV(0), 2, tGR, 1, eDir)

   Else
      ' Fill with solid brush:
      hBrush = CreateSolidBrush(ColorTranslate(oEndColor))
      Call FillRect(lHDC, tR, hBrush)
      Call DeleteObject(hBrush)
   End If

End Sub

Private Sub ColorSetTriVertex(tTV As TRIVERTEX, lColor As Long)
Dim lRed As Long
Dim lGreen As Long
Dim lBlue As Long
   lRed = (lColor And &HFF&) * &H100&
   lGreen = (lColor And &HFF00&)
   lBlue = (lColor And &HFF0000) \ &H100&
   Call ColorSetTriVertexComponent(tTV.Red, lRed)
   Call ColorSetTriVertexComponent(tTV.Green, lGreen)
   Call ColorSetTriVertexComponent(tTV.Blue, lBlue)
End Sub
Private Sub ColorSetTriVertexComponent(ByRef iColor As Integer, ByVal lComponent As Long)
   If (lComponent And &H8000&) = &H8000& Then
      iColor = (lComponent And &H7F00&)
      iColor = iColor Or &H8000
   Else
      iColor = lComponent
   End If
End Sub
Private Function ColorTranslate(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, ColorTranslate) Then
        ColorTranslate = CLR_INVALID
    End If
End Function
Private Function ColorBlend(ByVal oColorFrom As OLE_COLOR, ByVal oColorTo As OLE_COLOR, Optional ByVal Alpha As Long = 128) As Long
    Dim lCFrom As Long
    Dim lCTo As Long
    lCFrom = ColorTranslate(oColorFrom)
    lCTo = ColorTranslate(oColorTo)

    Dim lSrcR As Long
    Dim lSrcG As Long
    Dim lSrcB As Long
    Dim lDstR As Long
    Dim lDstG As Long
    Dim lDstB As Long
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000

    ColorBlend = RGB( _
        ((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), _
        ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), _
        ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255) _
    )

End Function

'/****************************************************************************
' * Button Events
' ****************************************************************************/

Private Sub ButtonClick(ByVal Index As Long)
    Dim btn As ITEM_INFO
    Dim bCancel As Boolean
    
    btn.Key = m_uItems(Index).Key
    btn.Text = m_uItems(Index).Text
    
    RaiseEvent ButtonClick(btn.Key, btn.Text, bCancel)
    
    If Not (bCancel) Then
        If (m_lSelected <> Index) Then
            m_lSelected = Index
            Me.Refresh
        Else
            Call DrawOneButton(ItemVisible(Index))
        End If
    End If
End Sub

Private Sub ButtonPopupMenu(ByVal Index As Long, Optional ByVal bButtonGlyphyAlign As Boolean = True)
    Dim menuCount As Long
    Dim i As Long

    ' Remove os menus que já tenham sido configurados
    For i = mnuPopupItem.UBound To 1 Step -1
        Unload mnuPopupItem(i)
    Next

    For i = 0 To UBound(m_uItems)
        ' procura todos os menus que tenham o botão indicado como pai
        If (m_uItems(i).parent = m_uButtons(Index).ref) Then
            If (menuCount > 0) Then
                Load mnuPopupItem(menuCount)
            End If
            mnuPopupItem(menuCount).Caption = m_uItems(i).Text
            mnuPopupItem(menuCount).Visible = True
            mnuPopupItem(menuCount).Tag = i
            mnuPopupItem(menuCount).Checked = IIf((m_lSelected = i), vbChecked, vbUnchecked)
            menuCount = menuCount + 1
        End If
    Next

    If (menuCount > 0) Then
        If (bButtonGlyphyAlign) Then
            Call PopupMenu(mnuPopup, vbPopupMenuLeftButton, _
                           m_uButtons(Index).pos.Right - m_lButtonGlyphWidth, _
                           m_uButtons(Index).pos.Bottom)
        Else
            Call PopupMenu(mnuPopup, vbPopupMenuLeftButton, _
                           m_uButtons(Index).pos.Left, _
                           m_uButtons(Index).pos.Bottom)
        End If
        Call ButtonTrack(0, Index, False, False)
    End If

End Sub

'/****************************************************************************
' * Button Functions
' ****************************************************************************/

Private Sub PrepareButtonList()
    Dim btn As Long
    Dim lSize As Long
    Dim lDisplaySize As Long
    Dim lDisplayButtonsSize As Long

    ReDim m_uButtons(0) 'As DISPLAY_BUTTON_INFO

    If (m_hWnd = 0) Or (Not m_bRedraw) Then Exit Sub

    Dim cR As RECT
    Call GetClientRect(m_hWnd, cR)
    lDisplaySize = cR.Right - cR.Left

    ' A dimensão mínima é o tamanho do root + último nível
    lDisplayButtonsSize = GetButtonSize(0)
    ' Se o último nível tb não for o root então verifica quais botões devem ser exibidos
    If (m_lSelected > 0) Then
        ' Mostra sempre o último nível(botão)
        Call AddButtonToList(m_lSelected)
        ' Acrescenta o tamanho do botão selecionado
        lDisplayButtonsSize = lDisplayButtonsSize + GetButtonSize(m_lSelected)
        ' Acrescenta, se houver espaço, os níveis acessados
        btn = m_uItems(m_lSelected).parent
        Do While (btn > 0)
            lSize = GetButtonSize(btn)
            ' Verifica a área disponível consegue exibir o botão
            If ((lSize + lDisplayButtonsSize) > lDisplaySize) Then
                Exit Do
            End If
            Call AddButtonToList(btn)
            lDisplayButtonsSize = lDisplayButtonsSize + lSize
            btn = m_uItems(btn).parent
        Loop
    End If
    ' Mostra sempre o primeiro nível(root)
    Call AddButtonToList(0)
    
    ' Calculate positions
    Dim offSet As Long
    For btn = UBound(m_uButtons) To 1 Step -1
        Call SetButtonSize(btn)
        Call OffsetRect(m_uButtons(btn).pos, offSet, 0)
        offSet = m_uButtons(btn).pos.Right
    Next

End Sub

Private Sub AddButtonToList(ByVal buttonRef As Long)
    Dim lNewIndex As Long
    lNewIndex = UBound(m_uButtons) + 1
    ReDim Preserve m_uButtons(lNewIndex) 'As DISPLAY_BUTTON_INFO
    m_uButtons(lNewIndex).ref = buttonRef
    m_uButtons(lNewIndex).mouseDown = False
    m_uButtons(lNewIndex).mouseOver = False
End Sub

Private Sub GetButtonPictureSize(ByRef lWidth As Long, ByRef lHeight As Long)
    ' Converts from HiMetric to pixel
    Dim lpWidth As Long
    Dim lpHeight As Long
    If Not (m_Picture Is Nothing) Then
        If (m_Picture.Handle <> 0) Then
            lpWidth = CLng(((m_Picture.Width / 2540) * 1440) / Screen.TwipsPerPixelX)
            lpHeight = CLng(((m_Picture.Height / 2540) * 1440) / Screen.TwipsPerPixelY)
        End If
    End If
    lWidth = lpWidth
    lHeight = lpHeight
End Sub

Private Function GetButtonSize(ByVal Index As Long) As Long
'param: button index
    Dim tR As RECT

    Call SetRect(tR, 0, 0, 64, 32)
    Call DrawText(m_hDC, m_uItems(Index).Text, -1, tR, DT_CALCRECT Or DT_LEFT)
    tR.Right = (tR.Right - tR.Left) + (m_lBorderSize * 2) + (m_lPaddingSize * 2)

    If (m_uItems(Index).Children) Then
        tR.Right = tR.Right + m_lPaddingSize + m_lButtonGlyphWidth
    End If

    If (m_uItems(Index).Key = ROOT_KEY) Then
        Dim lWdt As Long
        Call GetButtonPictureSize(lWdt, 0)
        tR.Right = tR.Right + lWdt + CLng(IIf(m_uItems(Index).Text = "", 0, m_lPaddingSize))
    End If

    GetButtonSize = tR.Right

End Function

Private Sub SetButtonSize(ByVal DisplayIndex As Long)
'param: display button index
    Dim tR As RECT
    Dim tCR As RECT

    Call GetClientRect(m_hWnd, tCR)

    Call SetRect(tR, m_lBorderSize, m_lBorderSize, 64, 32)
    Call DrawText(m_hDC, m_uItems(m_uButtons(DisplayIndex).ref).Text, -1, tR, DT_CALCRECT Or DT_LEFT)
    tR.Right = (tR.Right - tR.Left) + (m_lBorderSize * 2) + (m_lPaddingSize * 2)
    tR.Bottom = tCR.Bottom - m_lBorderSize

    If (m_uItems(m_uButtons(DisplayIndex).ref).Children) Then
        tR.Right = tR.Right + m_lPaddingSize + m_lButtonGlyphWidth
    End If

    If (m_uItems(m_uButtons(DisplayIndex).ref).Key = ROOT_KEY) Then
        Dim lWdt As Long
        Call GetButtonPictureSize(lWdt, 0)
        tR.Right = tR.Right + lWdt + CLng(IIf(m_uItems(m_uButtons(DisplayIndex).ref).Text = "", 0, m_lPaddingSize))
    End If

    Call CopyRect(m_uButtons(DisplayIndex).pos, tR)

End Sub

'/****************************************************************************
' * Items Functions
' ****************************************************************************/

Public Function ItemAdd(ByVal sKey As String, ByVal sParentKey As String, ByVal sText As String, Optional ByVal sToolTip As String) As Boolean
    Dim newIndex As Long
    ' Prevent duplicated key
    If Not (ItemIndex(sKey) = -1) Then
        Exit Function
    End If
    ' Create new item
    newIndex = UBound(m_uItems) + 1
    ReDim Preserve m_uItems(newIndex)
    ' Setup new item
    With m_uItems(newIndex)
        .Key = sKey
        .Text = Trim$(sText)
        .ToolTip = Trim$(sToolTip)
        .parent = ItemIndex(sParentKey)
        .parent = IIf((.parent = -1), 0, .parent)
        If Not (m_uItems(.parent).Children) Then
            m_uItems(.parent).Children = True
        End If
    End With
    ItemAdd = True
End Function

Public Function ItemRemove(ByVal sKey As String) As Boolean
    If (LCase$(sKey) = ROOT_KEY) Then Exit Function
    Dim Index As Long
    Dim bNeedUpdate As Boolean
    Index = ItemIndex(sKey)
    If (Index > 0) Then
        Dim tmp() As ITEM_INFO
        Dim i As Long
        Dim n As Long
        For i = LBound(m_uItems) To UBound(m_uItems)
            If (Not i = Index) Then
                If (Not ItemRelated(i, Index)) Then
                    ReDim Preserve tmp(n)
                    tmp(n) = m_uItems(i)
                    If (m_uItems(i).Children) Then
                        Dim j As Long
                        For j = LBound(m_uItems) To UBound(m_uItems)
                            If (m_uItems(j).parent = i) Then
                                m_uItems(j).parent = n
                            End If
                        Next
                    End If
                    n = n + 1
                End If
            End If
        Next
        If (m_lSelected = Index) Then
            m_lSelected = m_uItems(Index).parent
            bNeedUpdate = True
        ElseIf (ItemRelated(m_lSelected, Index)) Then
            m_lSelected = m_uItems(Index).parent
            bNeedUpdate = True
        End If
        m_uItems = tmp
        If (m_uItems(m_lSelected).Key <> ROOT_KEY) Then
            m_uItems(m_lSelected).Children = (ItemChildCount(m_lSelected) > 0)
        End If
        ItemRemove = True
        If (bNeedUpdate) Then
            Call Me.Refresh
        End If
    End If
End Function

Public Sub ItemClear()
    ReDim Preserve m_uItems(0)
    m_lSelected = 0
    m_uItems(0).Text = "Default"
    m_uItems(0).ToolTip = ""
    m_uItems(0).Children = True
    Call Me.Refresh
End Sub

Public Function ItemCount() As Long
    ItemCount = UBound(m_uItems)
End Function

Public Function ItemSetText(ByVal sKey As String, ByVal sText As String, Optional ByVal sToolTip) As Boolean
    Dim Index As Long
    Index = ItemIndex(sKey)
    If (Index > -1) Then
        m_uItems(Index).Text = Trim$(sText)
        If Not IsMissing(sToolTip) Then
            m_uItems(Index).ToolTip = Trim$(CStr(sToolTip))
        End If
        If Not (ItemVisible(Index) = -1) Then
            Me.Refresh
        End If
        ItemSetText = True
        Exit Function
    End If
    ItemSetText = False
End Function

Public Function ItemGetText(ByVal sKey As String) As String
    Dim Index As Long
    Index = ItemIndex(sKey)
    If (Index > -1) Then
        ItemGetText = m_uItems(Index).Text
        Exit Function
    End If
    ItemGetText = ""
End Function

Public Function ItemGetKeyByIndex(ByVal iIndex As Long) As String
    If (iIndex > -1) And (iIndex <= UBound(m_uItems)) Then
        ItemGetKeyByIndex = m_uItems(iIndex).Key
        Exit Function
    End If
    ItemGetKeyByIndex = ""
End Function

Public Sub ItemSelect(ByVal sKey As String)
    Dim Index As Long
    Index = ItemIndex(sKey)
    If (Index > -1) Then
        Call ButtonClick(Index)
    End If
End Sub

Public Function ItemSelected() As String
    ItemSelected = m_uItems(m_lSelected).Key
End Function

Private Function ItemIndex(ByRef sKey As String) As Long
    Dim i As Long
    sKey = LCase$(Trim$(sKey))
    For i = 0 To UBound(m_uItems)
        If (m_uItems(i).Key = sKey) Then
            ItemIndex = i
            Exit Function
        End If
    Next
    ItemIndex = -1
End Function
Private Function ItemRelated(Index As Long, parentIndex As Long) As Boolean
    Dim itm As Long
    itm = m_uItems(Index).parent
    Do While (itm > 0)
        If (itm = parentIndex) Then
            ItemRelated = True
            Exit Do
        End If
        itm = m_uItems(itm).parent
    Loop
End Function
Private Function ItemChildCount(Index As Long) As Long
    Dim n As Long
    Dim i As Long
    For i = LBound(m_uItems) To UBound(m_uItems)
        If (m_uItems(i).parent = Index) Then
            n = n + 1
        End If
    Next i
    ItemChildCount = n
End Function
Private Function ItemVisible(Index As Long) As Long
    Dim i As Long
    For i = LBound(m_uButtons) To UBound(m_uButtons)
        If (m_uButtons(i).ref = Index) Then
            ItemVisible = i
            Exit Function
        End If
    Next
    ItemVisible = -1
End Function

'/****************************************************************************
' * Theme Functions and properties
' ****************************************************************************/

Private Sub ThemeInitialize(ByVal lhWnd As Long)
    Dim hTheme As Long
    Dim bThemed As Boolean
    Dim lPtrColorName As Long
    Dim lPtrThemeFile As Long
    Dim sThemeFile As String
    Dim sColorName As String
    Dim hRes As Long
    Dim iPos As Long
    Dim lBitsPixel As Long
    Dim lHDC As Long

    Call ThemeVersionInitialize

    If (m_eStyle = Office2003) Then
        If (m_bIsXP) Then
            On Error Resume Next
            hTheme = OpenThemeData(lhWnd, StrPtr("ExplorerBar"))
            If Not (hTheme = 0) Then
                ReDim bThemeFile(0 To 260 * 2) As Byte
                lPtrThemeFile = VarPtr(bThemeFile(0))
                ReDim bColorName(0 To 260 * 2) As Byte
                lPtrColorName = VarPtr(bColorName(0))
                hRes = GetCurrentThemeName(lPtrThemeFile, 260, lPtrColorName, 260, 0, 0)

                sThemeFile = bThemeFile
                iPos = InStr(sThemeFile, vbNullChar)
                If (iPos > 1) Then sThemeFile = Left(sThemeFile, iPos - 1)
                sColorName = bColorName
                iPos = InStr(sColorName, vbNullChar)
                If (iPos > 1) Then sColorName = Left(sColorName, iPos - 1)

                m_eTheme = xpCustom  'default theme
                If (LCase$(Right$(sThemeFile, 13)) = "luna.msstyles") Then
                    Select Case sColorName
                        Case "NormalColor": m_eTheme = xpBlue
                        Case "Metallic":    m_eTheme = xpSilver
                        Case "HomeStead":   m_eTheme = xpOlive
                    End Select
                End If

                Call CloseThemeData(hTheme)
            End If
        End If
    End If

    lHDC = GetDC(lhWnd)
    lBitsPixel = GetDeviceCaps(lHDC, BITSPIXEL)
    Call ReleaseDC(lhWnd, lHDC)
    m_bTrueColor = (lBitsPixel > 8)

End Sub

Private Sub ThemeVersionInitialize()
    Dim tOSV As OSVERSIONINFO
    tOSV.dwVersionInfoSize = Len(tOSV)
    Call GetVersionEx(tOSV)

    m_bIsXP = False
    m_bHasGradientAndTransparency = False

    If (tOSV.dwMajorVersion > 5) Then
        m_bIsXP = True
        m_bHasGradientAndTransparency = True
    ElseIf (tOSV.dwMajorVersion = 5) Then
        If (tOSV.dwMinorVersion >= 1) Then
            m_bIsXP = True
        End If
        m_bHasGradientAndTransparency = True
    ElseIf (tOSV.dwMajorVersion = 4) Then ' NT4 or 9x/ME/SE
        If (tOSV.dwMinorVersion >= 10) Then
            m_bHasGradientAndTransparency = True
        End If
    End If

End Sub

Friend Property Get ThemeLightColor() As Long
    Select Case m_eStyle
        Case BarStylesConstants.Office97
            ThemeLightColor = ColorTranslate(vb3DHighlight)
        Case BarStylesConstants.OfficeXP
            ThemeLightColor = GetSysColor(vb3DHighlight And &H1F&)
        Case BarStylesConstants.Office2003
            Select Case m_eTheme
                Case xpBlue:   ThemeLightColor = RGB(255, 255, 255)
                Case xpSilver: ThemeLightColor = RGB(255, 255, 255)
                'Case xpOlive:
                Case xpCustom: ThemeLightColor = GetSysColor(vb3DHighlight And &H1F&)
            End Select
        Case BarStylesConstants.Office2007
            ThemeLightColor = GetSysColor(vb3DHighlight And &H1F&)
    End Select
End Property
Friend Property Get ThemeDarkColor() As Long
    Select Case m_eStyle
        Case BarStylesConstants.Office97
            ThemeDarkColor = ColorTranslate(vbButtonShadow)
        Case BarStylesConstants.OfficeXP
            ThemeDarkColor = GetSysColor(vbButtonShadow And &H1F&)
        Case BarStylesConstants.Office2003
            If (m_bTrueColor) Then
                Select Case m_eTheme
                    Case xpBlue:   ThemeDarkColor = RGB(106, 140, 203)
                    Case xpSilver: ThemeDarkColor = RGB(110, 109, 143)
                    'case xpOlive:
                    Case xpCustom: ThemeDarkColor = ColorBlend(vbButtonShadow, vb3DHighlight, 180)
                End Select
            Else
                ThemeDarkColor = ColorTranslate(vbButtonShadow)
            End If
        Case BarStylesConstants.Office2007
            ThemeDarkColor = GetSysColor(vbButtonShadow And &H1F&)
    End Select
End Property
Friend Property Get ThemeDisabledIconColor() As Long
    ThemeDisabledIconColor = Me.ThemeTextDisabledColor
End Property
Friend Property Get ThemeTextDownEffect() As Boolean
    ThemeTextDownEffect = (m_eStyle = Office97)
End Property
Friend Property Get ThemeTextColor() As Long
    If (Me.Enabled) Then
        ThemeTextColor = GetSysColor(vbWindowText And &H1F&)
    Else
        ThemeTextColor = Me.ThemeTextDisabledColor
    End If
End Property
Friend Property Get ThemeTextHotColor() As Long
    If (Me.Enabled) Then
        ThemeTextHotColor = Me.ThemeTextColor
    Else
        ThemeTextHotColor = Me.ThemeTextDisabledColor
    End If
End Property
Friend Property Get ThemeTextDisabledColor() As Long
    ThemeTextDisabledColor = Me.ThemeDarkColor
End Property
Friend Property Get ThemeGradientColorStart() As Long
    If (m_bCustomBackground) Then
        ThemeGradientColorStart = ColorTranslate(m_lCustomBackgroundColor)
        Exit Property
    End If
    ' no custom bkg
    Select Case m_eStyle
        Case BarStylesConstants.Office97
            ThemeGradientColorStart = ColorTranslate(vbButtonFace)
        Case BarStylesConstants.OfficeXP
            If (m_bTrueColor) Then
                ThemeGradientColorStart = ColorBlend(vb3DLight, vbButtonFace)
            Else
                ThemeGradientColorStart = ColorTranslate(vbButtonFace)
            End If
        Case BarStylesConstants.Office2003
            If (m_bTrueColor) Then
                Select Case m_eTheme
                    Case xpBlue:   ThemeGradientColorStart = RGB(209, 227, 251)
                    Case xpSilver: ThemeGradientColorStart = RGB(249, 249, 255)
                    Case xpOlive:  ThemeGradientColorStart = RGB(247, 249, 225)
                    Case xpCustom: ThemeGradientColorStart = ColorBlend(vbButtonFace, vb3DHighlight, 24)
                End Select
            Else
                ThemeGradientColorStart = ColorTranslate(vbButtonFace)
            End If
        Case BarStylesConstants.Office2007
            If (m_bTrueColor) Then
                ThemeGradientColorStart = RGB(227, 239, 255)
            Else
                ThemeGradientColorStart = ColorTranslate(vbButtonFace)
            End If
    End Select
End Property
Friend Property Get ThemeGradientColorEnd() As Long
    If (m_bCustomBackground) Then
        ThemeGradientColorEnd = ColorTranslate(m_lCustomBackgroundColor)
        Exit Property
    End If
    ' no custom bkg
    Select Case m_eStyle
        Case BarStylesConstants.Office97
            ThemeGradientColorEnd = Me.ThemeGradientColorStart
        Case BarStylesConstants.OfficeXP
            ThemeGradientColorEnd = Me.ThemeGradientColorStart
        Case BarStylesConstants.Office2003
            If (m_bTrueColor) Then
                Select Case m_eTheme
                    Case xpBlue:   ThemeGradientColorEnd = RGB(129, 169, 226)
                    Case xpSilver: ThemeGradientColorEnd = RGB(159, 157, 185)
                    Case xpOlive:  ThemeGradientColorEnd = RGB(181, 197, 143)
                    Case xpCustom: ThemeGradientColorEnd = GetSysColor(vbButtonFace And &H1F&)
                End Select
            Else
                ThemeGradientColorEnd = ColorTranslate(vbButtonFace)
            End If
        Case BarStylesConstants.Office2007
            ThemeGradientColorEnd = RGB(177, 211, 255)
    End Select
End Property
Friend Property Get ThemeBackgroundHotColorStart() As Long
    Select Case m_eStyle
        Case BarStylesConstants.Office97
            ThemeBackgroundHotColorStart = ColorTranslate(vbButtonFace)
        Case BarStylesConstants.OfficeXP
            If (m_bTrueColor) Then
                ThemeBackgroundHotColorStart = ColorBlend(ColorBlend(vb3DHighlight, &HFFFFFF), vbHighlight, 178)
            Else
                ThemeBackgroundHotColorStart = ColorTranslate(vbButtonFace)
            End If
        Case BarStylesConstants.Office2003
            If (m_bTrueColor) Then
                Select Case m_eTheme
                    Case xpBlue:   ThemeBackgroundHotColorStart = RGB(253, 254, 211)
                    Case xpSilver: ThemeBackgroundHotColorStart = RGB(255, 239, 192)
                    Case xpOlive:  ThemeBackgroundHotColorStart = RGB(255, 245, 206) 'review
                    Case xpCustom: ThemeBackgroundHotColorStart = ColorBlend(vbHighlight, vb3DHighlight, 77)
                End Select
            Else
                ThemeBackgroundHotColorStart = ColorTranslate(vbButtonFace)
            End If
        Case BarStylesConstants.Office2007
            ThemeBackgroundHotColorStart = RGB(255, 245, 204)
    End Select
End Property
Friend Property Get ThemeBackgroundHotColorEnd() As Long
    Select Case m_eStyle
        Case BarStylesConstants.Office97
            ThemeBackgroundHotColorEnd = Me.ThemeBackgroundHotColorStart
        Case BarStylesConstants.OfficeXP
            ThemeBackgroundHotColorEnd = Me.ThemeBackgroundHotColorStart
        Case BarStylesConstants.Office2003
            If (m_bTrueColor) Then
                Select Case m_eTheme
                    Case xpBlue:   ThemeBackgroundHotColorEnd = RGB(253, 221, 152)
                    Case xpSilver: ThemeBackgroundHotColorEnd = RGB(255, 220, 115)
                    Case xpOlive:  ThemeBackgroundHotColorEnd = RGB(255, 207, 142) 'review
                    Case xpCustom: ThemeBackgroundHotColorEnd = ColorBlend(vbHighlight, vb3DHighlight, 84)
                End Select
            Else
                ThemeBackgroundHotColorEnd = Me.ThemeBackgroundHotColorStart
            End If
        Case BarStylesConstants.Office2007
            ThemeBackgroundHotColorEnd = RGB(255, 219, 117)
    End Select
End Property
Friend Property Get ThemeBackgroundCheckedColorStart() As Long
    If (m_bTrueColor) Then
        Select Case m_eStyle
            Case BarStylesConstants.Office97
                ThemeBackgroundCheckedColorStart = ColorBlend(vb3DHighlight, vbButtonFace, 220)
            Case BarStylesConstants.OfficeXP
                ThemeBackgroundCheckedColorStart = ColorBlend(vbHighlight, Me.ThemeGradientColorStart, 21)
            Case BarStylesConstants.Office2003
                Select Case m_eTheme
                    Case xpBlue:   ThemeBackgroundCheckedColorStart = RGB(251, 223, 128)
                    Case xpSilver: ThemeBackgroundCheckedColorStart = RGB(250, 218, 152)
                    Case xpOlive:  ThemeBackgroundCheckedColorStart = RGB(254, 211, 142) 'review
                    Case xpCustom: ThemeBackgroundCheckedColorStart = ColorBlend(Me.ThemeGradientColorStart, Me.ThemeBackgroundHotColorStart, 16)
                End Select
            Case BarStylesConstants.Office2007
                ThemeBackgroundCheckedColorStart = RGB(255, 207, 146)
        End Select
    Else
        ThemeBackgroundCheckedColorStart = ColorTranslate(vbButtonFace)
    End If
End Property
Friend Property Get ThemeBackgroundCheckedColorEnd() As Long
    Select Case m_eStyle
        Case BarStylesConstants.Office97
            ThemeBackgroundCheckedColorEnd = Me.ThemeBackgroundCheckedColorStart
        Case BarStylesConstants.OfficeXP
            ThemeBackgroundCheckedColorEnd = Me.ThemeBackgroundCheckedColorStart
        Case BarStylesConstants.Office2003
            If (m_bTrueColor) Then
                Select Case m_eTheme
                    Case xpBlue:   ThemeBackgroundCheckedColorEnd = RGB(245, 185, 74)
                    Case xpSilver: ThemeBackgroundCheckedColorEnd = RGB(229, 165, 33)
                    Case xpOlive:  ThemeBackgroundCheckedColorEnd = RGB(254, 145, 78) 'review
                    Case xpCustom: ThemeBackgroundCheckedColorEnd = Me.ThemeGradientColorStart
                End Select
            Else
                ThemeBackgroundCheckedColorEnd = ColorTranslate(vbButtonFace)
            End If
        Case BarStylesConstants.Office2007
            ThemeBackgroundCheckedColorEnd = RGB(255, 175, 73)
    End Select
End Property
Friend Property Get ThemeBackgroundCheckedHotColorStart() As Long
    If (m_bTrueColor) Then
        Select Case m_eStyle
            Case BarStylesConstants.Office97
                ThemeBackgroundCheckedHotColorStart = ColorTranslate(vbButtonFace)
            Case BarStylesConstants.OfficeXP
                ThemeBackgroundCheckedHotColorStart = ColorBlend(vb3DHighlight, vbHighlight)
            Case BarStylesConstants.Office2003
                Select Case m_eTheme
                    Case xpBlue:   ThemeBackgroundCheckedHotColorStart = RGB(251, 139, 89)
                    Case xpSilver: ThemeBackgroundCheckedHotColorStart = RGB(236, 176, 139)
                    Case xpOlive:  ThemeBackgroundCheckedHotColorStart = RGB(251, 139, 89) 'review
                    Case xpCustom: ThemeBackgroundCheckedHotColorStart = ColorBlend(vbHighlight, vb3DHighlight)
                End Select
            Case BarStylesConstants.Office2007
                ThemeBackgroundCheckedHotColorStart = RGB(252, 151, 61)
        End Select
    Else
        ThemeBackgroundCheckedHotColorStart = ColorTranslate(vbButtonFace)
    End If
End Property
Friend Property Get ThemeBackgroundCheckedHotColorEnd() As Long
    Select Case m_eStyle
        Case BarStylesConstants.Office97
            ThemeBackgroundCheckedHotColorEnd = Me.ThemeBackgroundCheckedHotColorStart
        Case BarStylesConstants.OfficeXP
            ThemeBackgroundCheckedHotColorEnd = Me.ThemeBackgroundCheckedHotColorStart
        Case BarStylesConstants.Office2003
            If (m_bTrueColor) Then
                Select Case m_eTheme
                    Case xpBlue:   ThemeBackgroundCheckedHotColorEnd = RGB(206, 47, 3)
                    Case xpSilver: ThemeBackgroundCheckedHotColorEnd = RGB(196, 103, 48)
                    Case xpOlive:  ThemeBackgroundCheckedHotColorEnd = RGB(206, 47, 3) 'review
                    Case xpCustom: ThemeBackgroundCheckedHotColorEnd = ColorBlend(vbHighlight, vb3DHighlight, 150)
                End Select
            Else
                ThemeBackgroundCheckedHotColorEnd = Me.ThemeBackgroundCheckedHotColorStart
            End If
        Case BarStylesConstants.Office2007
            ThemeBackgroundCheckedHotColorEnd = RGB(255, 184, 94)
    End Select
End Property
Friend Property Get ThemeBorderHotColor(borderPos As BorderColorPositionConstants) As Long
    Select Case m_eStyle
        Case BarStylesConstants.Office97
            Select Case borderPos
                Case BorderColorPositionConstants.bcLeft, BorderColorPositionConstants.bcTop
                    ThemeBorderHotColor = Me.ThemeLightColor
                Case BorderColorPositionConstants.bcRight, BorderColorPositionConstants.bcBottom
                    ThemeBorderHotColor = Me.ThemeDarkColor
            End Select
        Case BarStylesConstants.OfficeXP
            ThemeBorderHotColor = GetSysColor(vbHighlight And &H1F&)
        Case BarStylesConstants.Office2003
            If (m_bTrueColor) Then
                Select Case m_eTheme
                    Case xpBlue:   ThemeBorderHotColor = RGB(0, 0, 128)
                    Case xpSilver: ThemeBorderHotColor = RGB(75, 75, 111)
                    'case xpOlive:
                    Case xpCustom: ThemeBorderHotColor = GetSysColor(vbHighlight And &H1F&)
                End Select
            Else
                ThemeBorderHotColor = GetSysColor(vbHighlight And &H1F&)
            End If
        Case BarStylesConstants.Office2007
            ThemeBorderHotColor = RGB(255, 189, 105)
    End Select
End Property
Friend Property Get ThemeBorderCheckedColor(borderPos As BorderColorPositionConstants) As Long
    Select Case m_eStyle
        Case BarStylesConstants.Office97
            Select Case borderPos
                Case BorderColorPositionConstants.bcLeft, BorderColorPositionConstants.bcTop
                    ThemeBorderCheckedColor = Me.ThemeDarkColor
                Case BorderColorPositionConstants.bcRight, BorderColorPositionConstants.bcBottom
                    ThemeBorderCheckedColor = Me.ThemeLightColor
            End Select
        Case BarStylesConstants.OfficeXP
            ThemeBorderCheckedColor = GetSysColor(vbHighlight And &H1F&)
        Case BarStylesConstants.Office2003
            If (m_bTrueColor) Then
                Select Case m_eTheme
                    Case xpBlue:   ThemeBorderCheckedColor = RGB(0, 0, 128)
                    Case xpSilver: ThemeBorderCheckedColor = RGB(75, 75, 111)
                    'case xpOlive:
                    Case xpCustom: ThemeBorderCheckedColor = GetSysColor(vbHighlight And &H1F&)
                End Select
            Else
                ThemeBorderCheckedColor = GetSysColor(vbHighlight And &H1F&)
            End If
        Case BarStylesConstants.Office2007
            ThemeBorderCheckedColor = RGB(255, 189, 105) 'RGB(141, 141, 141)
    End Select
End Property
Friend Property Get ThemeBorderCheckedHotColor(borderPos As BorderColorPositionConstants) As Long
    Select Case m_eStyle
        Case BarStylesConstants.Office2007
            ThemeBorderCheckedHotColor = RGB(251, 140, 60)
        Case Else
            ThemeBorderCheckedHotColor = Me.ThemeBorderCheckedColor(borderPos)
    End Select
End Property
Friend Property Get ThemeInnerBorderHotColor(borderPos As BorderColorPositionConstants) As Long
    ThemeInnerBorderHotColor = TRANSPARENT
End Property
Friend Property Get ThemeInnerBorderCheckedColor(borderPos As BorderColorPositionConstants) As Long
    ThemeInnerBorderCheckedColor = TRANSPARENT
End Property
Friend Property Get ThemeInnerBorderCheckedHotColor(borderPos As BorderColorPositionConstants) As Long
    ThemeInnerBorderCheckedHotColor = TRANSPARENT
End Property


'/****************************************************************************
' * Start Subclass code - The programmer may call any of the following
' * Subclass_??? routines
' ****************************************************************************/

'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    On Error GoTo Errs
    'Parameters:
      'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
      'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
      'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If When And eMsgWhen.MSG_AFTER Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
Errs:
End Sub


'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
    On Error GoTo Errs
    'Parameters:
      'lng_hWnd  - The handle of the window to be subclassed
    'Returns;
      'The sc_aSubData() index
    Const CODE_LEN              As Long = 202                           'Length of the machine code in bytes
    Const FUNC_CWP              As String = "CallWindowProcA"           'We use CallWindowProc to call the original WndProc
    Const FUNC_EBM              As String = "EbMode"                    'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
    Const FUNC_SWL              As String = "SetWindowLongA"            'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
    Const MOD_USER              As String = "user32"                    'Location of the SetWindowLongA & CallWindowProc functions
    Const MOD_VBA5              As String = "vba5"                      'Location of the EbMode function if running VB5
    Const MOD_VBA6              As String = "vba6"                      'Location of the EbMode function if running VB6
    Const PATCH_01              As Long = 18                            'Code buffer offset to the location of the relative address to EbMode
    Const PATCH_02              As Long = 68                            'Address of the previous WndProc
    Const PATCH_03              As Long = 78                            'Relative address of SetWindowsLong
    Const PATCH_06              As Long = 116                           'Address of the previous WndProc
    Const PATCH_07              As Long = 121                           'Relative address of CallWindowProc
    Const PATCH_0A              As Long = 186                           'Address of the owner object
    Static aBuf(1 To CODE_LEN)  As Byte                                 'Static code buffer byte array
    Static pCWP                 As Long                                 'Address of the CallWindowsProc
    Static pEbMode              As Long                                 'Address of the EbMode IDE break/stop/running function
    Static pSWL                 As Long                                 'Address of the SetWindowsLong function
    Dim i                       As Long                                 'Loop index
    Dim j                       As Long                                 'Loop index
    Dim nSubIdx                 As Long                                 'Subclass data index
    Dim sHex                    As String                               'Hex code string

    'If it's the first time through here..
    If aBuf(1) = 0 Then

        'The hex pair machine code representation.
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
               "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
               "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
               "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

        'Convert the string from hex pairs to bytes and store in the static machine code buffer
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            aBuf(j) = Val("&H" & Mid$(sHex, i, 2))                      'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
            i = i + 2
        Loop                                                            'Next pair of hex characters

        'Get API function addresses
        If Subclass_InIDE Then                                          'If we're running in the VB IDE
            aBuf(16) = &H90                                               'Patch the code buffer to enable the IDE state code
            aBuf(17) = &H90                                               'Patch the code buffer to enable the IDE state code
            pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                       'Get the address of EbMode in vba6.dll
            If pEbMode = 0 Then                                           'Found?
                pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                     'VB5 perhaps
            End If
        End If

        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                            'Get the address of the CallWindowsProc function
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                            'Get the address of the SetWindowLongA function
        ReDim sc_aSubData(0 To 0) As tSubData                           'Create the first sc_aSubData element
    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If nSubIdx = -1 Then                                            'If an sc_aSubData element isn't being re-cycled
            nSubIdx = UBound(sc_aSubData()) + 1                           'Calculate the next element
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData          'Create a new sc_aSubData element
        End If

        Subclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)
        .hWnd = lng_hWnd                                                'Store the hWnd
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                   'Allocate memory for the machine code WndProc
        .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)      'Set our WndProc in place
        Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)          'Copy the machine code from the static byte array to the code array in sc_aSubData
        Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                    'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                 'Original WndProc address for CallWindowProc, call the original WndProc
        Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                       'Patch the relative address of the SetWindowLongA api function
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                 'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
        Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                       'Patch the relative address of the CallWindowProc api function
        Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                 'Patch the address of this object instance into the static machine code buffer
    End With
Errs:
End Function

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
    On Error GoTo Errs
    'Parameters:
      'lng_hWnd  - The handle of the window to stop being subclassed
    With sc_aSubData(zIdx(lng_hWnd))
        Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)             'Restore the original WndProc
        Call zPatchVal(.nAddrSub, PATCH_05, 0)                          'Patch the Table B entry count to ensure no further 'before' callbacks
        Call zPatchVal(.nAddrSub, PATCH_09, 0)                          'Patch the Table A entry count to ensure no further 'after' callbacks
        Call GlobalFree(.nAddrSub)                                      'Release the machine code memory
        .hWnd = 0                                                       'Mark the sc_aSubData element as available for re-use
        .nMsgCntB = 0                                                   'Clear the before table
        .nMsgCntA = 0                                                   'Clear the after table
        Erase .aMsgTblB                                                 'Erase the before table
        Erase .aMsgTblA                                                 'Erase the after table
    End With
Errs:
End Sub

'=======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    On Error GoTo Errs
    Dim nEntry  As Long                                                 'Message table entry index
    Dim nOff1   As Long                                                 'Machine code buffer offset 1
    Dim nOff2   As Long                                                 'Machine code buffer offset 2

    If uMsg = ALL_MESSAGES Then                                         'If all messages
        nMsgCnt = ALL_MESSAGES                                            'Indicates that all messages will callback
    Else                                                                'Else a specific message number
        Do While nEntry < nMsgCnt                                         'For each existing entry. NB will skip if nMsgCnt = 0
            nEntry = nEntry + 1

            If aMsgTbl(nEntry) = 0 Then                                 'This msg table slot is a deleted entry
                aMsgTbl(nEntry) = uMsg                                    'Re-use this entry
                Exit Sub                                                  'Bail
            ElseIf aMsgTbl(nEntry) = uMsg Then                          'The msg is already in the table!
                Exit Sub                                                  'Bail
            End If
        Loop                                                            'Next entry

        nMsgCnt = nMsgCnt + 1                                           'New slot required, bump the table entry count
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                    'Bump the size of the table.
        aMsgTbl(nMsgCnt) = uMsg                                         'Store the message number in the table
    End If

    If When = eMsgWhen.MSG_BEFORE Then                                  'If before
        nOff1 = PATCH_04                                                  'Offset to the Before table
        nOff2 = PATCH_05                                                  'Offset to the Before table entry count
    Else                                                                'Else after
        nOff1 = PATCH_08                                                  'Offset to the After table
        nOff2 = PATCH_09                                                  'Offset to the After table entry count
    End If

    If uMsg <> ALL_MESSAGES Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)                               'Patch the appropriate table entry count
Errs:
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc                                              'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
    On Error GoTo Errs
    'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0                                                  'Iterate through the existing sc_aSubData() elements
        With sc_aSubData(zIdx)
            If .hWnd = lng_hWnd Then                                    'If the hWnd of this element is the one we're looking for
                If Not bAdd Then                                        'If we're searching not adding
                    Exit Function                                       'Found
                End If
            ElseIf .hWnd = 0 Then                                       'If this an element marked for reuse.
                If bAdd Then                                            'If we're adding
                    Exit Function                                       'Re-use it
                End If
            End If
        End With
        zIdx = zIdx - 1                                                 'Decrement the index
    Loop

    'If Not bAdd Then
    '    Debug.Assert False                                             'hWnd not found, programmer error
    'End If
Errs:
'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    zSetTrue = True
    bValue = True
End Function

'/****************************************************************************
' * END Subclassing Code
' ****************************************************************************/

