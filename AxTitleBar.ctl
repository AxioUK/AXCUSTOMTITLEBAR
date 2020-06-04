VERSION 5.00
Begin VB.UserControl axCustomTitleBar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   630
   ScaleWidth      =   2715
   ToolboxBitmap   =   "AxTitleBar.ctx":0000
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      Picture         =   "AxTitleBar.ctx":0312
      ScaleHeight     =   495
      ScaleWidth      =   2535
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2535
      Begin VB.PictureBox PicL1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1770
         Picture         =   "AxTitleBar.ctx":1854
         ScaleHeight     =   210
         ScaleWidth      =   765
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   285
         Width           =   765
      End
      Begin VB.Label lblAxio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ax"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   330
         Left            =   450
         TabIndex        =   5
         Top             =   -30
         Width           =   345
      End
      Begin VB.Label lblLock 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CustomTitleBar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   810
         TabIndex        =   4
         Top             =   30
         Width           =   1560
      End
      Begin VB.Label lbVers 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2355
         TabIndex        =   3
         Top             =   -30
         Width           =   165
      End
      Begin VB.Label lblBy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modded by"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1005
         TabIndex        =   2
         Top             =   315
         Width           =   690
      End
   End
End
Attribute VB_Name = "axCustomTitleBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-CLASE-ORIGINAL---------------------------
'Cls Name: ClsCustomTitleBar
'Autor: Leandro Ascierto
'Date:  09/09/2018
'Web:   www.leandroascierto.com
'Note:  Custom Title bar for "windows 10"
'------------------------------------
'-UC-VERSION----------------------
'UC Name: axCustomTitleBar
'Editor: David Rojas [AxioUK]
'Date:  02/06/2020
'Note:  para usar como OCX y reducir codigo/depuración en proyectos grandes
'------------------------------------
Private Enum eMsgWhen                                                   'When to callback
  MSG_BEFORE = 1                                                        'Callback before the original WndProc
  MSG_AFTER = 2                                                         'Callback after the original WndProc
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                            'Callback before and after the original WndProc
End Enum

Private Enum eThunkType
    SubclassThunk = 0
    HookThunk = 1
    CallbackThunk = 2
End Enum

Private z_IDEflag                       As Long                         'Flag indicating we are in IDE
Private z_ScMem                         As Long                         'Thunk base address
Private z_scFunk                        As Collection                   'hWnd/thunk-address collection
Private z_hkFunk                        As Collection                   'hook/thunk-address collection
Private z_cbFunk                        As Collection                   'callback/thunk-address collection
Private Const IDX_INDEX                 As Long = 2                     'index of the subclassed hWnd OR hook type
Private Const IDX_CALLBACKORDINAL       As Long = 22                    ' Ubound(callback thunkdata)+1, index of the callback

Private Const IDX_WNDPROC               As Long = 9                     'Thunk data index of the original WndProc
Private Const IDX_BTABLE                As Long = 11                    'Thunk data index of the Before table
Private Const IDX_ATABLE                As Long = 12                    'Thunk data index of the After table
Private Const IDX_PARM_USER             As Long = 13                    'Thunk data index of the User-defined callback parameter data index
Private Const IDX_UNICODE               As Long = 75                    'Must be Ubound(subclass thunkdata)+1; index for unicode support
Private Const ALL_MESSAGES              As Long = -1                    'All messages callback
Private Const MSG_ENTRIES               As Long = 32                    'Number of msg table entries. Set to 1 if using ALL_MESSAGES for all subclassed windows

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProcW Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SendMessageA Lib "user32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SendMessageW Lib "user32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
'Private Declare Function IsWindowVisible Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Private Declare Function WaitMessage Lib "user32" () As Long
'Private Declare Function EndMenu Lib "user32.dll" () As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function AdjustWindowRectEx Lib "user32.dll" (ByRef lpRect As RECT, ByVal dsStyle As Long, ByVal bMenu As Long, ByVal dwEsStyle As Long) As Long
Private Declare Function MonitorFromWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
'Private Declare Function MapWindowPoints Lib "user32.dll" (ByVal hwndFrom As Long, ByVal hwndTo As Long, ByRef lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function GetDpiForMonitor Lib "Shcore.dll" (ByVal hMonitor As Long, ByVal dpiType As Long, ByRef dpiX As Long, ByRef dpiY As Long) As Long

'----------------User32 Api--------------
'Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'Private Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As PointAPI) As Long
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal hwnd As Long, ByVal lptpm As Any) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function InvalidateRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uiParam As Long, pvParam As Any, ByVal fWinIni As Long) As Long
Private Declare Function CopyImage Lib "user32.dll" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hwnd As Long) As Long

'----------------Gdi32 Api------------------
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetTextColor Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Private Declare Function GetBkColor Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function GdiAlphaBlend& Lib "gdi32" (ByVal hdc&, ByVal X&, ByVal Y&, ByVal dx&, ByVal dy&, ByVal hdcSrc&, ByVal SrcX&, ByVal SrcY&, ByVal SrcdX&, ByVal SrcdY&, ByVal lBlendFunction&)
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, ByRef pBMI As BITMAPINFO, ByVal iUsage As Long, ByRef ppvBits As Long, ByVal hSection As Long, ByVal dwOffset As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
'Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
'Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Private Const NULL_BRUSH As Long = 5

'----------------Kernel32 Api-----------------
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
'Private Declare Function LoadLibraryEx Lib "kernel32.dll" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
'Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

'----------------Shlwapi  Api-----------------
'Private Declare Function SHRegGetPathA Lib "shlwapi.dll" (ByVal hKey As Long, ByVal pcszSubKey As String, ByVal pcszValue As String, ByVal pszPath As String, ByVal dwFlags As Long) As Long
Private Declare Function SHGetValue Lib "shlwapi.dll" Alias "SHGetValueA" (ByVal hKey As Long, ByVal pszSubKey As String, ByVal pszValue As String, ByRef pdwType As Long, ByRef pvData As Any, ByRef pcbData As Long) As Long

'----------------UxTheme  Api-----------------
'Private Declare Function GetThemeStream Lib "UxTheme" (ByVal hTheme As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal iPropId As Long, ByRef ppvStream As Long, ByRef pcbStream As Long, ByVal hInst As Long) As Long
'Private Declare Function GetThemeRect Lib "UxTheme" (ByVal hTheme As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal iPropId As Long, ByRef pRect As RECT) As Long
'Private Declare Function OpenThemeData Lib "UxTheme" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
'Private Declare Function CloseThemeData Lib "UxTheme" (ByVal hTheme As Long) As Long
'Private Declare Function GetImmersiveUserColorSetPreference Lib "uxtheme.dll" Alias "#98" (ByVal bForceCheckRegistry As Long, ByVal bSkipCheckOnFail As Long) As Long
'Private Declare Function GetImmersiveColorTypeFromName Lib "uxtheme.dll" Alias "#96" (ByVal name As Long) As Long
'Private Declare Function GetImmersiveColorFromColorSetEx Lib "uxtheme.dll" Alias "#95" (ByVal dwImmersiveColorSet As Long, ByVal dwImmersiveColorType As Long, ByVal bIgnoreHighContrast As Long, ByVal dwHighContrastCacheMode As Long) As Long

'----------------Gdiplus  Api-----------------
'Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, ByRef Image As Long) As Long
'Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As Long
'Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As Long, ByVal InterpolationMode As Long) As Long
'Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal graphics As Long, ByVal PixelOffsetMode As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, ByRef graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As GDIPlusStartupInput, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
'Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "GdiPlus.dll" (ByVal mColor As Long, ByRef mBrush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal mBrush As Long) As Long
Private Declare Function GdipFillRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
'Private Declare Function GdipCloneBitmapArea Lib "GdiPlus.dll" (ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single, ByVal mPixelFormat As Long, ByVal mSrcBitmap As Long, ByRef mDstBitmap As Long) As Long
Private Declare Function GdipDrawLineI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipDrawRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long

'Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long
'Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Function DwmSetWindowAttribute Lib "dwmapi.dll" (ByVal hwnd As Long, ByVal dwAttribute As Long, ByRef pvAttribute As Long, ByVal cbAttribute As Long) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As PointAPI) As Long

'Private Declare Function GetDCEx Lib "user32.dll" (ByVal hwnd As Long, ByVal hrgnclip As Long, ByVal fdwOptions As Long) As Long
Private Const DCX_WINDOW As Long = &H1&
Private Const DCX_INTERSECTRGN As Long = &H80&

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long

Private Const UnitPixel                             As Long = 2
Private Const GdiPlusVersion                        As Long = 1&
Private Const QualityModeHigh                       As Long = 2&
Private Const InterpolationModeNearestNeighbor      As Long = QualityModeHigh + 3
Private Const PixelOffsetModeHalf                   As Long = QualityModeHigh + 2

'----------------WindowMensage----------------
Private Const WM_CLOSE                  As Long = &H10
Private Const WM_STYLECHANGED           As Long = &H7D
Private Const WM_HOTKEY                 As Long = &H312
Private Const WM_MENUSELECT             As Long = &H11F
Private Const WM_ENTERIDLE              As Long = &H121
Private Const WM_COMMAND                As Long = &H111
Private Const WM_MENUCHAR               As Long = &H120
Private Const WM_DESTROY                As Long = &H2
Private Const WM_SETCURSOR              As Long = &H20
Private Const WM_SIZE                   As Long = &H5
Private Const WM_SYSCOMMAND             As Long = &H112
Private Const WM_NCPAINT                As Long = &H85
Private Const WM_NCACTIVATE             As Long = &H86
Private Const WM_ACTIVATE               As Long = &H6
Private Const WM_ERASEBKGND             As Long = &H14
Private Const WM_SETICON                As Long = &H80
Private Const WM_NCLBUTTONDOWN          As Long = &HA1
Private Const WM_NCLBUTTONUP            As Long = &HA2
Private Const WM_NCRBUTTONUP            As Long = &HA5
Private Const WM_NCHITTEST              As Long = &H84
Private Const WM_NCMOUSEMOVE            As Long = &HA0
Private Const WM_GETICON                As Long = &H7F
Private Const WM_PAINT                  As Long = &HF&
Private Const WM_NCCALCSIZE             As Long = &H83
Private Const WM_SYSCHAR                As Long = &H106
Private Const WM_SYSKEYDOWN             As Long = &H104
Private Const WM_GETMINMAXINFO          As Long = &H24
Private Const WM_ACTIVATEAPP            As Long = &H1C
Private Const WM_ASIVAAANDAR            As Long = &HAE          'NO IDENTIFICADA
Private Const WM_MYCUSTOMMSG            As Long = &H555
Private Const WM_MYCUSTOMWPARAM         As Long = &H1111
Private Const WM_MDIGETACTIVE           As Long = &H229
Private Const WM_MDISETMENU             As Long = &H230
Private Const WM_SHOWWINDOW             As Long = &H18
Private Const WM_CREATE                 As Long = &H1
Private Const WM_MOUSEMOVE              As Long = &H200
Private Const WM_LBUTTONDBLCLK          As Long = &H203
Private Const WM_LBUTTONDOWN            As Long = &H201
Private Const WM_LBUTTONUP              As Long = &H202
Private Const WM_NCLBUTTONDBLCLK        As Long = &HA3
Private Const WM_ENTERSIZEMOVE          As Long = &H231
Private Const WM_EXITSIZEMOVE           As Long = &H232
'Private Const WM_WINDOWPOSCHANGED       As Long = &H47
Private Const WM_WININICHANGE           As Long = &H1A
Private Const WM_NCMOUSELEAVE           As Long = &H2A2
Private Const WM_SIZING                 As Long = &H214
Private Const WM_DPICHANGED             As Long = &H2E0
Private Const WM_DISPLAYCHANGE          As Long = &H7E

'----------------WindowStyle and Ex----------
Private Const WS_MAXIMIZEBOX            As Long = &H10000
Private Const WS_MINIMIZEBOX            As Long = &H20000
Private Const WS_THICKFRAME             As Long = &H40000
Private Const WS_CAPTION                As Long = &HC00000
Private Const WS_BORDER                 As Long = &H800000
Private Const WS_SYSMENU                As Long = &H80000
Private Const WS_DLGFRAME               As Long = &H400000
Private Const WS_VISIBLE                As Long = &H10000000
Private Const WS_CHILD                  As Long = &H40000000
Private Const WS_MINIMIZE               As Long = &H20000000
Private Const WS_MAXIMIZE               As Long = &H1000000

Private Const WS_EX_CLIENTEDGE          As Long = &H200&
Private Const WS_EX_MDICHILD            As Long = &H40&
Private Const WS_EX_APPWINDOW           As Long = &H40000
Private Const WS_EX_WINDOWEDGE          As Long = &H100&
Private Const WS_EX_TRANSPARENT         As Long = &H20&

Private Const GWL_STYLE                 As Long = -16
Private Const GWL_EXSTYLE               As Long = -20

'----------------DrawText--------------------
Private Const DT_CALCRECT               As Long = &H400
Private Const DT_VCENTER                As Long = &H4
Private Const DT_CENTER                 As Long = &H1
Private Const DT_SINGLELINE             As Long = &H20
Private Const DT_WORD_ELLIPSIS          As Long = &H40000

'----------------HitTest---------------------
Private Const HTCLIENT                  As Long = 1
Private Const HTCAPTION                 As Long = 2
Private Const HTSYSMENU                 As Long = 3
Private Const HTGROWBOX                 As Long = 4
Private Const HTMENU                    As Long = 5
Private Const HTMINBUTTON               As Long = 8
Private Const HTMAXBUTTON               As Long = 9
Private Const HTLEFT                    As Long = 10
Private Const HTRIGHT                   As Long = 11
Private Const HTTOP                     As Long = 12
Private Const HTTOPLEFT                 As Long = 13
Private Const HTTOPRIGHT                As Long = 14
Private Const HTBOTTOM                  As Long = 15
Private Const HTBOTTOMLEFT              As Long = 16
Private Const HTBOTTOMRIGHT             As Long = 17
Private Const HTBORDER                  As Long = 18
Private Const HTCLOSE                   As Long = 20
Private Const HTHELP                    As Long = 21
Private Const HTMenuPlus                As Long = 30

'---------------SYSCOMMAND-------------------
Private Const SC_CLOSE                  As Long = &HF060&
Private Const SC_MINIMIZE               As Long = &HF020&
Private Const SC_RESTORE                As Long = &HF120&
Private Const SC_MAXIMIZE               As Long = &HF030&

Private Const HORZRES           As Long = 8
Private Const VERTRES           As Long = 10
Private Const HORZSIZE          As Long = 4
Private Const VERTSIZE          As Long = 6

Private Const THEME_MIDDLE_IMAGES          As Integer = 3
Private Const THEME_MIDDLE_IMAGES2         As Integer = 4  ' Inactive Window
Private Const THEME_LEFT_IMAGES            As Integer = 5
Private Const THEME_LEFT_IMAGES2           As Integer = 6  ' Inactive Window
Private Const THEME_CLOSE_RIGHT_IMAGES     As Integer = 7
Private Const THEME_CLOSE_RIGHT_IMAGES2    As Integer = 8  ' Inactive Window
Private Const THEME_CLOSE_SINGLE_IMAGES    As Integer = 9
Private Const THEME_CLOSE_SINGLE_IMAGES2   As Integer = 10 ' Inactive Window
Private Const THEME_CLOSE_HOT_FRAME        As Integer = 11 ' Win7, Vista
Private Const THEME_HOT_FRAME              As Integer = 16 ' Win7, Vista
Private Const THEME_CLOSE_ICONS            As Integer = 11
Private Const THEME_HELP_ICONS             As Integer = 15 ' Win7, Vista: 17
Private Const THEME_MAX_ICONS              As Integer = 19 ' Win7, Vista: 21
Private Const THEME_MIN_ICONS              As Integer = 23 ' Win7, Vista: 25
Private Const THEME_RESTORE_ICONS          As Integer = 27 ' Win7, Vista: 29
Private Const THEME_CLOSE_SMALL_IMAGES     As Integer = 37 ' Win7, Vista: 45
Private Const THEME_CLOSE_SMALL_IMAGES2    As Integer = 38 ' Win7, Vista: 46  Inactive Window
Private Const THEME_CLOSE_SMALL_HOT_FRAME  As Integer = 47  ' Win7, Vista
Private Const THEME_CLOSE_SMALL_ICONS      As Integer = 39 ' Win7, Vista: 48
Private Const THEME_CLOSE_ICONS_CLEAR      As Integer = 64 ' Win7, Vista:
Private Const THEME_HELP_ICONS_CLEAR       As Integer = 68 ' Win7, Vista:
Private Const THEME_MAX_ICONS_CLEAR        As Integer = 72 ' Win7, Vista:
Private Const THEME_MIN_ICONS_CLEAR        As Integer = 76 ' Win7, Vista:
Private Const THEME_RESTORE_ICONS_CLEAR    As Integer = 80 ' Win7, Vista:

'---------------Others-----------------------
Private Const ICON_SMALL                As Long = 0
Private Const ICON_BIG                  As Long = 1
Private Const LR_COPYFROMRESOURCE       As Long = &H4000
Private Const IMAGE_ICON                As Long = 1
Private Const TPM_RETURNCMD             As Long = &H100&
Private Const CS_DROPSHADOW             As Long = &H20000
Private Const GCL_STYLE                 As Long = -26
Private Const BI_RGB                    As Long = 0&
Private Const DIB_RGB_COLORS            As Long = 0&
Private Const COLOR_MENU                As Long = 4
Private Const LOGPIXELSX                As Long = 88

Private Const LOAD_LIBRARY_AS_DATAFILE As Long = &H2
Private Const TMT_DISKSTREAM = 213
Private Const TMT_ATLASRECT = 8002
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const DWMNCRP_ENABLED = 2
Private Const DWMWA_NCRENDERING_POLICY = 2
Private Const SPI_GETNONCLIENTMETRICS As Long = 41
Private Const SPI_GETWORKAREA = 48
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2 '
  
Private Type GDIPlusStartupInput
    GdiPlusVersion                      As Long
    DebugEventCallback                  As Long
    SuppressBackgroundThread            As Long
    SuppressExternalCodecs              As Long
End Type

Private Type PointAPI
    X                                   As Long
    Y                                   As Long
End Type

Private Type RECT
    Left                                As Long
    Top                                 As Long
    Right                               As Long
    Bottom                              As Long
End Type

Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

Private Type BITMAPINFOHEADER
    biSize                              As Long
    biWidth                             As Long
    biHeight                            As Long
    biPlanes                            As Integer
    biBitCount                          As Integer
    biCompression                       As Long
    biSizeImage                         As Long
    biXPelsPerMeter                     As Long
    biYPelsPerMeter                     As Long
    biClrUsed                           As Long
    biClrImportant                      As Long
End Type

Private Type RGBQUAD
    rgbBlue                             As Byte
    rgbGreen                            As Byte
    rgbRed                              As Byte
    rgbReserved                         As Byte
End Type

Private Type BITMAPINFO
    bmiHeader                           As BITMAPINFOHEADER
    bmiColors                           As RGBQUAD
End Type

Private Type MINMAXINFO
    ptReserved                          As PointAPI
    ptMaxSize                           As PointAPI
    ptMaxPosition                       As PointAPI
    ptMinTrackSize                      As PointAPI
    ptMaxTrackSize                      As PointAPI
End Type

Private Type WINDOWPOS
   hwnd                                 As Long
   hWndInsertAfter                      As Long
   X                                    As Long
   Y                                    As Long
   cx                                   As Long
   cy                                   As Long
   Flags                                As Long
End Type

Private Type NCCALCSIZE_PARAMS
   rgrc(0 To 2)                         As RECT
   lppos                                As Long
End Type

Private Type LOGFONT
    lfHeight                            As Long
    lfWidth                             As Long
    lfEscapement                        As Long
    lfOrientation                       As Long
    lfWeight                            As Long
    lfItalic                            As Byte
    lfUnderline                         As Byte
    lfStrikeOut                         As Byte
    lfCharSet                           As Byte
    lfOutPrecision                      As Byte
    lfClipPrecision                     As Byte
    lfQuality                           As Byte
    lfPitchAndFamily                    As Byte
    lfFaceName(1 To 32)                 As Byte
End Type

Private Type NONCLIENTMETRICS
  cbSize As Long
  iBorderWidth As Long
  iScrollWidth As Long
  iScrollHeight As Long
  iCaptionWidth As Long
  iCaptionHeight As Long
  lfCaptionFont As LOGFONT
  iSMCaptionWidth As Long
  iSMCaptionHeight As Long
  lfSMCaptionFont As LOGFONT
  iMenuWidth As Long
  iMenuHeight As Long
  lfMenuFont As LOGFONT
  lfStatusFont As LOGFONT
  lfMessageFont As LOGFONT
End Type

Private Type ThemePart
    hImage As Long
    Widht As Long
    Height As Long
End Type

Private tMinIco         As ThemePart
Private tMaxIco         As ThemePart
Private tRestoreIco     As ThemePart
Private tHelpIco        As ThemePart
Private tCloseIco       As ThemePart
Private tLeftBtn        As ThemePart
Private tMidleBtn       As ThemePart
Private tRightBtn       As ThemePart

Private FrmH            As Long
Private m_Focus         As Long
Private m_Icon          As Long
Private m_Caption       As String
Private WinStyle        As Long
Private WinExStyle      As Long
Private m_bSubClass     As Boolean
Private m_DrawTitle     As Boolean
Private TimerEnabled    As Boolean
Private m_Hittest       As Long

Private mCBW            As Long
Private mCBH            As Long
Private mBtnW           As Long
Private hFont           As Long
Private m_ShowIcon      As Boolean
Private m_WhiteIcons    As Boolean
Private m_ShowTitlebar  As Boolean
Private m_ShowCaption   As Boolean
Private m_UseSystemTheme As Boolean
Private m_TitleBarBackColor             As OLE_COLOR
Private m_TitleBarBackColorDesactivate  As OLE_COLOR

Private m_MinFrmSize    As PointAPI
Private m_MaxFrmSize    As PointAPI
Private DibDC           As Long
Private hOldBMP         As Long
Private BorderWidth     As Long
Private BorderHeight    As Long
Private GdipToken       As Long
Private nScale          As Double
Private IsWinZoomed     As Boolean
Private diff            As Long

Private Const BorderPixels  As Long = 5

Private lHwnd   As Long
Private lFrm     As Form

Public Event ENTERSIZEMOVE()
Public Event EXITSIZEMOVE()

'-The following routines are exclusively for the ssc_subclass routines----------------------------
Private Function ssc_Subclass(ByVal lng_hWnd As Long, _
    Optional ByVal lParamUser As Long = 0, _
    Optional ByVal nOrdinal As Long = 1, _
    Optional ByVal oCallback As Object = Nothing, _
    Optional ByVal bIdeSafety As Boolean = True, _
    Optional ByVal bUnicode As Boolean = False) As Boolean                          'Subclass the specified window handle

    '*************************************************************************************************
    '* lng_hWnd   - Handle of the window to subclass
    '* lParamUser - Optional, user-defined callback parameter
    '* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
    '* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl for design-time subclassing
    '* bUnicode - Optional, if True, Unicode API calls will be made to the window vs ANSI calls
    '*************************************************************************************************
    '* cSelfSub - self-subclassing class template
    '* Paul_Caton@hotmail.com
    '* Copyright free, use and abuse as you see fit.
    '*
    '* v1.0 Re-write of the SelfSub/WinSubHook-2 submission to Planet Source Code............ 20060322
    '* v1.1 VirtualAlloc memory to prevent Data Execution Prevention faults on Win64......... 20060324
    '* v1.2 Thunk redesigned to handle unsubclassing and memory release...................... 20060325
    '* v1.3 Data array scrapped in favour of property accessors.............................. 20060405
    '* v1.4 Optional IDE protection added
    '*      User-defined callback parameter added
    '*      All user routines that pass in a hWnd get additional validation
    '*      End removed from zError.......................................................... 20060411
    '* v1.5 Added nOrdinal parameter to ssc_Subclass
    '*      Switched machine-code array from Currency to Long................................ 20060412
    '* v1.6 Added an optional callback target object
    '*      Added an IsBadCodePtr on the callback address in the thunk prior to callback..... 20060413
    '*************************************************************************************************
    ' Subclassing procedure must be declared identical to the one at the end of this class (Sample at Ordinal #1)

    ' \\LaVolpe - reworked routine a bit, revised the ASM to allow auto-unsubclass on WM_DESTROY
    
    Dim z_Sc(0 To IDX_UNICODE)          As Long                                     'Thunk machine-code initialised here
    Const CODE_LEN                      As Long = 4 * IDX_UNICODE                   'Thunk length in bytes
    
    Const MEM_LEN                       As Long = CODE_LEN + (8 * (MSG_ENTRIES))    'Bytes to allocate per thunk, data + code + msg tables
    Const PAGE_RWX                      As Long = &H40&                             'Allocate executable memory
    Const MEM_COMMIT                    As Long = &H1000&                           'Commit allocated memory
    Const MEM_RELEASE                   As Long = &H8000&                           'Release allocated memory flag
    Const IDX_EBMODE                    As Long = 3                                 'Thunk data index of the EbMode function address
    Const IDX_CWP                       As Long = 4                                 'Thunk data index of the CallWindowProc function address
    Const IDX_SWL                       As Long = 5                                 'Thunk data index of the SetWindowsLong function address
    Const IDX_FREE                      As Long = 6                                 'Thunk data index of the VirtualFree function address
    Const IDX_BADPTR                    As Long = 7                                 'Thunk data index of the IsBadCodePtr function address
    Const IDX_OWNER                     As Long = 8                                 'Thunk data index of the Owner object's vTable address
    Const IDX_CALLBACK                  As Long = 10                                'Thunk data index of the callback method address
    Const IDX_EBX                       As Long = 16                                'Thunk code patch index of the thunk data
    Const GWL_WNDPROC                   As Long = -4                                'SetWindowsLong WndProc index
    Const WNDPROC_OFF                   As Long = &H38                              'Thunk offset to the WndProc execution address
    Const SUB_NAME                      As String = "ssc_Subclass"                  'This routine's name
    
    Dim nAddr                           As Long
    Dim nID                             As Long
    Dim nMyID                           As Long

    If IsWindow(lng_hWnd) = 0 Then                                                  'Ensure the window handle is valid
        zError SUB_NAME, "Invalid window handle"
        Exit Function
    End If
    
    nMyID = GetCurrentProcessId                                                     'Get this process's ID
    GetWindowThreadProcessId lng_hWnd, nID                                          'Get the process ID associated with the window handle
    If nID <> nMyID Then                                                            'Ensure that the window handle doesn't belong to another process
        zError SUB_NAME, "Window handle belongs to another process"
        Exit Function
    End If
      
    If oCallback Is Nothing Then Set oCallback = Me                                 'If the user hasn't specified the callback owner
    
    nAddr = zAddressOf(oCallback, nOrdinal)                                         'Get the address of the specified ordinal method
    If nAddr = 0 Then                                                               'Ensure that we've found the ordinal method
        zError SUB_NAME, "Callback method not found"
        Exit Function
    End If
        
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                        'Allocate executable memory
    
    If z_ScMem <> 0 Then                                                            'Ensure the allocation succeeded
  
        If z_scFunk Is Nothing Then Set z_scFunk = New Collection                   'If this is the first time through, do the one-time initialization
        On Error GoTo CatchDoubleSub                                                'Catch double subclassing
        z_scFunk.Add z_ScMem, "h" & lng_hWnd                                        'Add the hWnd/thunk-address to the collection
        On Error GoTo 0
        
        ' \\Tai Chi Minh Ralph Eastwood - fixed bug where the MSG_AFTER was not being honored
        ' \\LaVolpe - modified thunks to allow auto-unsubclassing when WM_DESTROY received
        z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(16) = &H12345678: z_Sc(17) = &HF63103FF: z_Sc(18) = &H750C4339: z_Sc(19) = &H7B8B4A38: z_Sc(20) = &H95E82C: z_Sc(21) = &H7D810000: z_Sc(22) = &H228&: z_Sc(23) = &HC70C7500: z_Sc(24) = &H20443: z_Sc(25) = &H5E90000: z_Sc(26) = &H39000000: z_Sc(27) = &HF751475: z_Sc(28) = &H25E8&: z_Sc(29) = &H8BD23100: z_Sc(30) = &H6CE8307B: z_Sc(31) = &HFF000000: z_Sc(32) = &H10C2610B: z_Sc(33) = &HC53FF00: z_Sc(34) = &H13D&: z_Sc(35) = &H85BE7400: z_Sc(36) = &HE82A74C0: z_Sc(37) = &H2&: z_Sc(38) = &H75FFE5EB: z_Sc(39) = &H2C75FF30: z_Sc(40) = &HFF2875FF: z_Sc(41) = &H73FF2475: z_Sc(42) = &H1053FF24: z_Sc(43) = &H811C4589: z_Sc(44) = &H13B&: z_Sc(45) = &H39727500:
        z_Sc(46) = &H6D740473: z_Sc(47) = &H2473FF58: z_Sc(48) = &HFFFFFC68: z_Sc(49) = &H873FFFF: z_Sc(50) = &H891453FF: z_Sc(51) = &H7589285D: z_Sc(52) = &H3045C72C: z_Sc(53) = &H8000&: z_Sc(54) = &H8920458B: z_Sc(55) = &H4589145D: z_Sc(56) = &HC4816124: z_Sc(57) = &H4&: z_Sc(58) = &H8B1862FF: z_Sc(59) = &H853AE30F: z_Sc(60) = &H810D78C9: z_Sc(61) = &H4C7&: z_Sc(62) = &H28458B00: z_Sc(63) = &H2975AFF2: z_Sc(64) = &H2873FF52: z_Sc(65) = &H5A1C53FF: z_Sc(66) = &H438D1F75: z_Sc(67) = &H144D8D34: z_Sc(68) = &H1C458D50: z_Sc(69) = &HFF3075FF: z_Sc(70) = &H75FF2C75: z_Sc(71) = &H873FF28: z_Sc(72) = &HFF525150: z_Sc(73) = &H53FF2073: z_Sc(74) = &HC328C328
        
        z_Sc(IDX_EBX) = z_ScMem                                                     'Patch the thunk data address
        z_Sc(IDX_INDEX) = lng_hWnd                                                  'Store the window handle in the thunk data
        z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                       'Store the address of the before table in the thunk data
        z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)             'Store the address of the after table in the thunk data
        z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                         'Store the callback owner's object address in the thunk data
        z_Sc(IDX_CALLBACK) = nAddr                                                  'Store the callback address in the thunk data
        z_Sc(IDX_PARM_USER) = lParamUser                                            'Store the lParamUser callback parameter in the thunk data
        
        ' \\LaVolpe - validate unicode request & cache unicode usage
        If bUnicode Then bUnicode = (IsWindowUnicode(lng_hWnd) <> 0&)
        z_Sc(IDX_UNICODE) = bUnicode                                                'Store whether the window is using unicode calls or not
        
        ' \\LaVolpe - added extra parameter "bUnicode" to the zFnAddr calls
        z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree", bUnicode)               'Store the VirtualFree function address in the thunk data
        z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", bUnicode)            'Store the IsBadCodePtr function address in the thunk data
        
        Debug.Assert zInIDE
        If bIdeSafety = True And z_IDEflag = 1 Then                                 'If the user wants IDE protection
            z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode", bUnicode)                  'Store the EbMode function address in the thunk data
        End If
    
        ' \\LaVolpe - use ANSI for non-unicode usage, else use WideChar calls
        If bUnicode Then
            z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcW", bUnicode)          'Store CallWindowProc function address in the thunk data
            z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongW", bUnicode)           'Store the SetWindowLong function address in the thunk data
            z_Sc(IDX_UNICODE) = 1
            RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                        'Copy the thunk code/data to the allocated memory
            nAddr = SetWindowLongW(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc
        Else
            z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA", bUnicode)          'Store CallWindowProc function address in the thunk data
            z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA", bUnicode)           'Store the SetWindowLong function address in the thunk data
            RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                        'Copy the thunk code/data to the allocated memory
            nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc
        End If
        If nAddr = 0 Then                                                           'Ensure the new WndProc was set correctly
            zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
            GoTo ReleaseMemory
        End If
        'Store the original WndProc address in the thunk data
        RtlMoveMemory z_ScMem + IDX_WNDPROC * 4, VarPtr(nAddr), 4&                  'z_Sc(IDX_WNDPROC) = nAddr
        ssc_Subclass = True                                                         'Indicate success
    Else
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
    End If

    Exit Function                                                                   'Exit ssc_Subclass
    
CatchDoubleSub:
    zError SUB_NAME, "Window handle is already subclassed"
      
ReleaseMemory:
    VirtualFree z_ScMem, 0, MEM_RELEASE                                             'ssc_Subclass has failed after memory allocation, so release the memory
    
End Function


'Terminate all subclassing
Private Sub ssc_Terminate()
    ' can be made public. Releases all subclassing
    ' can be removed and zTerminateThunks can be called directly
    zTerminateThunks SubclassThunk
End Sub


'UnSubclass the specified window handle
Private Sub ssc_UnSubclass(ByVal lng_hWnd As Long)
    ' can be made public. Releases a specific subclass
    ' can be removed and zUnThunk can be called directly
    zUnThunk lng_hWnd, SubclassThunk
End Sub


'Add the message value to the window handle's specified callback table
Private Sub ssc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    ' Note: can be removed if not needed and zAddMsg can be called directly
    
    If IsBadCodePtr(zMap_VFunction(lng_hWnd, SubclassThunk)) = 0 Then               'Ensure that the thunk hasn't already released its memory
        If When And MSG_BEFORE Then                                                 'If the message is to be added to the before original WndProc table...
            zAddMsg uMsg, IDX_BTABLE                                                'Add the message to the before table
        End If
        If When And MSG_AFTER Then                                                  'If message is to be added to the after original WndProc table...
            zAddMsg uMsg, IDX_ATABLE                                                'Add the message to the after table
        End If
    End If
    
End Sub


'Delete the message value from the window handle's specified callback table
Private Sub ssc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    ' Note: can be removed if not needed and zDelMsg can be called directly
    If IsBadCodePtr(zMap_VFunction(lng_hWnd, SubclassThunk)) = 0 Then               'Ensure that the thunk hasn't already released its memory
        If When And MSG_BEFORE Then                                                 'If the message is to be deleted from the before original WndProc table...
            zDelMsg uMsg, IDX_BTABLE                                                'Delete the message from the before table
        End If
        If When And MSG_AFTER Then                                                  'If the message is to be deleted from the after original WndProc table...
            zDelMsg uMsg, IDX_ATABLE                                                'Delete the message from the after table
        End If
    End If
    
End Sub


'Call the original WndProc
Private Function ssc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' Note: can be removed if you do not use this function inside of your window procedure
    
    If IsBadCodePtr(zMap_VFunction(lng_hWnd, SubclassThunk)) = 0 Then               'Ensure that the thunk hasn't already released its memory
        If zData(IDX_UNICODE) Then
            ssc_CallOrigWndProc = CallWindowProcW(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        Else
            ssc_CallOrigWndProc = CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        End If
    End If
    
End Function


'Get the subclasser lParamUser callback parameter
Private Function zGet_lParamUser(ByVal hWnd_Hook_ID As Long, vType As eThunkType) As Long
    'Note: can be removed if you never need to retrieve/update your user-defined paramter. See ssc_Subclass
    
    If vType <> CallbackThunk Then
        If IsBadCodePtr(zMap_VFunction(hWnd_Hook_ID, vType)) = 0 Then               'Ensure that the thunk hasn't already released its memory
            zGet_lParamUser = zData(IDX_PARM_USER)                                  'Get the lParamUser callback parameter
        End If
    End If
    
End Function


'Let the subclasser lParamUser callback parameter
Private Sub zSet_lParamUser(ByVal hWnd_Hook_ID As Long, vType As eThunkType, NewValue As Long)
    'Note: can be removed if you never need to retrieve/update your user-defined paramter. See ssc_Subclass
    
    If vType <> CallbackThunk Then
        If IsBadCodePtr(zMap_VFunction(hWnd_Hook_ID, vType)) = 0 Then               'Ensure that the thunk hasn't already released its memory
            zData(IDX_PARM_USER) = NewValue                                         'Set the lParamUser callback parameter
        End If
    End If
    
End Sub


'Add the message to the specified table of the window handle
Private Sub zAddMsg(ByVal uMsg As Long, ByVal nTable As Long)

    Dim nCount                          As Long                                     'Table entry count
    Dim nBase                           As Long                                     'Remember z_ScMem
    Dim i                               As Long                                     'Loop index
    
    nBase = z_ScMem                                                                 'Remember z_ScMem so that we can restore its value on exit
    z_ScMem = zData(nTable)                                                         'Map zData() to the specified table
    
    If uMsg = ALL_MESSAGES Then                                                     'If ALL_MESSAGES are being added to the table...
        nCount = ALL_MESSAGES                                                       'Set the table entry count to ALL_MESSAGES
    Else
        nCount = zData(0)                                                           'Get the current table entry count
        If nCount >= MSG_ENTRIES Then                                               'Check for message table overflow
            zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
            GoTo Bail
        End If
    
        For i = 1 To nCount                                                         'Loop through the table entries
            If zData(i) = 0 Then                                                    'If the element is free...
                zData(i) = uMsg                                                     'Use this element
                GoTo Bail                                                           'Bail
            ElseIf zData(i) = uMsg Then                                             'If the message is already in the table...
                GoTo Bail                                                           'Bail
            End If
        Next i                                                                      'Next message table entry
    
        nCount = i                                                                  'On drop through: i = nCount + 1, the new table entry count
        zData(nCount) = uMsg                                                        'Store the message in the appended table entry
    End If
    
    zData(0) = nCount                                                               'Store the new table entry count
Bail:
    z_ScMem = nBase                                                                 'Restore the value of z_ScMem
    
End Sub


'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long)

    Dim nCount                          As Long                                     'Table entry count
    Dim nBase                           As Long                                     'Remember z_ScMem
    Dim i                               As Long                                     'Loop index
    
    nBase = z_ScMem                                                                 'Remember z_ScMem so that we can restore its value on exit
    z_ScMem = zData(nTable)                                                         'Map zData() to the specified table
    
    If uMsg = ALL_MESSAGES Then                                                     'If ALL_MESSAGES are being deleted from the table...
        zData(0) = 0                                                                'Zero the table entry count
    Else
        nCount = zData(0)                                                           'Get the table entry count
        
        For i = 1 To nCount                                                         'Loop through the table entries
            If zData(i) = uMsg Then                                                 'If the message is found...
                zData(i) = 0                                                        'Null the msg value -- also frees the element for re-use
                GoTo Bail                                                           'Bail
            End If
        Next i                                                                      'Next message table entry
        
        zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
    End If
      
Bail:
    z_ScMem = nBase                                                                 'Restore the value of z_ScMem
    
End Sub


'-SelfCallback code------------------------------------------------------------------------------------
'-The following routines are exclusively for the scb_SetCallbackAddr routines----------------------------
Private Function scb_SetCallbackAddr(ByVal nParamCount As Long, _
    Optional ByVal nOrdinal As Long = 1, _
    Optional ByVal oCallback As Object = Nothing, _
    Optional ByVal bIdeSafety As Boolean = True) As Long   'Return the address of the specified callback thunk
    '*************************************************************************************************
    '* nParamCount  - The number of parameters that will callback
    '* nOrdinal     - Callback ordinal number, the final private method is ordinal 1, the second last is ordinal 2, etc...
    '* oCallback    - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety   - Optional, set to false to disable IDE protection.
    '*************************************************************************************************
    ' Callback procedure must return a Long even if, per MSDN, the callback procedure is a Sub vs Function
    ' The number of parameters are dependent on the individual callback procedures
    
    Const MEM_LEN                       As Long = IDX_CALLBACKORDINAL * 4 + 4       'Memory bytes required for the callback thunk
    Const PAGE_RWX                      As Long = &H40&                             'Allocate executable memory
    Const MEM_COMMIT                    As Long = &H1000&                           'Commit allocated memory
    Const SUB_NAME                      As String = "scb_SetCallbackAddr"           'This routine's name
    Const INDX_OWNER                    As Long = 0
    Const INDX_CALLBACK                 As Long = 1
    Const INDX_EBMODE                   As Long = 2
    Const INDX_BADPTR                   As Long = 3
    Const INDX_EBX                      As Long = 5
    Const INDX_PARAMS                   As Long = 12
    Const INDX_PARAMLEN                 As Long = 17

    Dim z_Cb()                          As Long                                     'Callback thunk array
    Dim nCallback                       As Long
      
    If z_cbFunk Is Nothing Then
        Set z_cbFunk = New Collection                                               'If this is the first time through, do the one-time initialization
    Else
        On Error Resume Next                                                        'Catch already initialized?
        z_ScMem = z_cbFunk.Item("h" & nOrdinal)                                     'Test it
        If Err = 0 Then
            scb_SetCallbackAddr = z_ScMem + 16                                      'we had this one, just reference it
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    If nParamCount < 0 Then                                                         ' validate parameters
        zError SUB_NAME, "Invalid Parameter count"
        Exit Function
    End If
    
    If oCallback Is Nothing Then Set oCallback = Me                                 'If the user hasn't specified the callback owner
    nCallback = zAddressOf(oCallback, nOrdinal)                                     'Get the callback address of the specified ordinal
    If nCallback = 0 Then
        zError SUB_NAME, "Callback address not found."
        Exit Function
    End If
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                        'Allocate executable memory
        
    If z_ScMem = 0& Then
        zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError          ' oops
        Exit Function
    End If
    z_cbFunk.Add z_ScMem, "h" & nOrdinal                                            'Add the callback/thunk-address to the collection
        
    ReDim z_Cb(0 To IDX_CALLBACKORDINAL) As Long                                    'Allocate for the machine-code array
    
    ' Create machine-code array
    z_Cb(4) = &HBB60E089: z_Cb(6) = &H73FFC589: z_Cb(7) = &HC53FF04: z_Cb(8) = &H7B831F75: z_Cb(9) = &H20750008: z_Cb(10) = &HE883E889: z_Cb(11) = &HB9905004: z_Cb(13) = &H74FF06E3: z_Cb(14) = &HFAE2008D: z_Cb(15) = &H53FF33FF: z_Cb(16) = &HC2906104: z_Cb(18) = &H830853FF: z_Cb(19) = &HD87401F8: z_Cb(20) = &H4589C031: z_Cb(21) = &HEAEBFC
    
    z_Cb(INDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", False)
    z_Cb(INDX_OWNER) = ObjPtr(oCallback)                                            'Set the Owner
    z_Cb(INDX_CALLBACK) = nCallback                                                 'Set the callback address
    z_Cb(IDX_CALLBACKORDINAL) = nOrdinal                                            'Cache ordinal used for zTerminateThunks
      
    Debug.Assert zInIDE
    If bIdeSafety = True And z_IDEflag = 1 Then                                     'If the user wants IDE protection
        z_Cb(INDX_EBMODE) = zFnAddr("vba6", "EbMode", False)                        'EbMode Address
    End If
        
    z_Cb(INDX_PARAMS) = nParamCount                                                 'Set the parameter count
    z_Cb(INDX_PARAMLEN) = nParamCount * 4                                           'Set the number of stck bytes to release on thunk return
      
    '\\LaVolpe - redirect address to proper location in virtual memory. Was: z_Cb(INDX_EBX) = VarPtr(z_Cb(INDX_OWNER))
    z_Cb(INDX_EBX) = z_ScMem                                                        'Set the data address relative to virtual memory pointer
        
    RtlMoveMemory z_ScMem, VarPtr(z_Cb(INDX_OWNER)), MEM_LEN                        'Copy thunk code to executable memory
    scb_SetCallbackAddr = z_ScMem + 16                                              'Thunk code start address
    
End Function


Private Sub scb_ReleaseCallback(ByVal nOrdinal As Long)
    ' can be made public. Releases a specific callback
    ' can be removed and zUnThunk can be called directly
    zUnThunk nOrdinal, CallbackThunk
End Sub


Private Sub scb_TerminateCallbacks()
    ' can be made public. Releases all callbacks
    ' can be removed and zTerminateThunks can be called directly
    zTerminateThunks CallbackThunk
End Sub


'========================================================================
' COMMON USE ROUTINES
'-The following routines are used for each of the three types of thunks
'========================================================================

'Map zData() to the thunk address for the specified window handle
Private Function zMap_VFunction(ByVal vFuncTarget As Long, vType As eThunkType) As Long
    
    ' vFuncTarget is one of the following, depending on vType
    '   - Subclassing:  the hWnd of the window subclassed
    '   - Hooking:      the hook type created
    '   - Callbacks:    the ordinal of the callback
    
    Dim thunkCol As Collection
    
    If vType = CallbackThunk Then
        Set thunkCol = z_cbFunk
    ElseIf vType = HookThunk Then
        Set thunkCol = z_hkFunk
    ElseIf vType = SubclassThunk Then
        Set thunkCol = z_scFunk
    Else
        zError "zMap_Vfunction", "Invalid thunk type passed"
        Exit Function
    End If
    
    If thunkCol Is Nothing Then
        zError "zMap_VFunction", "Thunk hasn't been initialized"
    Else
        On Error GoTo Catch
        z_ScMem = thunkCol("h" & vFuncTarget)                                       'Get the thunk address
        zMap_VFunction = z_ScMem
    End If
    Exit Function                                                                   'Exit returning the thunk address
    
Catch:
    zError "zMap_VFunction", "Thunk type for ID of " & vFuncTarget & " does not exist"
    
End Function


'Error handler
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String)
    ' \\LaVolpe -  Note. These two lines can be rem'd out if you so desire. But don't remove the routine
    App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
    MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub


'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String, ByVal asUnicode As Boolean) As Long

    If asUnicode Then
        zFnAddr = GetProcAddress(GetModuleHandleW(StrPtr(sDLL)), sProc)             'Get the specified procedure address
    Else
        zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                     'Get the specified procedure address
    End If
    Debug.Assert zFnAddr                                                            'In the IDE, validate that the procedure address was located
    ' ^^ FYI VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")
    
End Function


'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
    ' Note: used both in subclassing and hooking routines
    
    Dim bSub                            As Byte                                     'Value we expect to find pointed at by a vTable method entry
    Dim bVal                            As Byte
    Dim nAddr                           As Long                                     'Address of the vTable
    Dim i                               As Long                                     'Loop index
    Dim j                               As Long                                     'Loop limit
  
    RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                               'Get the address of the callback object's instance
    If Not zProbe(nAddr + &H1C, i, bSub) Then                                       'Probe for a Class method
        If Not zProbe(nAddr + &H6F8, i, bSub) Then                                  'Probe for a Form method
            ' \\LaVolpe - Added propertypage offset
            If Not zProbe(nAddr + &H710, i, bSub) Then                              'Probe for a PropertyPage method
                If Not zProbe(nAddr + &H7A4, i, bSub) Then                          'Probe for a UserControl method
                    Exit Function                                                   'Bail...
                End If
            End If
        End If
    End If
  
    i = i + 4                                                                       'Bump to the next entry
    j = i + 1024                                                                    'Set a reasonable limit, scan 256 vTable entries
    Do While i < j
        RtlMoveMemory VarPtr(nAddr), i, 4                                           'Get the address stored in this vTable entry
    
        If IsBadCodePtr(nAddr) Then                                                 'Is the entry an invalid code address?
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4                 'Return the specified vTable entry address
            Exit Do                                                                 'Bad method signature, quit loop
        End If

        RtlMoveMemory VarPtr(bVal), nAddr, 1                                        'Get the byte pointed to by the vTable entry
        If bVal <> bSub Then                                                        'If the byte doesn't match the expected value...
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4                 'Return the specified vTable entry address
            Exit Do                                                                 'Bad method signature, quit loop
        End If
    
        i = i + 4                                                                   'Next vTable entry
    Loop
    
End Function


'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean

    Dim bVal                        As Byte
    Dim nAddr                       As Long
    Dim nLimit                      As Long
    Dim nEntry                      As Long
  
    nAddr = nStart                                                                  'Start address
    nLimit = nAddr + 32                                                             'Probe eight entries
    Do While nAddr < nLimit                                                         'While we've not reached our probe depth
        RtlMoveMemory VarPtr(nEntry), nAddr, 4                                      'Get the vTable entry
    
        If nEntry <> 0 Then                                                         'If not an implemented interface
            RtlMoveMemory VarPtr(bVal), nEntry, 1                                   'Get the value pointed at by the vTable entry
            If bVal = &H33 Or bVal = &HE9 Then                                      'Check for a native or pcode method signature
                nMethod = nAddr                                                     'Store the vTable entry
                bSub = bVal                                                         'Store the found method signature
                zProbe = True                                                       'Indicate success
                Exit Do                                                             'Return
            End If
        End If
    
        nAddr = nAddr + 4                                                           'Next vTable entry
    Loop
    
End Function


Private Function zInIDE() As Long
    ' This is only run in IDE; it is never run when compiled
    z_IDEflag = 1
    zInIDE = z_IDEflag
End Function


Private Property Get zData(ByVal nIndex As Long) As Long
    ' retrieves stored value from virtual function's memory location
    RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
End Property


Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)
    ' sets value in virtual function's memory location
    RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
End Property


Private Sub zUnThunk(ByVal thunkID As Long, ByVal vType As eThunkType)
    ' Releases a specific subclass, hook or callback
    ' thunkID depends on vType:
    '   - Subclassing:  the hWnd of the window subclassed
    '   - Hooking:      the hook type created
    '   - Callbacks:    the ordinal of the callback

    Const IDX_SHUTDOWN                  As Long = 1
    Const MEM_RELEASE                   As Long = &H8000&                           'Release allocated memory flag
    
    If zMap_VFunction(thunkID, vType) Then
        Select Case vType
            Case SubclassThunk
                If IsBadCodePtr(z_ScMem) = 0 Then                                   'Ensure that the thunk hasn't already released its memory
                    zData(IDX_SHUTDOWN) = 1                                         'Set the shutdown indicator
                    zDelMsg ALL_MESSAGES, IDX_BTABLE                                'Delete all before messages
                    zDelMsg ALL_MESSAGES, IDX_ATABLE                                'Delete all after messages
                    '\\LaVolpe - Force thunks to replace original window procedure handle. Without this, app can crash when a window is subclassed multiple times simultaneously
                    If zData(IDX_UNICODE) Then                                      'Force window procedure handle to be replaced
                        SendMessageW thunkID, 0&, 0&, ByVal 0&
                    Else
                        SendMessageA thunkID, 0&, 0&, ByVal 0&
                    End If
                End If
                z_scFunk.Remove "h" & thunkID                                       'Remove the specified thunk from the collection
            Case HookThunk
                If IsBadCodePtr(z_ScMem) = 0 Then                                   'Ensure that the thunk hasn't already released its memory
                    zData(IDX_SHUTDOWN) = 1                                         'Set the shutdown indicator
                    zData(IDX_ATABLE) = 0                                           ' want no more After messages
                    zData(IDX_BTABLE) = 0                                           ' want no more Before messages
                End If
                z_hkFunk.Remove "h" & thunkID                                       'Remove the specified thunk from the collection
            Case CallbackThunk
                If IsBadCodePtr(z_ScMem) = 0 Then                                   'Ensure that the thunk hasn't already released its memory
                    VirtualFree z_ScMem, 0, MEM_RELEASE                             'Release allocated memory
                End If
                z_cbFunk.Remove "h" & thunkID                                       'Remove the specified thunk from the collection
        End Select
    End If

End Sub


Private Sub zTerminateThunks(ByVal vType As eThunkType)
    ' Removes all thunks of a specific type: subclassing, hooking or callbacks
    Dim i                               As Long
    Dim thunkCol                        As Collection
    
    Select Case vType
        Case SubclassThunk
            Set thunkCol = z_scFunk
        Case HookThunk
            Set thunkCol = z_hkFunk
        Case CallbackThunk
            Set thunkCol = z_cbFunk
        Case Else
            Exit Sub
    End Select
    
    If Not (thunkCol Is Nothing) Then                                               'Ensure that hooking has been started
        With thunkCol
            For i = .Count To 1 Step -1                                             'Loop through the collection of hook types in reverse order
                z_ScMem = .Item(i)                                                  'Get the thunk address
                If IsBadCodePtr(z_ScMem) = 0 Then                                   'Ensure that the thunk hasn't already released its memory
                    Select Case vType
                        Case SubclassThunk
                            zUnThunk zData(IDX_INDEX), SubclassThunk                'Unsubclass
                        Case HookThunk
                            zUnThunk zData(IDX_INDEX), HookThunk                    'Unhook
                        Case CallbackThunk
                            zUnThunk zData(IDX_CALLBACKORDINAL), CallbackThunk      'Release callback
                    End Select
                End If
            Next i                                                                  'Next member of the collection
        End With
        Set thunkCol = Nothing                                                      'Destroy the hook/thunk-address collection
    End If

End Sub

' Incia el Timer.
Private Sub TimerSet(lngLapse As Long)
    If TimerEnabled = False Then
        Call SetTimer(FrmH, ObjPtr(Me) + 1, lngLapse, scb_SetCallbackAddr(4, 2))
        TimerEnabled = True
    End If
End Sub


' Detiene el Timer.
Private Sub TimerKill()
    If TimerEnabled Then
        Call KillTimer(FrmH, ObjPtr(Me) + 1)
        TimerEnabled = False
    End If
End Sub


'=====================================================================================================================================================================
'Rutina del Skin
'=====================================================================================================================================================================

Public Property Get MinFormWidth() As Long: MinFormWidth = m_MinFrmSize.X: End Property
Public Property Get MinFormHeight() As Long: MinFormHeight = m_MinFrmSize.Y: End Property
Public Property Get MaxFormWidth() As Long: MaxFormWidth = m_MaxFrmSize.X: End Property
Public Property Get MaxFormHeight() As Long: MaxFormHeight = m_MaxFrmSize.Y: End Property
Public Property Let MinFormWidth(Value As Long): m_MinFrmSize.X = Value: End Property
Public Property Let MinFormHeight(Value As Long): m_MinFrmSize.Y = Value: End Property
Public Property Let MaxFormWidth(Value As Long): m_MaxFrmSize.X = Value: End Property
Public Property Let MaxFormHeight(Value As Long): m_MaxFrmSize.Y = Value: End Property

Public Property Get GetSystemTitleBarColor() As Long
    GetSystemTitleBarColor = DWM_AccentColor And &HFFFFFF
End Property

Public Property Get UseSystemTheme() As Boolean
    UseSystemTheme = m_UseSystemTheme
End Property

Public Property Let UseSystemTheme(NewUseSystemTheme As Boolean)
    m_UseSystemTheme = NewUseSystemTheme
    If NewUseSystemTheme Then
        m_TitleBarBackColor = DWM_AccentColor And &HFFFFFF
        m_WhiteIcons = Not IsLightColor(m_TitleBarBackColor)
        m_TitleBarBackColorDesactivate = vbWhite
        m_ShowTitlebar = True
        PropertyChanged "UseSystemTheme"
        Call Refresh
    End If
End Property

Public Property Get ShowCaption() As Boolean
    ShowCaption = m_ShowCaption
    'DrawTitle = m_ShowCaption
End Property

Public Property Let ShowCaption(ByVal NewShowCaption As Boolean)
    m_ShowCaption = NewShowCaption
    PropertyChanged "ShowCaption"
    Call Refresh
End Property

Public Property Get DrawTitle() As Boolean
    DrawTitle = m_DrawTitle
End Property

Public Property Let DrawTitle(NewValue As Boolean)
    m_DrawTitle = NewValue
    Call Refresh
End Property

Public Property Get TitleBarBackColorDesactivate() As OLE_COLOR
    TitleBarBackColorDesactivate = m_TitleBarBackColorDesactivate
End Property

Public Property Let TitleBarBackColorDesactivate(ByVal NewTitleBarBackColorDesactivate As OLE_COLOR)
    m_TitleBarBackColorDesactivate = NewTitleBarBackColorDesactivate
    OleTranslateColor NewTitleBarBackColorDesactivate, 0, m_TitleBarBackColorDesactivate
    PropertyChanged "TitleBarBackColorDesactivate"
    Call Refresh
End Property

Public Property Get TitleBarBackColor() As OLE_COLOR
    TitleBarBackColor = m_TitleBarBackColor
End Property

Public Property Let TitleBarBackColor(ByVal NewTitleBarBackColor As OLE_COLOR)
    OleTranslateColor NewTitleBarBackColor, 0, m_TitleBarBackColor
    m_TitleBarBackColor = NewTitleBarBackColor
    PropertyChanged "TitleBarBackColor"
    Call Refresh
End Property

Public Property Get ShowTitlebar() As Boolean
    ShowTitlebar = m_ShowTitlebar
    WhiteIcons = m_ShowTitlebar
End Property

Public Property Let ShowTitlebar(ByVal NewShowTitlebar As Boolean)
    m_ShowTitlebar = NewShowTitlebar
    PropertyChanged "ShowTitlebar"
    Call Refresh
End Property

Public Property Get WhiteIcons() As Boolean
    WhiteIcons = m_WhiteIcons
End Property

Public Property Let WhiteIcons(ByVal NewWhiteIcons As Boolean)
    m_WhiteIcons = NewWhiteIcons
    PropertyChanged "WhiteIcons"
    Call Refresh
End Property

Public Property Get ShowIcon() As Boolean
    ShowIcon = m_ShowIcon
End Property

Public Property Let ShowIcon(ByVal NewShowIcon As Boolean)
    m_ShowIcon = NewShowIcon
    If FrmH Then
        If NewShowIcon Then
            If m_Icon Then DestroyIcon m_Icon
            m_Icon = GetWindowIcon(FrmH, 16 * nScale)
        Else
            If m_Icon Then DestroyIcon m_Icon: m_Icon = 0
        End If
        RedrawWindow FrmH, ByVal 0&, ByVal 0&, &H1
    End If

    PropertyChanged "ShowIcon"
    Call Refresh
End Property

Public Property Get TitleBarWidth() As Long
    Dim Rec As RECT
    GetClientRect FrmH, Rec
    TitleBarWidth = Rec.Right - mCBW
End Property

Public Property Get TitleBarHeight() As Long
    TitleBarHeight = mCBH
End Property

Public Property Get ControlBoxWidth() As Long
    ControlBoxWidth = mCBW
End Property

Public Property Get ControlBoxHeight() As Long
    ControlBoxHeight = mCBH
End Property

Public Property Let OnTop(NewValue As Boolean)
    If NewValue Then
        SetWindowPos FrmH, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos FrmH, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
End Property

Public Sub DragForm()
    Call ReleaseCapture
    Call SendMessage(FrmH, &HA1, 2, 0&)
End Sub

' Función que retorna si la ventana es del estilo indicado.
Private Function IsStyle(WS_Style As Long) As Boolean
    IsStyle = WinStyle = (WinStyle Or WS_Style)
End Function

' Función que retorna si la ventana es del estilo extendido indicado.
Private Function IsExStyle(WS_EX_Style As Long) As Boolean
    IsExStyle = WinExStyle = (WinExStyle Or WS_EX_Style)
End Function

Private Sub Refresh()
    If FrmH Then
        PaintBuffer
        RedrawWindow FrmH, ByVal 0&, ByVal 0&, &H1
    End If
End Sub

Private Sub InitGDI()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = GdiPlusVersion
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
End Sub
  
Private Sub TerminateGDI()
    Call GdiplusShutdown(GdipToken)
End Sub

Private Function IsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress As Long
    Call CopyMemory(lAddress, ByVal lpArray, &H4)
    IsArrayDim = Not (lAddress = 0)
End Function

' Función principal que llama a subclasificar dicha ventana.
Private Function HookForm() As Boolean     '(lHwnd As Long) As Boolean
    
    Dim Rec As RECT
    Dim PT As PointAPI
    Dim hdc As Long
    Dim NCM As NONCLIENTMETRICS
    Dim lStyle As Long
    
    Set lFrm = UserControl.Parent
    lHwnd = lFrm.hwnd
    
    nScale = GetWindowsDPI
    
    NCM.cbSize = Len(NCM)
    SystemParametersInfo SPI_GETNONCLIENTMETRICS, 0, NCM, 0

    hFont = CreateFontIndirect(NCM.lfCaptionFont)
 
    mCBH = 31 * nScale
    mCBW = 138 * nScale
    mBtnW = mCBW / 3
    BorderWidth = BorderPixels * nScale
    BorderHeight = BorderPixels * nScale
    
    If m_bSubClass Then UnHookForm 'FrmH

    InitGDI
    
    If m_ShowIcon Then
        If m_Icon Then DestroyIcon m_Icon
        m_Icon = GetWindowIcon(lHwnd, 16 * nScale)
    End If
    
    'Call LoadTheme(lHwnd)
    
    FrmH = lHwnd
    m_Caption = GetWindowCaption(FrmH)
  
    ' Obtiene los estilos de la ventana.
    WinStyle = GetWindowLong(FrmH, GWL_STYLE)
    WinExStyle = GetWindowLong(FrmH, GWL_EXSTYLE)

    Call PaintBuffer

    Call DwmSetWindowAttribute(lHwnd, DWMWA_NCRENDERING_POLICY, DWMNCRP_ENABLED, 4)
    Call DwmSetWindowAttribute(lHwnd, 5, DWMNCRP_ENABLED, 4)
    
    lStyle = WS_BORDER Or WS_THICKFRAME
    If IsStyle(WS_MINIMIZEBOX) Then lStyle = lStyle Or WS_MINIMIZEBOX
    If IsStyle(WS_MAXIMIZEBOX) Then lStyle = lStyle Or WS_MAXIMIZEBOX
    
    SetWindowLongA lHwnd, GWL_EXSTYLE, WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW
    SetWindowLongA lHwnd, GWL_STYLE, lStyle 'And WS_CAPTION
       
    If ssc_Subclass(FrmH, , 1) Then                                                 ' Subclasifica la ventana principal.
        ssc_AddMsg FrmH, WM_NCHITTEST, MSG_AFTER
        ssc_AddMsg FrmH, WM_SIZE, MSG_BEFORE
        ssc_AddMsg FrmH, WM_LBUTTONDOWN, MSG_BEFORE
        ssc_AddMsg FrmH, WM_LBUTTONUP, MSG_BEFORE
        ssc_AddMsg FrmH, WM_LBUTTONDBLCLK, MSG_BEFORE
        ssc_AddMsg FrmH, WM_NCACTIVATE, MSG_BEFORE_AFTER
        ssc_AddMsg FrmH, WM_ACTIVATE, MSG_BEFORE_AFTER
        ssc_AddMsg FrmH, WM_ACTIVATEAPP, MSG_BEFORE_AFTER
        ssc_AddMsg FrmH, WM_NCLBUTTONUP, MSG_BEFORE
        ssc_AddMsg FrmH, WM_NCRBUTTONUP, MSG_BEFORE
        ssc_AddMsg FrmH, WM_NCLBUTTONDOWN, MSG_BEFORE
        ssc_AddMsg FrmH, WM_NCLBUTTONDBLCLK, MSG_BEFORE
        ssc_AddMsg FrmH, WM_NCCALCSIZE, MSG_BEFORE
        ssc_AddMsg FrmH, WM_NCPAINT, MSG_AFTER
        ssc_AddMsg FrmH, WM_STYLECHANGED, MSG_AFTER
        ssc_AddMsg FrmH, WM_PAINT, MSG_BEFORE_AFTER
        ssc_AddMsg FrmH, WM_GETMINMAXINFO, MSG_BEFORE_AFTER
        ssc_AddMsg FrmH, WM_WININICHANGE, MSG_AFTER
        ssc_AddMsg FrmH, WM_NCMOUSELEAVE, MSG_AFTER
        ssc_AddMsg FrmH, WM_SIZING, MSG_AFTER
        ssc_AddMsg FrmH, WM_DPICHANGED, MSG_AFTER
        ssc_AddMsg FrmH, WM_DISPLAYCHANGE, MSG_AFTER
        ssc_AddMsg FrmH, WM_ENTERSIZEMOVE, MSG_AFTER
        ssc_AddMsg FrmH, WM_EXITSIZEMOVE, MSG_AFTER
        m_bSubClass = True
        
        SetWindowPos FrmH, 0, 0, 0, 0, 0, 551
        
        HookForm = True
    End If
        
End Function

' Detiene todas las subclasificaciones.
Public Sub UnHookForm()   '(hwnd As Long)
    Dim i As Integer
    
    TimerKill
        
    If m_bSubClass Then
        ssc_UnSubclass lHwnd
        m_bSubClass = False
    End If
    
    ssc_Terminate
    scb_TerminateCallbacks

    TerminateGDI
    
    If hFont Then DeleteObject hFont: hFont = 0
    
    If DibDC <> 0 Then DeleteDIB
    If m_Icon Then DestroyIcon m_Icon: m_Icon = 0
      
    FrmH = 0
End Sub

' Rutina que obtiene el ícono de la ventana (éste debe ser destruido).
Private Function GetWindowIcon(hwnd As Long, Size As Long) As Long

    Dim hIcon As Long

    If Size <= 20 Then
        hIcon = SendMessage(hwnd, WM_GETICON, ICON_SMALL, ByVal 0)
        If hIcon = 0 Then
            hIcon = SendMessage(hwnd, WM_GETICON, ICON_BIG, ByVal 0)
        End If
    Else
        hIcon = SendMessage(hwnd, WM_GETICON, ICON_BIG, ByVal 0)
        If hIcon = 0 Then
            hIcon = SendMessage(hwnd, WM_GETICON, ICON_SMALL, ByVal 0)
        End If
    End If
        
    If hIcon <> 0 Then
        GetWindowIcon = CopyImage(hIcon, IMAGE_ICON, Size, Size, LR_COPYFROMRESOURCE)
    End If
    
End Function

' Obtiene el caption de la ventana.
Private Function GetWindowCaption(ByVal hwnd As Long) As String
    Dim sBuff As String
    sBuff = String(255, Chr$(0))
    GetWindowText hwnd, sBuff, 255
    GetWindowCaption = Left$(sBuff, InStr(sBuff, Chr$(0)) - 1)
End Function

'Retrieves the signed x-coordinate from the specified LPARAM value.
Public Function Get_X_lParam(ByVal lParam As Long) As Long
    Get_X_lParam = lParam And &H7FFF&
    If lParam And &H8000& Then Get_X_lParam = Get_X_lParam Or &HFFFF8000
End Function

'Retrieves the signed y-coordinate from the specified LPARAM value.
Public Function Get_Y_lParam(ByVal lParam As Long) As Long
    Get_Y_lParam = (lParam And &H7FFF0000) \ &H10000
    If lParam And &H80000000 Then Get_Y_lParam = Get_Y_lParam Or &HFFFF8000
End Function

' Retorna el valor LoWord.
Private Function LoWord(ByVal Numero As Long) As Long
    LoWord = Numero And &HFFFF&
End Function

' Retorna el valor HiWord.
Private Function HiWord(ByVal Numero As Long) As Long
    HiWord = Numero \ &H10000 And &HFFFF&
End Function

Private Sub UserControl_InitProperties()
    
m_TitleBarBackColor = DWM_AccentColor And &HFFFFFF
m_WhiteIcons = Not IsLightColor(m_TitleBarBackColor)
m_TitleBarBackColorDesactivate = vbWhite
m_ShowCaption = True
m_ShowIcon = True

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
           Call DragForm
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'ReadProperties Bag-------------------
With PropBag
    m_ShowCaption = .ReadProperty("ShowCaption", True)
    m_ShowIcon = .ReadProperty("ShowIcon", True)
    m_ShowTitlebar = .ReadProperty("ShowTitlebar", True)
    m_UseSystemTheme = .ReadProperty("UseSystemTheme", False)
    m_WhiteIcons = .ReadProperty("WhiteIcons", False)
    m_TitleBarBackColor = .ReadProperty("TitleBarBackColor", &HFFC0C0)
    m_TitleBarBackColorDesactivate = .ReadProperty("TitleBarBackColorDesactivate", &HE0E0E0)
End With


Dim IsInIDE
IsInIDE = App.LogMode
If IsInIDE = 0 Then
    'Running In IDE
    If Not Ambient.UserMode Then
        Call UnHookForm
    Else
        Call HookForm
    End If
ElseIf IsInIDE = 1 Then
    'Running Compiled
    Call HookForm
End If

lbVers.Caption = "2"

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
'WriteProperties Bag-------------------
    Call .WriteProperty("ShowCaption", m_ShowCaption, True)
    Call .WriteProperty("ShowIcon", m_ShowIcon, True)
    Call .WriteProperty("ShowTitlebar", m_ShowTitlebar, True)
    Call .WriteProperty("UseSystemTheme", m_UseSystemTheme, False)
    Call .WriteProperty("WhiteIcons", m_WhiteIcons, False)
    Call .WriteProperty("TitleBarBackColor", m_TitleBarBackColor, &HFFC0C0)
    Call .WriteProperty("TitleBarBackColorDesactivate", m_TitleBarBackColorDesactivate, &HE0E0E0)
End With
End Sub

Private Sub UserControl_Resize()
picLogo.Move 8, 0
UserControl.Height = picLogo.Height + 30
UserControl.Width = picLogo.Width + 40
End Sub

' Detiene el SubClass y elimina el HDC del StdPicture.
Private Sub UserControl_Terminate()
    UnHookForm 'lHwnd
    Debug.Print "TERMINATE"
End Sub


Private Function GetHitTestt(ByVal X As Long, ByVal Y As Long) As Long
    Dim Rec As RECT, lReturn As Long
    Dim PT As PointAPI
    
    PT.X = X: PT.Y = Y
    ScreenToClient FrmH, PT
    GetClientRect FrmH, Rec
    X = PT.X: Y = PT.Y

    If IsWinZoomed = False Then
        If X <= BorderWidth Then lReturn = HTLEFT
        If X >= Rec.Right - BorderWidth Then lReturn = HTRIGHT
        If Y <= BorderWidth Then lReturn = HTTOP
        If Y >= Rec.Bottom - BorderWidth Then lReturn = HTBOTTOM
        
        If (X <= BorderWidth) And (Y <= BorderWidth) Then lReturn = HTTOPLEFT
        If (X <= BorderWidth) And (Y >= Rec.Bottom - BorderWidth) Then lReturn = HTBOTTOMLEFT
        If (X >= Rec.Right - BorderWidth) And (Y <= BorderWidth) Then lReturn = HTTOPRIGHT
        If (X >= Rec.Right - BorderWidth) And (Y >= Rec.Bottom - BorderWidth) Then lReturn = HTBOTTOMRIGHT
    End If
    
    If m_ShowTitlebar Then
        If (X >= BorderWidth) And (X < Rec.Right - BorderWidth - mCBW) And (Y > BorderWidth) And (Y <= mCBH) Then lReturn = HTCAPTION
    End If
    
    If (Y > 1) And (Y <= mCBH) And (X < Rec.Right) And (X >= Rec.Right - mBtnW) Then lReturn = HTCLOSE
    If (Y > 1) And (Y <= mCBH) And (X < Rec.Right - mBtnW) And (X >= Rec.Right - mBtnW * 2) Then lReturn = HTMAXBUTTON
    If (Y > 1) And (Y <= mCBH) And (X < Rec.Right - (mBtnW * 2)) And (X >= Rec.Right - mBtnW * 3) Then lReturn = HTMINBUTTON
    
    If m_ShowIcon Then
        Dim IconSize As Long

        IconSize = 16 * nScale
        If (Y > (mCBH / 2 - IconSize)) And (Y < (mCBH / 2 + IconSize)) And (X > (4 * nScale)) And (X < (28 * nScale)) Then lReturn = HTSYSMENU
    End If
    
    If m_ShowTitlebar Then
        If lReturn = 0 And (X > BorderWidth) And (Y > BorderWidth) And (X < Rec.Right - BorderWidth) And (Y < Rec.Bottom - BorderWidth) Then
            lReturn = HTCLIENT
        End If
    Else
        If lReturn = 0 Then lReturn = HTCLIENT
    End If
    
    GetHitTestt = lReturn

End Function

Private Sub InitDIB(ByVal Width As Long, ByVal Height As Long, Optional copyB As Boolean)

    Dim DIBInf  As BITMAPINFO
    Dim hDIB As Long
    
    If DibDC <> 0 Then DeleteDIB

    DibDC = CreateCompatibleDC(0)
    
    With DIBInf.bmiHeader
        .biSize = Len(DIBInf.bmiHeader)
        .biWidth = Width
        .biHeight = Height
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
        .biSizeImage = AlignScan(.biWidth, .biBitCount) * .biHeight
        .biXPelsPerMeter = (GetDeviceCaps(DibDC, HORZRES) / GetDeviceCaps(DibDC, HORZSIZE)) * 1000
        .biYPelsPerMeter = (GetDeviceCaps(DibDC, VERTRES) / GetDeviceCaps(DibDC, VERTSIZE)) * 1000
    End With
   
    hDIB = CreateDIBSection(DibDC, DIBInf, DIB_RGB_COLORS, 0, 0, 0)
    hOldBMP = SelectObject(DibDC, hDIB)

End Sub

Private Sub DeleteDIB()
    DeleteObject SelectObject(DibDC, hOldBMP)
    DeleteDC DibDC
    DibDC = 0: hOldBMP = 0
End Sub

Private Function AlignScan(ByVal inWidth As Long, ByVal inDepth As Integer) As Long
    AlignScan = (((inWidth * inDepth) + &H1F) And Not &H1F&) \ &H8
End Function

' funcion para convertir un color long a un BGRA(Blue, Green, Red, Alpha)
Private Function ConvertColor(Color As Long, ByVal Opacity As Long) As Long
    Dim BGRA(0 To 3) As Byte
  
    BGRA(3) = CByte((Abs(Opacity) / 100) * 255)
    BGRA(0) = ((Color \ &H10000) And &HFF)
    BGRA(1) = ((Color \ &H100) And &HFF)
    BGRA(2) = (Color And &HFF)
    CopyMemory ConvertColor, BGRA(0), 4&
End Function
  

Private Sub DrawCloseButton(ByVal hGraphics As Long, ByVal hpen As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
    Dim P5 As Long

    P5 = 5 * nScale
    GdipDrawLineI hGraphics, hpen, Left + (Width / 2) - P5, Top + (Height / 2) - P5, Left + (Width / 2) + P5, Top + (Height / 2) + P5
    GdipDrawLineI hGraphics, hpen, Left + (Width / 2) - P5, Top + (Height / 2) + P5, Left + (Width / 2) + P5, Top + (Height / 2) - P5
End Sub

Private Sub DrawMaximizeButton(ByVal hGraphics As Long, ByVal hpen As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
    Dim P5 As Byte
    Dim P8 As Byte
    
    P5 = 5 * nScale
    P8 = 8 * nScale
    GdipDrawRectangleI hGraphics, hpen, Left + (Width / 2) - P5, Top + (Height / 2) - P5, P8, P8
End Sub

Private Sub DrawRestoreButton(ByVal hGraphics As Long, ByVal hpen As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
    Dim P2 As Byte, P3 As Byte, P4 As Byte, P5 As Byte, P6 As Byte, P8 As Byte
    
    P2 = 2 * nScale: P3 = 3 * nScale
    P4 = 4 * nScale: P5 = 5 * nScale
    P6 = 6 * nScale: P8 = 8 * nScale
   
    GdipDrawRectangleI hGraphics, hpen, Left + (Width / 2) - P5, Top + (Height / 2) - P4, P8, P8
    GdipDrawLineI hGraphics, hpen, Left + (Width / 2) - P3, Top + (Height / 2) - P4, Left + (Width / 2) - P3, Top + (Height / 2) - P6
    GdipDrawLineI hGraphics, hpen, Left + (Width / 2) - P2, Top + (Height / 2) - P6, Left + (Width / 2) + P5, Top + (Height / 2) - P6
    GdipDrawLineI hGraphics, hpen, Left + (Width / 2) + P3, Top + (Height / 2) + P2, Left + (Width / 2) + P5, Top + (Height / 2) + P2
    GdipDrawLineI hGraphics, hpen, Left + (Width / 2) + P5, Top + (Height / 2) - P6, Left + (Width / 2) + P5, Top + (Height / 2) + P2
End Sub

Private Sub DrawMinimizeButton(ByVal hGraphics As Long, ByVal hpen As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
    Dim P5 As Byte

    P5 = 5 * nScale
   GdipDrawLineI hGraphics, hpen, Left + (Width / 2) - P5, Top + (Height / 2), Left + (Width / 2) + P5, Top + (Height / 2)
End Sub

Private Function PaintBuffer()
     Dim hGraphics As Long
     Dim BtnH As Long
     Dim BtnW As Long
     Dim hdc As Long
     Dim hBrush As Long, hpen As Long
     Dim Rec As RECT
     Dim BtnState As Integer
     Dim isBtnPress As Boolean
     Dim lColor As Long
     Dim LightIcon As Boolean

     If m_ShowTitlebar And Not m_Focus Then
        LightIcon = Not IsLightColor(m_TitleBarBackColorDesactivate)
     Else
        LightIcon = m_WhiteIcons
     End If

     If m_UseSystemTheme And Not m_Focus Then
        lColor = vbBlack
     Else
        lColor = IIf(LightIcon, vbWhite, vbBlack)
     End If
     
     If GetKeyState(vbLeftButton) < 0 Then isBtnPress = True
     
     InitDIB mCBW, mCBH
     
     If GdipCreateFromHDC(DibDC, hGraphics) = 0 Then
         
         BtnState = IIf(m_Hittest = HTCLOSE, IIf(isBtnPress, 40, 90), 0)
         Call GdipCreateSolidFill(ConvertColor(vbRed, BtnState), hBrush)
         GdipFillRectangleI hGraphics, hBrush, mCBW - mBtnW, 0, mBtnW, mCBH
         GdipDeleteBrush hBrush
         
         BtnState = IIf(m_Hittest = HTMAXBUTTON, IIf(isBtnPress, 50, 25), 0)
         If Not IsStyle(WS_MAXIMIZEBOX) Then BtnState = 0
         Call GdipCreateSolidFill(ConvertColor(lColor, BtnState), hBrush)
         GdipFillRectangleI hGraphics, hBrush, mCBW - mBtnW * 2, 0, mBtnW, mCBH
         GdipDeleteBrush hBrush
         
         BtnState = IIf(m_Hittest = HTMINBUTTON, IIf(isBtnPress, 50, 25), 0)
         If Not IsStyle(WS_MINIMIZEBOX) Then BtnState = 0
         Call GdipCreateSolidFill(ConvertColor(lColor, BtnState), hBrush)
         GdipFillRectangleI hGraphics, hBrush, mCBW - mBtnW * 3, 0, mBtnW, mCBH
         GdipDeleteBrush hBrush
         
         If LightIcon Then
            lColor = vbWhite
         Else
            If m_Hittest = HTCLOSE And Not isBtnPress Then lColor = vbWhite Else lColor = vbBlack
         End If
         
         If m_UseSystemTheme And Not m_Focus Then
            If m_Hittest = HTCLOSE And Not isBtnPress Then lColor = vbWhite Else lColor = &H989487
         End If
         
         GdipCreatePen1 ConvertColor(lColor, 100), 1 * nScale, UnitPixel, hpen
         DrawCloseButton hGraphics, hpen, mCBW - mBtnW, 0, mBtnW, mCBH
         GdipDeletePen hpen
         '------------------------------------------------------------------------
         If LightIcon Then
            lColor = vbWhite
         Else
            lColor = IIf(m_Hittest = HTMAXBUTTON, vbWhite, vbBlack)
         End If
         
         If Not IsStyle(WS_MAXIMIZEBOX) Then lColor = &H989487
         If m_UseSystemTheme And Not m_Focus Then
            lColor = &H989487
         End If
         
         GdipCreatePen1 ConvertColor(lColor, 100), 1 * nScale, UnitPixel, hpen
         
         If IsWinZoomed Then
            DrawRestoreButton hGraphics, hpen, mCBW - mBtnW * 2, 0, mBtnW, mCBH
         Else
            DrawMaximizeButton hGraphics, hpen, mCBW - mBtnW * 2, 0, mBtnW, mCBH
         End If
         GdipDeletePen hpen
         '------------------------------------------------------------------------
         If LightIcon Then
            lColor = vbWhite
         Else
            lColor = IIf(m_Hittest = HTMINBUTTON, vbWhite, vbBlack)
         End If
         
         If Not IsStyle(WS_MINIMIZEBOX) Then lColor = &H989487
         
         If m_UseSystemTheme And Not m_Focus Then
            lColor = &H989487
         End If
         
         GdipCreatePen1 ConvertColor(lColor, 100), 1 * nScale, UnitPixel, hpen
         DrawMinimizeButton hGraphics, hpen, mCBW - mBtnW * 3, 0, mBtnW, mCBH
         GdipDeletePen hpen

         Call GdipDeleteGraphics(hGraphics)
     End If

End Function

Public Sub DrawCaption(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Optional hdc As Long)
        Dim OldColor As Long, OldFont As Long
        Dim Rec As RECT
        Dim LightIcon As Boolean
        Dim bCreate As Boolean
        
        If m_ShowTitlebar And Not m_Focus Then
           LightIcon = Not IsLightColor(m_TitleBarBackColorDesactivate)
        Else
           LightIcon = m_WhiteIcons
        End If
        
        If 0 Then
            hdc = GetDC(FrmH)
            bCreate = True
        End If
        
        SetRect Rec, X, Y, X + Width, Y + Height

        OldColor = GetTextColor(hdc)
        OldFont = SelectObject(hdc, hFont)
        
        If m_UseSystemTheme And Not m_Focus Then
            SetTextColor hdc, &H989487
        Else
            SetTextColor hdc, IIf(LightIcon, vbWhite, vbBlack)
        End If
        

        DrawText hdc, m_Caption, -1, Rec, DT_VCENTER Or DT_SINGLELINE Or DT_WORD_ELLIPSIS
        
        SetTextColor OldColor, OldColor
        Call SelectObject(hdc, OldFont)
        
        If bCreate Then Call ReleaseDC(FrmH, hdc)
End Sub


Private Function DWM_AccentColor() As Long
    Const AccentColorPath = "Software\Microsoft\Windows\DWM"
    Dim oColor As Long
    If SHGetValue(HKEY_CURRENT_USER, AccentColorPath, "AccentColor", 0, oColor, 4) = 0 Then
        DWM_AccentColor = oColor
    Else
        OleTranslateColor vbActiveTitleBar, 0, DWM_AccentColor
    End If
End Function



Private Sub ConfigureNCandClient(lParam As Long)
    Dim tNCR                As NCCALCSIZE_PARAMS
    Dim tWP                 As WINDOWPOS
    
    CopyMemory tNCR, ByVal lParam, Len(tNCR)
    CopyMemory tWP, ByVal tNCR.lppos, Len(tWP)
    
    Dim PT As PointAPI
    Dim RECT As RECT
    
    If diff = 0 Then
        ClientToScreen FrmH, PT
        GetWindowRect FrmH, RECT
        
        diff = (RECT.Top - PT.Y) + (1 * nScale)
      'Form1.Label1.Caption = diff
        If diff < -18 Then diff = 0
    End If
    
    With tNCR.rgrc(0)
        '.Left = tWP.X
        If Not IsWinZoomed Then
        '***************************
        '*****************************
            .Top = tWP.Y + diff
            If nScale > 1 Then .Top = .Top
        Else
            .Top = tWP.Y - 10
        End If
     End With
 
    LSet tNCR.rgrc(1) = tNCR.rgrc(0)
    CopyMemory ByVal lParam, tNCR, Len(tNCR)
    
End Sub

'Private Function GetImmersiveColor(ByVal sName As String) As Long
'    Dim lColorSet       As Long
'    lColorSet = GetImmersiveUserColorSetPreference(0, 0)
'
'    Dim lColorType      As Long
'    lColorType = GetImmersiveColorTypeFromName(StrPtr(sName))
'
'    Dim lRawColor       As Long
'    GetImmersiveColor = GetImmersiveColorFromColorSetEx(lColorSet, lColorType, 0, 0)
'End Function

Public Function IsLightColor(Color As Long) As Boolean
    Dim R As Integer, G As Integer, B As Integer
    R = &HFF& And Color
    G = (&HFF00& And Color) \ 256
    B = (&HFF0000 And Color) \ 65536
    If (R + G + B) / 3 > 127 Then IsLightColor = True
End Function

Public Function GetWindowsDPI() As Double
    Dim hdc As Long, LPX  As Double
    hdc = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hdc, LOGPIXELSX))
    ReleaseDC 0, hdc

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function


Public Function GetDpiMonitor() As Double
    Dim hMonitor As Long
    Dim X As Long, Y As Long
    
    hMonitor = MonitorFromWindow(FrmH, &H0)
    GetDpiForMonitor hMonitor, &H0, X, Y
    GetDpiMonitor = X / 96#
    
End Function


' Ordinal 3
Private Sub WndProc2(ByVal bBefore As Boolean, _
       ByRef bHandled As Boolean, _
       ByRef lReturn As Long, _
       ByVal hwnd As Long, _
       ByVal uMsg As Long, _
       ByVal wParam As Long, _
       ByVal lParam As Long, _
       ByRef lParamUser As Long)
       

End Sub

' Ordinal 2
Private Function TimerProc( _
       ByVal hwnd As Long, _
       ByVal tMsg As Long, _
       ByVal TimerID As Long, _
       ByVal tickCount As Long) As Long
    
    Dim PT As PointAPI
    GetCursorPos PT
    
    If WindowFromPoint(PT.X, PT.Y) <> hwnd Then
        TimerKill
        m_Hittest = 0
        PaintBuffer
        SendMessage hwnd, WM_PAINT, 0, 0
    End If

End Function


' Ordinal 1
Private Sub WndProc(ByVal bBefore As Boolean, _
       ByRef bHandled As Boolean, _
       ByRef lReturn As Long, _
       ByVal hwnd As Long, _
       ByVal uMsg As Long, _
       ByVal wParam As Long, _
       ByVal lParam As Long, _
       ByRef lParamUser As Long)
       
Dim Rec     As RECT
Dim X       As Long
Dim Y       As Long
Dim lRet    As Long


Select Case uMsg

    Case WM_DISPLAYCHANGE, WM_DPICHANGED

        nScale = GetWindowsDPI
        mCBH = 31 * nScale
        mCBW = 138 * nScale
        mBtnW = mCBW / 3
        BorderWidth = BorderPixels * nScale
        BorderHeight = BorderPixels * nScale
    
        If m_ShowIcon Then
            If m_Icon Then DestroyIcon m_Icon
            m_Icon = GetWindowIcon(hwnd, 16 * nScale)
        End If
        Call Refresh
    
    Case WM_NCHITTEST
        GetClientRect hwnd, Rec
        X = Get_X_lParam(lParam)
        Y = Get_Y_lParam(lParam)

        lReturn = GetHitTestt(X, Y)
   
        If m_Hittest <> lReturn Then
            If lReturn = HTCLOSE Or lReturn = HTMAXBUTTON Or lReturn = HTMINBUTTON Then
                m_Hittest = lReturn
                PaintBuffer
                SendMessage hwnd, WM_PAINT, 0, 0
                TimerSet 10
                
            Else
                 If m_Hittest = HTCLOSE Or m_Hittest = HTMAXBUTTON Or m_Hittest = HTMINBUTTON Then
                    m_Hittest = 0
                    PaintBuffer
                    SendMessage hwnd, WM_PAINT, 0, 0
                    TimerKill
                 End If
            End If
        End If
  
    Case WM_SIZE
        
        If m_Hittest = HTCLOSE Or m_Hittest = HTMAXBUTTON Or m_Hittest = HTMINBUTTON Or IsWinZoomed Then
            m_Hittest = 0
            PaintBuffer
        End If
        GetClientRect hwnd, Rec
        Rec.Bottom = mCBH
        InvalidateRect hwnd, Rec, 0

    Case WM_NCLBUTTONDOWN
        If m_Hittest = HTCLOSE Or m_Hittest = HTMAXBUTTON Or m_Hittest = HTMINBUTTON Then
            PaintBuffer
            SendMessage hwnd, WM_PAINT, 0, 0
            bHandled = True
        End If
   
        If GetHitTestt(Get_X_lParam(lParam), Get_Y_lParam(lParam)) = HTSYSMENU Then
            Static e As Long

            If GetTickCount - e = 0 Then Exit Sub

            GetWindowRect hwnd, Rec
            lRet = TrackPopupMenuEx(GetSystemMenu(hwnd, False), &H100&, Rec.Left + 3, Rec.Top + mCBH, hwnd, ByVal 0&)
            If lRet Then PostMessage hwnd, WM_SYSCOMMAND, lRet, ByVal 0&
            
            e = GetTickCount
            
        End If
        
    Case WM_NCRBUTTONUP
        m_Hittest = GetHitTestt(Get_X_lParam(lParam), Get_Y_lParam(lParam))
        If m_Hittest = HTSYSMENU Or m_Hittest = HTCAPTION Then
            GetWindowRect hwnd, Rec
            lRet = TrackPopupMenuEx(GetSystemMenu(hwnd, False), TPM_RETURNCMD, Get_X_lParam(lParam), Get_Y_lParam(lParam), hwnd, ByVal 0&)
            PostMessage hwnd, WM_SYSCOMMAND, lRet, ByVal 0&
        End If
        
    Case WM_NCLBUTTONUP
        If m_Hittest = HTCLOSE Or m_Hittest = HTMAXBUTTON Or m_Hittest = HTMINBUTTON Then
            PaintBuffer
            bHandled = True
            
            If m_Hittest = HTMINBUTTON Then If IsStyle(WS_MINIMIZEBOX) Then SendMessage hwnd, WM_SYSCOMMAND, IIf(IsIconic(hwnd), SC_RESTORE, SC_MINIMIZE), ByVal 0&
            If m_Hittest = HTMAXBUTTON Then If IsStyle(WS_MAXIMIZEBOX) Then SendMessage hwnd, WM_SYSCOMMAND, IIf(IsWinZoomed, SC_RESTORE, SC_MAXIMIZE), ByVal 0&
            If m_Hittest = HTCLOSE Then SendMessage hwnd, WM_SYSCOMMAND, SC_CLOSE, ByVal 0&
        End If
        
    Case WM_PAINT
        Dim hdc As Long
        Dim OldFont As Long
        Dim hBrush As Long
        Dim OldColor  As Long
        Dim IconSize As Long
        
        GetClientRect hwnd, Rec
        
        If bBefore Then
            Rec.Bottom = mCBH
            InvalidateRect hwnd, Rec, 1&
            Exit Sub
        End If
  
        hdc = GetDC(hwnd)
        
        If m_ShowTitlebar Then
            Rec.Bottom = mCBH
            If m_Focus Then
                hBrush = CreateSolidBrush(m_TitleBarBackColor) 'GetBkColor(hdc)
            Else
                hBrush = CreateSolidBrush(m_TitleBarBackColorDesactivate)
            End If
            Call FillRect(hdc, Rec, hBrush)
            DeleteObject hBrush
        End If
        
        GdiAlphaBlend& hdc, Rec.Right - mCBW, 0&, mCBW, mCBH, DibDC, 0, 0, mCBW, mCBH, 2 ^ 24 + &HFF0000 * 1

        Rec.Left = (10 * nScale)
        
        If m_ShowIcon Then
            IconSize = 16 * nScale
            DrawIconEx hdc, Rec.Left, (mCBH / 2) - (IconSize / 2), m_Icon, IconSize, IconSize, 0, 0, &H1 Or &H2
            Rec.Left = (10 + 20) * nScale
        End If
        
        If m_ShowCaption Then
            DrawCaption Rec.Left, Rec.Top, Rec.Right - mCBW - Rec.Left, mCBH, hdc
        End If
        
        ReleaseDC hwnd, hdc
        
    Case WM_SETICON
        
        If m_Icon Then DestroyIcon m_Icon: m_Icon = 0
        If ShowIcon Then
            m_Icon = GetWindowIcon(FrmH, 16 * nScale)
        End If
        Call SendMessage(hwnd, WM_PAINT, 0, 0)
        
    Case WM_NCCALCSIZE
        IsWinZoomed = IsZoomed(hwnd)
        ConfigureNCandClient lParam

    Case WM_NCLBUTTONDBLCLK

        If m_ShowIcon And GetHitTestt(Get_X_lParam(lParam), Get_Y_lParam(lParam)) = HTSYSMENU Then
            SendMessage hwnd, WM_SYSCOMMAND, SC_CLOSE, ByVal 0&
        Else
            If IsStyle(WS_MAXIMIZEBOX) Then SendMessage hwnd, WM_SYSCOMMAND, IIf(IsWinZoomed, SC_RESTORE, SC_MAXIMIZE), ByVal 0&
            PaintBuffer
        End If
        bHandled = True
        
    Case WM_GETMINMAXINFO

        Dim tMINMAXINFO As MINMAXINFO
        Dim uMonInfo As MONITORINFO
        Dim hMonitor As Long
        
        Call CopyMemory(tMINMAXINFO, ByVal lParam, LenB(tMINMAXINFO))
            
            tMINMAXINFO.ptMinTrackSize = m_MinFrmSize
            tMINMAXINFO.ptMaxTrackSize = m_MaxFrmSize

            hMonitor = MonitorFromWindow(hwnd, &H0)
            If hMonitor <> 0 Then
                uMonInfo.cbSize = Len(uMonInfo)
                Call GetMonitorInfo(hMonitor, uMonInfo)
            Else
                Call SystemParametersInfo(SPI_GETWORKAREA, 0, uMonInfo.rcWork, 0)
            End If

            AdjustWindowRectEx uMonInfo.rcWork, GetWindowLong(hwnd, GWL_STYLE), 0, GetWindowLong(hwnd, GWL_EXSTYLE)
            
            If tMINMAXINFO.ptMinTrackSize.X = 0 Then
                tMINMAXINFO.ptMinTrackSize.X = mCBW ' + (20 * nScale)
            End If
            
            If tMINMAXINFO.ptMinTrackSize.Y = 0 Then
                tMINMAXINFO.ptMinTrackSize.Y = mCBH + BorderHeight
            End If
            
            If tMINMAXINFO.ptMaxTrackSize.X = 0 Then
                tMINMAXINFO.ptMaxTrackSize.X = uMonInfo.rcWork.Right - uMonInfo.rcWork.Left '+ (12 * nScale)
            End If
            
            If tMINMAXINFO.ptMaxTrackSize.Y = 0 Then
                tMINMAXINFO.ptMaxTrackSize.Y = uMonInfo.rcWork.Bottom - uMonInfo.rcWork.Top '+ (8 * nScale)
                'If nScale > 1 Then tMINMAXINFO.ptMaxTrackSize.Y = tMINMAXINFO.ptMaxTrackSize.Y '+ (4 * nScale)
            End If

        Call CopyMemory(ByVal lParam, tMINMAXINFO, LenB(tMINMAXINFO))
        lReturn = DefWindowProc(hwnd, uMsg, wParam, lParam)
        
        bHandled = False
        
    Case WM_STYLECHANGED
        Dim lStyle As Long
        
        If wParam = GWL_STYLE Then
            lStyle = WS_BORDER Or WS_VISIBLE Or WS_THICKFRAME
            If IsStyle(WS_MINIMIZEBOX) Then lStyle = lStyle Or WS_MINIMIZEBOX
            If IsStyle(WS_MAXIMIZEBOX) Then lStyle = lStyle Or WS_MAXIMIZEBOX
            SetWindowLongA hwnd, GWL_STYLE, lStyle
        End If
        m_Caption = GetWindowCaption(hwnd)
        Call SendMessage(hwnd, WM_PAINT, 0, 0)
                
    Case WM_NCACTIVATE, WM_ACTIVATE, WM_ACTIVATEAPP
        m_Focus = wParam <> 0
        PaintBuffer
        Call SendMessage(hwnd, WM_PAINT, 0, 0)
        SetWindowPos FrmH, 0, 0, 0, 0, 0, 551
        
    Case WM_WININICHANGE
        If m_UseSystemTheme Then
            m_TitleBarBackColor = DWM_AccentColor And &HFFFFFF
            m_WhiteIcons = Not IsLightColor(m_TitleBarBackColor)
            m_TitleBarBackColorDesactivate = vbWhite
            Refresh
        End If

    Case WM_SIZING
        If wParam = 9 Then 'WMSZ_CAPTION
            PaintBuffer

            Call SendMessage(hwnd, WM_PAINT, 0, 0)
        End If
        
    Case WM_ENTERSIZEMOVE
        RaiseEvent ENTERSIZEMOVE
    Case WM_EXITSIZEMOVE
        RaiseEvent EXITSIZEMOVE
    'Case WM_NCMOUSELEAVE 'El timer es mas efectivo
    '    PaintBuffer
    '    Call SendMessage(hwnd, WM_PAINT, 0, 0)
    'Case Else
    '    Debug.Print uMsg, Hex(uMsg)
    'Case WM_NCPAINT
    'Dim hpen As Long
    'hBrush = CreateSolidBrush(vbBlue)
    'hpen = CreatePen(0, 1, vbRed)
    'Dim hdc As Long
    'hdc = GetDCEx(hwnd, wParam, DCX_WINDOW Or DCX_INTERSECTRGN)
    'hdc = GetWindowDC(hwnd)
    'SelectObject hdc, hpen
    
    'Dim RECT As RECT
    

        'GetWindowRect hwnd, RECT
    'RECT.Right = 10000
    'RECT.Bottom = 1000
    
    'FillRect hdc, RECT, hBrush
    'Rectangle hdc, 4, 4, 100, 100
    
    'Call ReleaseDC(hwnd, hdc)
        
End Select
    
End Sub


