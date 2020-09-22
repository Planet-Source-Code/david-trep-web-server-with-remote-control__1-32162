Attribute VB_Name = "mdlAPI"
Option Explicit

Public retval As Long

Type POINTAPI
        x As Long
        y As Long
End Type

'  Pen Styles
Const PS_SOLID = 0
Const PS_DASH = 1                    '  -------
Const PS_DOT = 2                     '  .......
Const PS_DASHDOT = 3                 '  _._._._
Const PS_DASHDOTDOT = 4              '  _.._.._
Const PS_NULL = 5
Const PS_INSIDEFRAME = 6
Const PS_USERSTYLE = 7
Const PS_ALTERNATE = 8
Const PS_STYLE_MASK = &HF
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4

Declare Function CreateDC Lib "GDI32" Alias "CreateDCA" (ByVal _
    lpszDriverName As String, ByVal lpszDeviceName As String, _
    ByVal lpszOutput As String, lpInitData As DEVMODE) As Long
   
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Declare Function MoveToEx Lib "GDI32" (ByVal hDC As _
    Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
           
Declare Function LineTo Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, _
    ByVal y As Long) As Long
    
Declare Function CreatePen Lib "GDI32" (ByVal _
    nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) _
    As Long
    
Declare Function SelectObject Lib "GDI32" Alias "Selectobject" (ByVal hDC As _
    Long, ByVal hObject As Long) As Long
    
Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As _
    Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal _
    nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal _
    YSrc As Long, ByVal dwRop As Long) As Long
    
Declare Function Ellipse Lib "GDI32" (ByVal hDC _
    As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
    ByVal Y2 As Long) As Long
    
Declare Function Rectangle Lib "GDI32" (ByVal _
    hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
    ByVal Y2 As Long) As Long
    
Declare Function SetROP2 Lib "GDI32" Alias "SetRop2" (ByVal hDC As _
    Long, ByVal nDrawMode As Long) As Long
    
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
    
Declare Function SystemParametersInfo Lib "user32" _
    Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal _
    uParam As Long, ByVal lpvParam As Long, ByVal fuWinIni As _
    Long) As Long
    
Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    
Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long)
        
'  size of a device name string
Const CCHDEVICENAME = 32

'  size of a form name string
Const CCHFORMNAME = 32

Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

