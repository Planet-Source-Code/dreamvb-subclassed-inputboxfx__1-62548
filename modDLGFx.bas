Attribute VB_Name = "InputFx"
'DM InputboxFX API Subclassed Version
'This allows you to change many things with an inputbox and is all done
'by subclassing the classic VB inputbox.
' all in all I think this project took around 2 hours to make so it not that bad

'Features version 1.0
' Support to use your own Fonts includeing Styles like Bold,Italic well you know
' Support to change backcolor of the dialog, Forecolor of the promt text
' Support to change Forecolor and Backcolor of the EditBox
' Support to Disable Right Clicking of the EditBox
' Support to Disbale the X Button Close button
' Support to add a Timed out dilaog so it self closes
' Support to add alphablending. Note this will only work on Win2k and Better OS
' Support to find out what button was passed
' Support to add textures

' see below Enum Buttons

' anyway apart from all that hope you like it
' all code is commented I also suguest that beginners take a look at this to
' as you should pick up some help on subclassing and windows related stuff.

'as always use the code as you see fit. just remmber were it came from.

'Thanks
'Ben Jones
'Questions and Answers on any of my project or need a quick vb tips
'http://www.eraystudios.com/forum


Option Explicit

Public Enum Buttons
    [fxSystem] = 0
    [fxOk] = 1
    [fxCancel] = 2
End Enum

Private Type mDialogPos
    XPos As Long
    YPos As Long
End Type

Private Type CWPSTRUCT
    lParam As Long
    wParam As Long
    message As Long
    hwnd As Long
End Type

Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateBrushIndirect Lib "gdi32.dll" (ByRef lpLogBrush As LOGBRUSH) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetBkColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateFont Lib "gdi32.dll" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetSystemMenu Lib "user32.dll" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nBkMode As Long) As Long

Private Declare Function SetTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
Private Declare Function LoadBitmap Lib "user32.dll" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As String) As Long


'Used for hooking
Private Const WH_CALLWNDPROC = 4
Private Const GWL_WNDPROC = (-4)

'Window Messages
Private Const WM_CTLCOLORDLG As Long = &H136
Private Const WM_CTLCOLORSTATIC As Long = &H138
Private Const WM_CTLCOLOREDIT As Long = &H133
Private Const WM_SHOWWINDOW As Long = &H18
Private Const WM_CTLCOLORBTN = &H135
Private Const WM_DESTROY = &H2
Private Const WM_CREATE = &H1
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_CLOSE As Long = &H10
Private Const WM_COMMAND As Long = &H111

Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_CLIENTEDGE As Long = &H200&
Private Const WS_EX_STATICEDGE As Long = &H20000
Private Const WS_EX_DLGMODALFRAME As Long = &H1&

Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2

Private Const MF_BYPOSITION As Long = &H400&
Private Const MF_REMOVE = &H1000&
'Hook vars
Private m_InputProc As Long, m_Hook As Long, m_EditBoxProc As Long, m_EditHwnd As Long
Private m_DialogHwnd As Long

Private m_DialogPos As mDialogPos
Private WinStruc As CWPSTRUCT

'Backcolor and forecolor and style vars
Private m_BackColor As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private m_EditBkColor As OLE_COLOR
Private m_EditForeColor As OLE_COLOR
Private m_FlatEditBox As Boolean
Private m_BackGroundImg As Variant
Private m_UseBackImage As Boolean

'any Other vars we may have used
Private m_AlphaLevel As Integer
Private m_RightClick As Boolean
Private m_DisableX As Boolean
Private m_ButtonPressed As Buttons

'Timer related vars
Private TimerID As Long
Public is_active As Boolean
Private TimeHwnd As Long
Private m_TimeOut As Long
Private m_Enable_TimeOut As Boolean
Private TimeVal As Long 'Counter

'Font vars
Private hObjFont As Long
Private m_Bold As Boolean, m_Italic As Boolean, m_Underline As Boolean, m_FontName As String, m_FontSize As Long

Function GetWinText(lpWindLng As Long) As String
Dim sBuff As String, iTxtLen As Long
    sBuff = Space(256)
    iTxtLen = GetWindowTextLength(lpWindLng) + 1
    
    Call GetWindowText(lpWindLng, sBuff, iTxtLen)
    GetWinText = Left(sBuff, iTxtLen - 1)
    
    sBuff = "": iTxtLen = 0
    
End Function

Private Sub TimerProc(hwnd As Long, msg As Long, idTimer As Long, dwTime As Long)
    'This small little Timer Proc allows the dialog to be closed after an amount of time
    TimeVal = TimeVal + 1 'Little counter
    
    If (TimeVal >= m_TimeOut) Then
        PostMessage m_DialogHwnd, WM_CLOSE, 0, 0 'Send a close message
        StopTimer
    End If
    
End Sub

Public Sub StopTimer()
    If (is_active = True) And (TimerID <> 0) Then
        'If timer is active and we have a timer ID kill the timer
        KillTimer TimeHwnd, TimerID
        is_active = False 'Not active any more
    End If
End Sub

Public Function CreateTimer(hwnd As Long)
Dim pause As Long
    'Create a timer for the window
    If is_active Then StopTimer 'if timer is already running stop it
    'Createa new timer
    TimerID = SetTimer(hwnd, 0, m_TimeOut, AddressOf TimerProc)
    
    If TimerID <> 0 Then
        is_active = True 'Noe inform us that it's active
    End If
    
End Function

Private Function TranslateColor(OleClr As OLE_COLOR, Optional hPal As Integer = 0) As Long
    ' used to return the correct color value of OleClr as a long
    If OleTranslateColor(OleClr, hPal, TranslateColor) Then
        TranslateColor = &HFFFF&
    End If
End Function

Private Function StrGetClassName(lpWindLng As Long)
Dim sBuff As String, iCls As Long
    'Used for finding a classname of a window from it's hwnd
    sBuff = Space(256) 'create a buffer to hold classname
    iCls = GetClassName(lpWindLng, sBuff, 256) ' Get classname
    
    If iCls <> 0 Then
        'Strip away any unnessary nullchars and retun the classname
        StrGetClassName = Left(sBuff, iCls)
        sBuff = ""
    End If
End Function

Function EditBoxProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'This sub classes the small edit box to disbale the right click menu
    Select Case msg
        Case WM_RBUTTONDOWN
            If Not m_RightClick Then
                EditBoxProc = 0
            Else
                EditBoxProc = CallWindowProc(m_EditBoxProc, hwnd, msg, wParam, lParam)
            End If
        Case Else
            EditBoxProc = CallWindowProc(m_EditBoxProc, hwnd, msg, wParam, lParam)
    End Select
End Function

Private Function HookWindow(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    'This is where you need to Hook the inputbox
    CopyMemory WinStruc, ByVal lParam, Len(WinStruc)

    If WinStruc.message = WM_CREATE Then
        'Locate the class name of the inputbox '#32770'
        If StrGetClassName(WinStruc.hwnd) = "#32770" Then
            m_DialogHwnd = WinStruc.hwnd
            'Start subclassing the inputbox window
            m_InputProc = SetWindowLong(WinStruc.hwnd, GWL_WNDPROC, AddressOf InputProc)
        End If
    End If
    
    HookWindow = CallNextHookEx(m_Hook, nCode, wParam, ByVal lParam) 'Keep sending the messages back to the system
End Function

Private Function InputProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim WndRect As RECT
Dim lb As LOGBRUSH
Dim iRet As Long
Dim hMenu As Long
Dim sButton As String

Dim Wnd_Width As Long, Wnd_Height As Long, txtHwnd As Long, bSize As Integer

    'This is the main part were we can play around with the inputboxs interface or just skip any messages
    
    Select Case WinStruc.message
        
        Case WM_COMMAND
            'We use the code below to find out what button was pressed
            sButton = UCase(GetWinText(lParam))
            
            If sButton = "OK" Then
                m_ButtonPressed = fxOk 'OK  pressed
            ElseIf sButton = "CANCEL" Then
                m_ButtonPressed = fxCancel 'cancel press
            Else
                m_ButtonPressed = fxSystem 'Closed by system command X button or self close timer
            End If
            
            InputProc = CallWindowProc(m_InputProc, hwnd, msg, wParam, lParam)
        Case WM_CLOSE 'Close the window
            TimeVal = 0
            InputProc = CallWindowProc(m_InputProc, hwnd, msg, wParam, lParam)
        Case WM_SHOWWINDOW 'Show window
            If EnableTimeOut Then CreateTimer m_DialogHwnd 'if use has request a time out box create then a timer
            
            If m_DisableX Then
                'Disable the X Close button on the dialog
                hMenu = GetSystemMenu(hwnd, False)
                RemoveMenu hMenu, GetMenuItemCount(hMenu) - 1, MF_BYPOSITION Or MF_REMOVE
            End If
            
            iRet = iRet Or WS_EX_LAYERED Or WS_EX_DLGMODALFRAME  'Style to set on the window
            SetWindowLong hwnd, GWL_EXSTYLE, iRet 'Set the window with the new style
            SetLayeredWindowAttributes hwnd, 0, m_AlphaLevel, LWA_ALPHA
            'Get the Windows Rect, we need to use this for positioning the window, and also find the height and width
            GetWindowRect hwnd, WndRect
            Wnd_Width = (WndRect.Right - WndRect.Left) 'Window width
            Wnd_Height = (WndRect.Bottom - WndRect.Top) 'Window Height
            
            'This allows you to flatten down the edit box
            m_EditHwnd = FindWindowEx(hwnd, 0, "Edit", vbNullString) 'Locate the edit box's Hwnd
            'subclass edit box
            m_EditBoxProc = SetWindowLong(m_EditHwnd, GWL_WNDPROC, AddressOf EditBoxProc)
            'Fallten the textbox border is user has requested to
            If m_EditHwnd <> 0 Then FlatBorder m_EditHwnd, m_FlatEditBox 'Apply the flat style to the Edit box
            MoveWindow hwnd, m_DialogPos.XPos, m_DialogPos.YPos, Wnd_Width, Wnd_Height, True 'Move and resize the window
            
        Case WM_CTLCOLORDLG, WM_CTLCOLORSTATIC 'Clor of dialog and static text
            SetBkMode wParam, 0
            SetBkColor wParam, m_BackColor ' Set the dialogs backcolor
            'Below creates the font object to be used on the inputbox
            
            If Len(Trim(m_FontName)) = 0 Then m_FontName = "Arial" 'If not fontname is found use default one
            If m_Bold Then bSize = 700 Else bSize = 0 'Set on Bold text
            If m_FontSize = 0 Then m_FontSize = 16 'Set default fontsize if none is incldued
            hObjFont = CreateFont(m_FontSize, 0, 0, 0, bSize, m_Italic, m_Underline, 0, 1, 0, 0, 0, 0, m_FontName)
            SelectObject wParam, hObjFont 'Place the new Font Object onto the DC
            
            SetTextColor wParam, m_ForeColor 'Set the textcolor of the static text
            lb.lbColor = m_BackColor
            InputProc = CreateBrushIndirect(lb) 'Send back the Brush information to be used on the dialog
            
            'Tile check if texture is enabled and we have a vaild bitmap
            If (m_UseBackImage) And Not IsEmpty(m_BackGroundImg) Then InputProc = CreatePatternBrush(m_BackGroundImg)
            
        Case WM_CTLCOLOREDIT
            'Set the backcolor and forecolor of the edit box
            SetBkColor wParam, m_EditBkColor
            SetTextColor wParam, m_EditForeColor
            lb.lbColor = m_EditBkColor
            InputProc = CreateBrushIndirect(lb)
        Case WM_DESTROY
            TimeVal = 0
            StopTimer
            If hObjFont <> 0 Then DeleteObject hObjFont 'Delete the font Object
            Call UnHookInputProc 'Destroy the subclassing to the inputbox's window
        Case Else
            
            'Keep sending the messages back
            InputProc = CallWindowProc(m_InputProc, hwnd, msg, wParam, lParam)
    End Select
    
End Function

Private Sub UnHookInputProc()
    'Used to remove a hook on a window and return it back to normal
    SetWindowLong WinStruc.hwnd, GWL_WNDPROC, m_InputProc
    SetWindowLong m_EditHwnd, GWL_WNDPROC, m_EditBoxProc
End Sub

Public Function InputBoxFx(Prompt As String, Optional Title, Optional Default, Optional XPos As Long, Optional YPos As Long, Optional HelpFile As String, _
Optional Context As Long, Optional BackColor As OLE_COLOR = vbButtonFace, Optional PromtColor As OLE_COLOR = vbBlack, _
Optional AlphaLevel As Integer = 255, Optional EditBoxFC As OLE_COLOR = vbBlack, Optional EditBoxBK As OLE_COLOR = vbWhite) As Variant

    m_Hook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf HookWindow, App.hInstance, App.ThreadID)
    'Above sets a hook to locate a message box as it;s been created. we can then sublcass it latter
    
    m_BackColor = TranslateColor(BackColor) 'Back color of inputbox
    m_ForeColor = TranslateColor(PromtColor) ' Text color of for inputbox
    
    m_EditBkColor = TranslateColor(EditBoxBK) 'Back color of edit box
    m_EditForeColor = TranslateColor(EditBoxFC) 'forecolor of edit box

    m_AlphaLevel = AlphaLevel ' Alpha Level to be used. note only works on Windows 2000 and above,
    
    'Below is were the inputbox will be displayed on the screen
    m_DialogPos.XPos = XPos
    m_DialogPos.YPos = YPos
    
    InputBoxFx = InputBox(Prompt, Title, Default, XPos, YPos, HelpFile, Context) 'Show the input box
    Call UnhookWindowsHookEx(m_Hook) ' Unhook the inputbox
End Function

Private Function FlatBorder(ByVal hwnd As Long, bFlat As Boolean)
Dim m_OldStyle As Long
    
    m_OldStyle = GetWindowLong(hwnd, GWL_EXSTYLE) 'Get the windows orginal style
    
    If bFlat Then
        'Turn on Flat style
        m_OldStyle = m_OldStyle And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    Else
        'Turn of flat style
        m_OldStyle = m_OldStyle And Not WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE
    End If
    
    SetWindowLong hwnd, GWL_EXSTYLE, m_OldStyle 'Set the window with the style
    'Update the window
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
  
End Function

'Property stuff down here
Public Property Get FlatEditBox() As Boolean
    FlatEditBox = m_FlatEditBox
End Property

Public Property Let FlatEditBox(ByVal vFlatEdit As Boolean)
    m_FlatEditBox = vFlatEdit
End Property

Public Property Get fFontBold() As Boolean
    fFontBold = m_Bold
End Property

Public Property Let fFontBold(ByVal vNewBold As Boolean)
    m_Bold = vNewBold
End Property

Public Property Get fItalic() As Boolean
    fItalic = m_Italic
End Property

Public Property Let fItalic(ByVal vNewItalic As Boolean)
    m_Italic = vNewItalic
End Property

Public Property Get fUnderLine() As Boolean
    fUnderLine = m_Underline
End Property

Public Property Let fUnderLine(ByVal vNewUnderLine As Boolean)
    m_Underline = vNewUnderLine
End Property

Public Property Get fFontName() As String
    fFontName = m_FontName
End Property

Public Property Let fFontName(ByVal vNewFont As String)
   m_FontName = vNewFont
End Property

Public Property Get fFontSize() As Long
    fFontSize = m_FontSize
End Property

Public Property Let fFontSize(ByVal vNewfSize As Long)
    m_FontSize = vNewfSize
End Property

Public Property Get AllowRClick() As Boolean
    AllowRClick = m_RightClick
End Property

Public Property Let AllowRClick(ByVal vNewBool As Boolean)
    m_RightClick = vNewBool
End Property


Public Property Get DisableX() As Boolean
    DisableX = m_DisableX
End Property

Public Property Let DisableX(ByVal vNewX As Boolean)
    m_DisableX = vNewX
End Property

Public Property Get EnableTimeOut() As Boolean
    EnableTimeOut = m_Enable_TimeOut
End Property

Public Property Let EnableTimeOut(ByVal vNewTimeOut As Boolean)
    m_Enable_TimeOut = vNewTimeOut
End Property

Public Property Get TimeOut() As Long
    TimeOut = m_TimeOut
End Property

Public Property Let TimeOut(ByVal vNewTimer As Long)
    m_TimeOut = vNewTimer
End Property

Public Property Get ButtonPressed() As Buttons
    ButtonPressed = m_ButtonPressed
End Property

Public Property Get AllowBackImage() As Boolean
    AllowBackImage = m_UseBackImage
End Property

Public Property Let AllowBackImage(ByVal vNewBkImage As Boolean)
    m_UseBackImage = vNewBkImage
End Property

Public Property Let Image(ByVal vNewValue As Variant)
    Set m_BackGroundImg = vNewValue
End Property
