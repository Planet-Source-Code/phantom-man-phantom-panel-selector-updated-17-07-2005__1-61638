VERSION 5.00
Begin VB.UserControl PanelControl 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   ScaleHeight     =   297
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   201
End
Attribute VB_Name = "PanelControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Name      : PanelControl
' DateTime  : 08 July 2005 10:49
' Author    : Gary Noble
' Assumes   : pTab, pTabs, cMemdc And ITabEventHandler
' Purpose   : A Side Panel Control That Simulates A Tab Control But With A Better GUI
' Revision  : V1.02
'---------------------------------------------------------------------------------------
' History   :
'             08 July 2005 - Inital Version (GN)
'             12 July 2005 - Added Mask Color For Icons/Pictures (GN)
'                          - ReStructured The Messenger Style Drawing Routine (GN)
'                          - Messenger Style Panel Rect Not Set Properly. (GN)
'                            When Hovering Over The Top, The Tooltip Was Displaying
'                            The First Panel Data Instead Of The Selected Panel
'                            When The Mouse Was Out Side The Panel Button CoOrdinates
'             15 July 2005 - Embedded Control Height Was Wrong When In Messenger Draw Style
'                            Height Was not To Exact Scale and The Bottom of the Control Was
'                            Not Being Drawn Correctly.
'                          - Bottom Right Hand Pixel Was Not Being Draw Properly When Right
'                            To Left Mode Was Activated.
'                          - Set Version To 1.02
'             17 July 2005 - Updated Messenger Back Paint Colour To Blt Transparent.
'                          - The Selected Tab Was Not Painting Properly When The Backcolor
'                            Was White. (Thanks To Riccardo Cohen)
'---------------------------------------------------------------------------------------
'
'   This Control Also Uses Code From Other Authors, All The Original Copyrights
'   Credits Can Be Found Where They Put Them.
'
'   Give Credit Where Credit Is Due.
'
'   Special Thanks To: Vlad Vissoultchev  - Memdc Drawing Class
'                      Paul Caton - Subclassing Code
'
'
'   You are free to use this source as long as this copyright message
'               appears on your program's "About" dialog:
'
'   Panel Selector
'   Copyright (c) 2005 Gary Noble (gwnoble@msn.com)
'
'
'---------------------------------------------------------------------------------------
' Notes      : Please keep Any Copyright Notices
'            : Adhere To Copyright Laws
'---------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
'---Copyright (c) 2004-2005 Gary Noble (gwnoble@msn.com)
'------------------------------------------------------------------------
'
'--- Redistribution and use in source and binary forms, with or
'--- without modification, are permitted provided that the following
'--- conditions are met:
'
'--- 1. Redistributions of source code must retain the above copyright
'---    notice, this list of conditions and the following disclaimer.
'
'--- 2. Redistributions in binary form must reproduce the above copyright
'---    notice, this list of conditions and the following disclaimer in
'---    the documentation and/or other materials provided with the distribution.
'
'--- 3. The end-user documentation included with the redistribution, if any,
'---    must include the following acknowledgment:
'
'---    "This product includes software developed by Gary Noble"
'
'--- Alternately, this acknowledgment may appear in the software itself, if
'--- and wherever such third-party acknowledgments normally appear.
'
'--- THIS SOFTWARE IS PROVIDED "AS IS" AND ANY EXPRESSED OR IMPLIED WARRANTIES,
'--- INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY
'--- AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL
'--- GARY NOBLE OR ANY CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT,
'--- INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
'--- BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF
'--- USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
'--- THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
'--- (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
'--- THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
'------------------------------------------------------------------------

Option Explicit

'-- ToolTip
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

'-- Windows API Functions
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExToolTipStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwToolTipStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'-- Windows API Constants
Private Const WM_USER = &H400
Private Const CW_USEDEFAULT = &H80000000
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1


'-- Mouse Movement
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'-- Tooltip Window Constants
Private Const TTS_NOPREFIX = &H2
Private Const TTF_TRANSPARENT = &H100
Private Const TTF_CENTERTIP = &H2
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETToolTipTitle = (WM_USER + 32)
Private Const TTS_BALLOON = &H40
Private Const TTS_ALWAYSTIP = &H1
Private Const TTF_SUBCLASS = &H10
Private Const TTF_IDISHWND = &H1
Private Const TTM_SETDELAYTIME = (WM_USER + 3)
Private Const TTDT_AUTOPOP = 2
Private Const TTDT_INITIAL = 3

Private Const TOOLTIPS_CLASSA = "tooltips_class32"

'-- Tooltip Window Types
Private Type TOOLINFO
    lSize As Long
    lFlags As Long
    hwnd As Long
    lId As Long
    lpRect As RECT
    hInstance As Long
    lpStr As String
    lParam As Long
End Type


Public Enum ttIconType
    TTNoIcon = 0
    TTIconInfo = 1
    TTIconWarning = 2
    TTIconError = 3
End Enum

Public Enum ttToolTipStyleEnum
    TTStandard
    TTBalloon
End Enum

Private mvarBackColor As Long
Private mvarToolTipTitle As String
Private mvarForeColor As Long
Private mvarIcon As ttIconType
Private mvarCentered As Boolean
Private mvarToolTipStyle As ttToolTipStyleEnum
Private mvarTipText As String
Private mvarVisibleTime As Long
Private mvarDelayTime As Long

Private m_lTTHwnd As Long
Private m_lParentHwnd As Long
Private ti As TOOLINFO

'-- Events
Public Event PanelSelected(ByRef oPanel As pTab)
Public Event PanelHovering(ByRef oPanel As pTab)
Public Event PanelMouseDown(ByVal Button As Single)
Public Event PanelMouseUp(ByVal Button As Single)
Public Event SystemColorChanged()
Public Event ThemeChanged(ByVal ThemeName As String)

'-- Button Down Flag
Private m_bButtonDown As Boolean

'-- Focus
Private m_bInFocus As Boolean

'-- Theme Calls
Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, ByVal dwMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32.dll" (ByRef lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32.dll" (ByVal hRgn As Long, ByRef lpRect As RECT) As Long
Private Declare Function SelectClipRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Const ALTERNATE = 1
Private Type POINTAPI
    x As Long
    y As Long
End Type


'-- OLE Calls
Private Const CLR_INVALID = -1


Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

'-- Panel Draw Type
Public Enum EPanel_DrawStyle
    ETDS_Classic = 0
    ETDS_ClassicX = 1
    ETDS_MSMessenger = 2
End Enum

'-- Panel Draw Type
Public Enum EPanel_Alignment
    EDPA_Left = 0
    EDPA_Right = 1
End Enum

Const m_def_PanelStyle = 0
Dim m_PanelStyle As EPanel_DrawStyle

'-- Picture/Panel Min Size
Public Enum EPanel_PicSize
    EPanelSize8x8_ = 8
    EPanelSize16x16_ = 16
    EPanelSize24x24_ = 24
    EPanelSize32x32_ = 32
    EPanelSize48x48_ = 48
End Enum

Private m_RectControlBounderies As RECT               '-- Embedded control Boundaries
Private m_sCurrentSystemThemename As String           '-- Current Theme Name
Private m_oSelectedItem As pTab                       '-- Selected Panel
Private m_oHoverItem As pTab                          '-- Hovering Panel
Private m_oPaintDC As cMemDC                          '-- Vlads Paint Dc
Private m_lFontHeight As Long                         '-- Font Height
Private m_lPanelOffset As Long                        '-- Panel Height


'-- Color selection
Private m_lColorOneSelectedNormal As OLE_COLOR
Private m_lColorTwoSelectedNormal As OLE_COLOR
Private m_lColorOneNormal As OLE_COLOR
Private m_lColorTwoNormal As OLE_COLOR
Private m_lColorOneSelected As OLE_COLOR
Private m_lColorTwoSelected As OLE_COLOR
Private m_lColorHeaderColorOne As OLE_COLOR
Private m_lColorHeaderColorTwo As OLE_COLOR
Private m_lColorHeaderForeColor As OLE_COLOR
Private m_lColorHotOne As OLE_COLOR
Private m_lColorHotTwo As OLE_COLOR
Private m_lColorBorder As OLE_COLOR

'-- Panels
Dim m_Panels As pTabs

'-- Panel Size
Const m_def_PanelIconPictureSize = 16
Dim m_PanelIconPictureSize As EPanel_PicSize

Const m_def_MinEmbebedControlHeight = 25
Dim m_MinEmbebedControlHeight As Long

Const m_def_PanelAlignment = 0
Dim m_PanelAlignment As EPanel_Alignment

'-- Window Update
Const m_def_LockUpdate = False
Dim mb_LockUpdate As Boolean

'-- Custom color Values/Calls
Const m_def_CustomColor = vbInactiveTitleBar
Const m_def_UseCustomColor = False
Private m_CustomColor As OLE_COLOR
Private m_UseCustomColor As Boolean


'-- Paul Caton's Subclassing source
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_SYSCOLORCHANGE As Long = &H15
Private Const WM_THEMECHANGED As Long = &H31A

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

Private bTrack As Boolean
Private bTrackUser32 As Boolean
Private bInCtrl As Boolean

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

'==================================================================================================
'Subclasser declarations

Private Enum eMsgWhen
    MSG_AFTER = 1                                     'Message calls back after the original (previous) WndProc
    MSG_BEFORE = 2                                    'Message calls back before the original (previous) WndProc
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE    'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES As Long = -1               'All messages added or deleted
Private Const GMEM_FIXED As Long = 0                  'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC As Long = -4                'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04 As Long = 88                   'Panelle B (before) address patch offset
Private Const PATCH_05 As Long = 93                   'Panelle B (before) entry count patch offset
Private Const PATCH_08 As Long = 132                  'Panelle A (after) address patch offset
Private Const PATCH_09 As Long = 137                  'Panelle A (after) entry count patch offset

Private Type tSubData                                 'Subclass data type
    hwnd As Long                                      'Handle of the window being subclassed
    nAddrSub As Long                                  'The address of our new WndProc (allocated memory).
    nAddrOrig As Long                                 'The address of the pre-existing WndProc
    nMsgCntA As Long                                  'Msg after Panelle entry count
    nMsgCntB As Long                                  'Msg before Panelle entry count
    aMsgTblA() As Long                                'Msg after Panelle array
    aMsgTblB() As Long                                'Msg Before Panelle array
End Type

Private sc_aSubData() As tSubData                     'Subclass data array

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'==================================================================================================

Const m_def_ShowTips = True
Const m_def_SelectedItemColor = 0

'Property Variables:
Dim m_ShowTips As Boolean

'-- Selected Font Color
Dim m_SelectedItemColor As OLE_COLOR

'======================================================================================================
'UserControl private routines

'Determine if the passed function is supported
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
    Dim hMod As Long
    Dim bLibLoaded As Boolean

    hMod = GetModuleHandleA(sModule)

    If hMod = 0 Then
        hMod = LoadLibraryA(sModule)
        If hMod Then
            bLibLoaded = True
        End If
    End If

    If hMod Then
        If GetProcAddress(hMod, sFunction) Then
            IsFunctionExported = True
        End If
    End If

    If bLibLoaded Then
        Call FreeLibrary(hMod)
    End If
End Function

'Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
    Dim tme As TRACKMOUSEEVENT_STRUCT

    If bTrack Then
        With tme
            .cbSize = Len(tme)
            .dwFlags = TME_LEAVE
            .hwndTrack = lng_hWnd
        End With

        If bTrackUser32 Then
            Call TrackMouseEvent(tme)
        Else
            Call TrackMouseEventComCtl(tme)
        End If
    End If
End Sub

'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
'Parameters:
'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
'hWnd     - The window handle
'uMsg     - The message number
'wParam   - Message related data
'lParam   - Message related data
'Notes:
'If you really know what you're doing, it's possible to change the values of the
'hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
'values get passed to the default handler.. and optionaly, the 'after' callback
    Static bMoving As Boolean

    Select Case uMsg
        Case WM_MOUSEMOVE
            If Not bInCtrl Then
                bInCtrl = True
                Call TrackMouseLeave(lng_hWnd)
                Redraw
            End If

        Case WM_MOUSELEAVE
            bInCtrl = False
            Set m_oHoverItem = Nothing
            Redraw
        Case WM_SYSCOLORCHANGE
            GetThemeName hwnd
            RaiseEvent SystemColorChanged
            pvGetGradientColors
            Redraw
        Case WM_THEMECHANGED
            GetThemeName hwnd
            RaiseEvent ThemeChanged(m_sCurrentSystemThemename)
            pvGetGradientColors
            Redraw

    End Select
End Sub

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines

'Add a message to the Panelle of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback Panelle
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
End Sub

'Delete a message from the Panelle of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback Panelle
'uMsg      - The message number that will be removed from the callback Panelle. NB Can also be ALL_MESSAGES, ie all messages will callback
'When      - Whether the msg is to be removed from the before, after or both callback Panelles
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If When And eMsgWhen.MSG_AFTER Then
            Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
'Parameters:
'lng_hWnd  - The handle of the window to be subclassed
'Returns;
'The sc_aSubData() index
    Const CODE_LEN As Long = 200                      'Length of the machine code in bytes
    Const FUNC_CWP As String = "CallWindowProcA"      'We use CallWindowProc to call the original WndProc
    Const FUNC_EBM As String = "EbMode"               'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
    Const FUNC_SWL As String = "SetWindowLongA"       'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
    Const MOD_USER As String = "user32"               'Location of the SetWindowLongA & CallWindowProc functions
    Const MOD_VBA5 As String = "vba5"                 'Location of the EbMode function if running VB5
    Const MOD_VBA6 As String = "vba6"                 'Location of the EbMode function if running VB6
    Const PATCH_01 As Long = 18                       'Code buffer offset to the location of the relative address to EbMode
    Const PATCH_02 As Long = 68                       'Address of the previous WndProc
    Const PATCH_03 As Long = 78                       'Relative address of SetWindowsLong
    Const PATCH_06 As Long = 116                      'Address of the previous WndProc
    Const PATCH_07 As Long = 121                      'Relative address of CallWindowProc
    Const PATCH_0A As Long = 186                      'Address of the owner object
    Static aBuf(1 To CODE_LEN) As Byte                'Static code buffer byte array
    Static pCWP As Long                               'Address of the CallWindowsProc
    Static pEbMode As Long                            'Address of the EbMode IDE break/stop/running function
    Static pSWL As Long                               'Address of the SetWindowsLong function
    Dim i As Long                                     'Loop index
    Dim j As Long                                     'Loop index
    Dim nSubIdx As Long                               'Subclass data index
    Dim sHex As String                                'Hex code string

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
            aBuf(j) = Val("&H" & Mid$(sHex, i, 2))    'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
            i = i + 2
        Loop                                          'Next pair of hex characters

        'Get API function addresses
        If Subclass_InIDE Then                        'If we're running in the VB IDE
            aBuf(16) = &H90                           'Patch the code buffer to enable the IDE state code
            aBuf(17) = &H90                           'Patch the code buffer to enable the IDE state code
            pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)   'Get the address of EbMode in vba6.dll
            If pEbMode = 0 Then                       'Found?
                pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)    'VB5 perhaps
            End If
        End If

        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)          'Get the address of the CallWindowsProc function
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)          'Get the address of the SetWindowLongA function
        ReDim sc_aSubData(0 To 0) As tSubData         'Create the first sc_aSubData element
    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If nSubIdx = -1 Then                          'If an sc_aSubData element isn't being re-cycled
            nSubIdx = UBound(sc_aSubData()) + 1       'Calculate the next element
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData    'Create a new sc_aSubData element
        End If

        Subclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)
        .hwnd = lng_hWnd                              'Store the hWnd
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)    'Allocate memory for the machine code WndProc
        .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)    'Set our WndProc in place
        Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)    'Copy the machine code from the static byte array to the code array in sc_aSubData
        Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)  'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)    'Original WndProc address for CallWindowProc, call the original WndProc
        Call zPatchRel(.nAddrSub, PATCH_03, pSWL)     'Patch the relative address of the SetWindowLongA api function
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)    'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
        Call zPatchRel(.nAddrSub, PATCH_07, pCWP)     'Patch the relative address of the CallWindowProc api function
        Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))    'Patch the address of this object instance into the static machine code buffer
    End With
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()
    Dim i As Long

    i = UBound(sc_aSubData())                         'Get the upper bound of the subclass data array
    Do While i >= 0                                   'Iterate through each element
        With sc_aSubData(i)
            If .hwnd <> 0 Then                        'If not previously Subclass_Stop'd
                Call Subclass_Stop(.hwnd)             'Subclass_Stop
            End If
        End With

        i = i - 1                                     'Next element
    Loop
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
'Parameters:
'lng_hWnd  - The handle of the window to stop being subclassed
    With sc_aSubData(zIdx(lng_hWnd))
        Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)    'Restore the original WndProc
        Call zPatchVal(.nAddrSub, PATCH_05, 0)        'Patch the Panelle B entry count to ensure no further 'before' callbacks
        Call zPatchVal(.nAddrSub, PATCH_09, 0)        'Patch the Panelle A entry count to ensure no further 'after' callbacks
        Call GlobalFree(.nAddrSub)                    'Release the machine code memory
        .hwnd = 0                                     'Mark the sc_aSubData element as available for re-use
        .nMsgCntB = 0                                 'Clear the before Panelle
        .nMsgCntA = 0                                 'Clear the after Panelle
        Erase .aMsgTblB                               'Erase the before Panelle
        Erase .aMsgTblA                               'Erase the after Panelle
    End With
End Sub

'======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry As Long                                'Message Panelle entry index
    Dim nOff1 As Long                                 'Machine code buffer offset 1
    Dim nOff2 As Long                                 'Machine code buffer offset 2

    If uMsg = ALL_MESSAGES Then                       'If all messages
        nMsgCnt = ALL_MESSAGES                        'Indicates that all messages will callback
    Else                                              'Else a specific message number
        Do While nEntry < nMsgCnt                     'For each existing entry. NB will skip if nMsgCnt = 0
            nEntry = nEntry + 1

            If aMsgTbl(nEntry) = 0 Then               'This msg Panelle slot is a deleted entry
                aMsgTbl(nEntry) = uMsg                'Re-use this entry
                Exit Sub                              'Bail
            ElseIf aMsgTbl(nEntry) = uMsg Then        'The msg is already in the Panelle!
                Exit Sub                              'Bail
            End If
        Loop                                          'Next entry

        nMsgCnt = nMsgCnt + 1                         'New slot required, bump the Panelle entry count
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long  'Bump the size of the Panelle.
        aMsgTbl(nMsgCnt) = uMsg                       'Store the message number in the Panelle
    End If

    If When = eMsgWhen.MSG_BEFORE Then                'If before
        nOff1 = PATCH_04                              'Offset to the Before Panelle
        nOff2 = PATCH_05                              'Offset to the Before Panelle entry count
    Else                                              'Else after
        nOff1 = PATCH_08                              'Offset to the After Panelle
        nOff2 = PATCH_09                              'Offset to the After Panelle entry count
    End If

    If uMsg <> ALL_MESSAGES Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))    'Address of the msg Panelle, has to be re-patched because Redim Preserve will move it in memory.
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)             'Patch the appropriate Panelle entry count
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc                            'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry As Long

    If uMsg = ALL_MESSAGES Then                       'If deleting all messages
        nMsgCnt = 0                                   'Message count is now zero
        If When = eMsgWhen.MSG_BEFORE Then            'If before
            nEntry = PATCH_05                         'Patch the before Panelle message count location
        Else                                          'Else after
            nEntry = PATCH_09                         'Patch the after Panelle message count location
        End If
        Call zPatchVal(nAddr, nEntry, 0)              'Patch the Panelle message count to zero
    Else                                              'Else deleteting a specific message
        Do While nEntry < nMsgCnt                     'For each Panelle entry
            nEntry = nEntry + 1
            If aMsgTbl(nEntry) = uMsg Then            'If this entry is the message we wish to delete
                aMsgTbl(nEntry) = 0                   'Mark the Panelle slot as available
                Exit Do                               'Bail
            End If
        Loop                                          'Next entry
    End If
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0                                'Iterate through the existing sc_aSubData() elements
        With sc_aSubData(zIdx)
            If .hwnd = lng_hWnd Then                  'If the hWnd of this element is the one we're looking for
                If Not bAdd Then                      'If we're searching not adding
                    Exit Function                     'Found
                End If
            ElseIf .hwnd = 0 Then                     'If this an element marked for reuse.
                If bAdd Then                          'If we're adding
                    Exit Function                     'Re-use it
                End If
            End If
        End With
        zIdx = zIdx - 1                               'Decrement the index
    Loop

    If Not bAdd Then
        Debug.Assert False                            'hWnd not found, programmer error
    End If

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



'---------------------------------------------------------------------------------------
' Name      : AppThemed
' DateTime  : 08 July 2005 09:29
' Author    : Gary Noble
' Purpose   : Tells Us If The System Is Themed Or not
'---------------------------------------------------------------------------------------
'
Public Function AppThemed() As Boolean

    On Error Resume Next
    AppThemed = IsAppThemed()
    On Error GoTo 0

End Function

'---------------------------------------------------------------------------------------
' Name      : BlendColor
' DateTime  : 08 July 2005 09:29
' Author    : Gary Noble
' Purpose   : Belnd Two colours at A Given Aplpha Value
'---------------------------------------------------------------------------------------
'
Private Property Get BlendColor(ByVal oColorFrom As OLE_COLOR, _
                                ByVal oColorTo As OLE_COLOR, _
                                Optional ByVal Alpha As Long = 128) As Long

    Dim lSrcR As Long

    Dim lSrcG As Long
    Dim lSrcB As Long
    Dim lDstR As Long
    Dim lDstG As Long
    Dim lDstB As Long
    Dim lCFrom As Long
    Dim lCTo As Long
    lCFrom = TranslateColor(oColorFrom)
    lCTo = TranslateColor(oColorTo)
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000
    BlendColor = RGB(((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255))

End Property
'---------------------------------------------------------------------------------------
' Name      : TranslateColor
' DateTime  : 08 July 2005 09:30
' Author    : Gary Noble
' Purpose   : Convert Automation color to Windows color
'---------------------------------------------------------------------------------------
'
Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                                Optional hPal As Long = 0) As Long

    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function


'---------------------------------------------------------------------------------------
' Name      : GetThemeName
' DateTime  : 08 July 2005 09:30
' Author    : Gary Noble
' Purpose   : Gets The Current Windows Theme Name
' Assumes   : m_sCurrentSystemThemename
'---------------------------------------------------------------------------------------
'
Private Sub GetThemeName(Optional hwnd As Long)

    Dim hTheme As Long
    Dim sShellStyle As String
    Dim sThemeFile As String
    Dim lPtrThemeFile As Long, lPtrColorName As Long, hres As Long
    Dim iPos As Long
    On Error GoTo GetThemeName_Error

    'On Error Resume Next
    hTheme = OpenThemeData(hwnd, StrPtr("ExplorerBar"))

    If Not hTheme = 0 Then
        ReDim bThemeFile(0 To 260 * 2) As Byte
        lPtrThemeFile = VarPtr(bThemeFile(0))
        ReDim bColorName(0 To 260 * 2) As Byte
        lPtrColorName = VarPtr(bColorName(0))
        hres = GetCurrentThemeName(lPtrThemeFile, 260, lPtrColorName, 260, 0, 0)

        sThemeFile = bThemeFile
        iPos = InStr(sThemeFile, vbNullChar)
        If (iPos > 1) Then sThemeFile = Left(sThemeFile, iPos - 1)
        m_sCurrentSystemThemename = bColorName
        iPos = InStr(m_sCurrentSystemThemename, vbNullChar)
        If (iPos > 1) Then m_sCurrentSystemThemename = Left(m_sCurrentSystemThemename, iPos - 1)

        sShellStyle = sThemeFile
        For iPos = Len(sThemeFile) To 1 Step -1
            If (Mid(sThemeFile, iPos, 1) = "\") Then
                sShellStyle = Left(sThemeFile, iPos)
                Exit For
            End If
        Next iPos
        sShellStyle = sShellStyle & "Shell\" & m_sCurrentSystemThemename & "\ShellStyle.dll"
        CloseThemeData hTheme
    Else
        m_sCurrentSystemThemename = "Classic"
    End If


    On Error GoTo 0
    Exit Sub

GetThemeName_Error:
    m_sCurrentSystemThemename = "Classic"
End Sub


'---------------------------------------------------------------------------------------
' Name      : Panels
' DateTime  : 08 July 2005 09:31
' Author    : Gary Noble
' Purpose   : Panels Collection
'---------------------------------------------------------------------------------------
'
Public Property Get Panels() As pTabs
Attribute Panels.VB_Description = "Tabs collection\r\n"

    If m_Panels Is Nothing Then Set m_Panels = New pTabs: Set m_Panels.ParentControl = Me:    ' m_Panels.Eventhandler = Me
    Set Panels = m_Panels

End Property

Friend Function Eventhandler(sData As String)

    If Not mb_LockUpdate Then Redraw
    Debug.Print "Event Handler Notify: " & sData

End Function

Private Sub UserControl_EnterFocus()
    UserControl_GotFocus
End Sub

Private Sub UserControl_ExitFocus()
    UserControl_LostFocus
End Sub

Private Sub UserControl_GotFocus()

    m_bInFocus = True
    Call Redraw

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

'-- Set Panels collection
    Set m_Panels = New pTabs

    '-- Set The Panels PanelControl
    Set m_Panels.ParentControl = Me


    '-- New MemDC
    Set m_oPaintDC = New cMemDC

    '-- icon size
    m_PanelIconPictureSize = m_def_PanelIconPictureSize

    '-- Default font
    Set UserControl.Font = Ambient.Font

    '-- Panel Style
    m_PanelStyle = m_def_PanelStyle

    '-- Default Min Control Height
    m_MinEmbebedControlHeight = m_def_MinEmbebedControlHeight

    '-- Refresh The Control
    pvRefreshScreen


    mb_LockUpdate = m_def_LockUpdate
    m_PanelAlignment = m_def_PanelAlignment
    m_UseCustomColor = m_def_UseCustomColor
    m_CustomColor = m_def_CustomColor
    m_SelectedItemColor = m_def_SelectedItemColor
    m_ShowTips = m_def_ShowTips

End Sub


'---------------------------------------------------------------------------------------
' Name      : Redraw
' DateTime  : 08 July 2005 09:33
' Author    : Gary Noble
' Purpose   : Redraws The Panels
'---------------------------------------------------------------------------------------
'
Friend Function Redraw()


    If Not mb_LockUpdate Then

        Select Case Me.PanelStyle
            Case Is = ETDS_Classic
                Call pvDrawPanelsStandard
            Case Is = ETDS_ClassicX
                Call pvDrawPanelsStandardX
            Case Is = ETDS_MSMessenger
                Call pvDrawPanelsMessengerStyle

        End Select

    End If

End Function

'---------------------------------------------------------------------------------------
' Name      : pvDrawPanelsStandard
' DateTime  : 08 July 2005 12:31
' Author    : Gary Noble
' Purpose   : Draws The Panels In The Standard/Default Way
'---------------------------------------------------------------------------------------
'
Private Sub pvDrawPanelsStandard()

    Dim yOffset As Long                               '-- Top Draw Offset
    Dim lFontpos As Long                              '-- Font Draw Pos
    Dim lindex As Long                                '-- Current Index Drawn
    Dim oPanel As pTab                                '-- Panel
    Dim lSelectedstart As Long
    Dim bRight As Boolean                             '-- Alignment
    Dim lsIndex As Long                               '-- Draw Bottom Line Index Flag

    '-- Clear
    Cls
    m_oPaintDC.Cls BackColor
    bRight = CBool(Me.PanelAlignment = EDPA_Right)


    '-- Set The Position On Where To Draw The Text and Icon
    lFontpos = (m_lPanelOffset / 2) - (m_lFontHeight / 2)

    If Not m_Panels Is Nothing Then

        If m_Panels.Count > 0 Then


            '-- loop Through The Panels And Draw Them
            For Each oPanel In m_Panels

                With New cMemDC

                    '-- New Drawing DC
                    '-- initialise To the Panel Size
                    .Init ScaleWidth, m_lPanelOffset, m_oPaintDC.hdc
                    .BackStyle = BS_TRANSPARENT
                    .ForeColor = UserControl.ForeColor

                    '-- Set The Font
                    Set .Font = UserControl.Font

                    '-- Process The Panels If The Selected Item Is Not Nothing
                    If (Not m_oSelectedItem Is Nothing) Then

                        '-- Process The Selected Item
                        If (oPanel.Key = m_oSelectedItem.Key) Then

                            '-- Fill The Rect
                            '-- This Makes The Border Of The Panel
                            .Rectangle 0, 0, ScaleWidth, m_lPanelOffset, m_lColorBorder, , m_lColorBorder

                            '-- Gradient The Icon Part Of The Panel
                            .FillGradient IIf(bRight, ScaleWidth - m_lPanelOffset, 1), 1, IIf(bRight, ScaleWidth - 1, m_lPanelOffset), m_lPanelOffset - 1, m_lColorOneNormal, m_lColorTwoNormal, True

                            '-- Gradient The Rest
                            .FillGradient IIf(bRight, 1, m_lPanelOffset), 1, IIf(bRight, ScaleWidth - m_lPanelOffset, ScaleWidth), m_lPanelOffset, m_lColorOneSelectedNormal, m_lColorTwoSelectedNormal, True

                            .ForeColor = SelectedItemColor


                        Else

                            '-- Process The Hover Item
                            If (Not m_oHoverItem Is Nothing) Then

                                If (oPanel.Key = m_oHoverItem.Key) Then

                                    '-- Fill The Rect
                                    '-- This Makes The Border Of The Panel
                                    .Rectangle 0, 0, ScaleWidth, m_lPanelOffset, m_lColorBorder, , m_lColorBorder

                                    '-- Gradient The Icon Part Of The Panel
                                    .FillGradient IIf(bRight, ScaleWidth - m_lPanelOffset, m_lPanelOffset), 1, IIf(bRight, ScaleWidth - 1, 1), m_lPanelOffset, BlendColor(m_lColorOneNormal, vbWhite, 230), BlendColor(m_lColorTwoNormal, vbWhite, 50), True

                                    '-- Gradient The Rest
                                    If m_bButtonDown Then
                                        .FillGradient IIf(bRight, 1, m_lPanelOffset), 1, IIf(bRight, ScaleWidth - m_lPanelOffset, ScaleWidth - 1), m_lPanelOffset, m_lColorOneSelected, m_lColorTwoSelected, True
                                    Else
                                        .FillGradient IIf(bRight, 1, m_lPanelOffset), 1, IIf(bRight, ScaleWidth - m_lPanelOffset, ScaleWidth - 1), m_lPanelOffset, m_lColorHotOne, m_lColorHotTwo, True
                                    End If

                                    '  .DrawLine IIf(bRight, ScaleWidth - m_lPanelOffset, m_lPanelOffset - 1), 0, IIf(bRight, ScaleWidth - m_lPanelOffset, m_lPanelOffset - 1), m_lPanelOffset, m_lColorBorder

                                Else
                                    '-- Fill The Rect
                                    '-- This Makes The Border Of The Panel
                                    .Rectangle 0, 0, ScaleWidth, m_lPanelOffset, m_lColorBorder, , m_lColorBorder

                                    '-- Gradient The Icon Part Of The Panel
                                    .FillGradient IIf(bRight, ScaleWidth - m_lPanelOffset, m_lPanelOffset), 1, IIf(bRight, ScaleWidth - 1, 1), m_lPanelOffset, BlendColor(m_lColorOneNormal, vbWhite, 100), BlendColor(m_lColorTwoNormal, vbWhite, 200), True

                                    '-- Gradient The Rest
                                    .FillGradient IIf(bRight, 1, m_lPanelOffset), 1, IIf(bRight, ScaleWidth - m_lPanelOffset, ScaleWidth - 1), m_lPanelOffset, m_lColorOneNormal, m_lColorTwoNormal, True

                                    '   .DrawLine IIf(bRight, ScaleWidth - m_lPanelOffset, m_lPanelOffset - 1), 0, IIf(bRight, ScaleWidth - m_lPanelOffset, m_lPanelOffset - 1), m_lPanelOffset, m_lColorBorder

                                End If
                            Else
                                '-- No Panels Selected
                                '-- A Panel Should Always Be Selected but Just Incase

                                '-- Fill The Rect
                                '-- This Makes The Border Of The Panel
                                .Rectangle 0, 0, ScaleWidth, m_lPanelOffset, m_lColorBorder, , m_lColorBorder

                                '-- Gradient The Icon Part Of The Panel
                                .FillGradient IIf(bRight, ScaleWidth - m_lPanelOffset, m_lPanelOffset), 1, IIf(bRight, ScaleWidth - 1, 1), m_lPanelOffset, BlendColor(m_lColorOneNormal, vbWhite, 100), BlendColor(m_lColorTwoNormal, vbWhite, 200), True

                                '-- Gradient The Rest
                                .FillGradient IIf(bRight, 1, m_lPanelOffset), 1, IIf(bRight, ScaleWidth - m_lPanelOffset, ScaleWidth - 1), m_lPanelOffset, m_lColorOneNormal, m_lColorTwoNormal, True

                                '.DrawLine IIf(bRight, ScaleWidth - m_lPanelOffset, m_lPanelOffset - 1), 0, IIf(bRight, ScaleWidth - m_lPanelOffset, m_lPanelOffset - 1), m_lPanelOffset, m_lColorBorder

                            End If
                        End If

                    End If

                    If Not Enabled Then .ForeColor = &H80000011

                    '-- Draw The Panel Text
                    .DrawText oPanel.Caption, IIf(bRight, 2, m_lPanelOffset + 5), lFontpos, IIf(bRight, ScaleWidth - m_lPanelOffset - 10, ScaleWidth), m_lPanelOffset - 1, IIf(bRight, DT_RIGHT Or DT_WORD_ELLIPSIS, DT_LEFT Or DT_WORD_ELLIPSIS)

                    .ForeColor = UserControl.ForeColor

                    '-- Draw The Panel Picture
                    If bRight Then
                        If Enabled Then
                            .PaintPicture oPanel.picIcon, ScaleWidth - ((m_lPanelOffset) / 2) - (Me.PanelIconPictureSize / 2), (m_lPanelOffset / 2) - (Me.PanelIconPictureSize / 2), PanelIconPictureSize, PanelIconPictureSize, , , , oPanel.PicMaskcolor
                        Else
                            .PaintDisabledPicture oPanel.picIcon, ScaleWidth - ((m_lPanelOffset) / 2) - (Me.PanelIconPictureSize / 2), (m_lPanelOffset / 2) - (Me.PanelIconPictureSize / 2), PanelIconPictureSize, PanelIconPictureSize
                        End If
                    Else
                        If Enabled Then
                            .PaintPicture oPanel.picIcon, (m_lPanelOffset / 2) - (Me.PanelIconPictureSize / 2), (m_lPanelOffset / 2) - (Me.PanelIconPictureSize / 2), PanelIconPictureSize, PanelIconPictureSize, , , , oPanel.PicMaskcolor
                        Else
                            .PaintDisabledPicture oPanel.picIcon, (m_lPanelOffset / 2) - (Me.PanelIconPictureSize / 2), (m_lPanelOffset / 2) - (Me.PanelIconPictureSize / 2), PanelIconPictureSize, PanelIconPictureSize
                        End If

                    End If

                    '-- Blt To The Main paintDC
                    .BitBlt m_oPaintDC.hdc, 0, yOffset

                End With

                '-- Set The Panel Rect
                oPanel.RecLeft = 0
                oPanel.RecTop = yOffset
                oPanel.RecRight = ScaleWidth
                oPanel.RecBottom = yOffset + m_lPanelOffset


                '-- Set The OffSets to Draw From The Bottom
                '-- As We Have Drawn The selected Panel
                If (Not m_oSelectedItem Is Nothing) Then

                    If (oPanel.Key = m_oSelectedItem.Key) Then

                        '-- Selected index
                        lsIndex = lindex + 1

                        '-- Set The Embedded control Boundaries
                        With m_RectControlBounderies
                            .Top = yOffset + m_lPanelOffset
                            .Left = IIf(bRight, 1, m_lPanelOffset)
                            .Right = IIf(bRight, (ScaleWidth - m_lPanelOffset) - 2, (ScaleWidth - (1 + m_lPanelOffset)))
                            .Bottom = yOffset - 2
                        End With

                        '-- Embebbed control draw Line Start
                        lSelectedstart = yOffset + m_lPanelOffset

                        yOffset = ScaleHeight - (m_lPanelOffset * ((m_Panels.Count - 1) - (lindex)))
                        yOffset = yOffset - (lindex - m_Panels.Count) - 2

                        '-- Set The Next Draw yOffset
                        If lSelectedstart + MinEmbebedControlHeight > yOffset Then
                            yOffset = lSelectedstart + MinEmbebedControlHeight
                        End If

                        m_RectControlBounderies.Bottom = yOffset - m_RectControlBounderies.Top

                        '-- Move And Make The Selected Control Visible
                        If Not m_oSelectedItem.EmbededControl Is Nothing Then
                            If m_RectControlBounderies.Bottom < 0 Then
                                m_oSelectedItem.EmbededControl.Enabled = Enabled
                                m_oSelectedItem.EmbededControl.Visible = False
                            Else
                                On Error Resume Next
                                m_oSelectedItem.EmbededControl.Move m_RectControlBounderies.Left * Screen.TwipsPerPixelX, m_RectControlBounderies.Top * Screen.TwipsPerPixelY, m_RectControlBounderies.Right * Screen.TwipsPerPixelX, m_RectControlBounderies.Bottom * Screen.TwipsPerPixelY
                                m_oSelectedItem.EmbededControl.Visible = True
                                m_oSelectedItem.EmbededControl.Enabled = Enabled
                            End If
                        End If

                    Else
                        yOffset = yOffset + m_lPanelOffset - 1
                    End If
                Else
                    yOffset = yOffset + m_lPanelOffset - 1
                End If

                lindex = lindex + 1

            Next

            '-- Draw the Bottom Line Of The Control
            If lsIndex = m_Panels.Count Then
                m_oPaintDC.DrawLine IIf(bRight, 0, m_lPanelOffset - 1), ScaleHeight - 1, IIf(bRight, ScaleWidth - m_lPanelOffset + 1, ScaleWidth), ScaleHeight - 1, m_lColorBorder
            Else
                m_oPaintDC.DrawLine IIf(bRight, 0, 0), ScaleHeight - 1, IIf(bRight, ScaleWidth, ScaleWidth - 1), ScaleHeight - 1, m_lColorBorder
            End If

            '-- Draw The Left Or Right BorderLine Of The Embbeded object
            If bRight Then
                m_oPaintDC.DrawLine 0, 0, 0, ScaleHeight, m_lColorBorder
            Else
                m_oPaintDC.DrawLine ScaleWidth - 1, 0, ScaleWidth - 1, ScaleHeight, m_lColorBorder
            End If

            '-- Draw The Line Between Panel Space
            m_oPaintDC.DrawLine IIf(bRight, ScaleWidth - m_lPanelOffset, m_lPanelOffset - 1), m_RectControlBounderies.Top, IIf(bRight, ScaleWidth - m_lPanelOffset, m_lPanelOffset - 1), m_RectControlBounderies.Bottom + m_RectControlBounderies.Top, m_lColorBorder


        Else
            '-- Control Border
            With m_oPaintDC
                .DrawLine 0, 0, ScaleWidth - 1, 0, m_lColorBorder
                .DrawLine 0, ScaleHeight - 1, ScaleWidth, ScaleHeight - 1, m_lColorBorder
                .DrawLine ScaleWidth - 1, 0, ScaleWidth - 1, ScaleHeight, m_lColorBorder
                .DrawLine 0, yOffset, 0, ScaleHeight - 1, m_lColorBorder
            End With

        End If
    Else
        '-- Control Border
        With m_oPaintDC
            .DrawLine 0, 0, ScaleWidth - 1, 0, m_lColorBorder
            .DrawLine 0, ScaleHeight - 1, ScaleWidth, ScaleHeight - 1, m_lColorBorder
            .DrawLine ScaleWidth - 1, 0, ScaleWidth - 1, ScaleHeight, m_lColorBorder
            .DrawLine 0, yOffset, 0, ScaleHeight - 1, m_lColorBorder
        End With
    End If

    m_oPaintDC.BitBlt hdc


End Sub




'---------------------------------------------------------------------------------------
' Name      : pvDrawPanelsStandardX
' DateTime  : 08 July 2005 12:31
' Author    : Gary Noble
' Purpose   : Draws The Panels In The Custom Way
'---------------------------------------------------------------------------------------
'
Private Sub pvDrawPanelsStandardX()

    Dim yOffset As Long                               '-- Top Draw Offset
    Dim lFontpos As Long                              '-- Font Draw Pos
    Dim lindex As Long                                '-- Current Index Drawn
    Dim oPanel As pTab                                '-- Panel
    Dim lSelectedstart As Long
    Dim bRight As Boolean                             '-- Alignment
    Dim lsIndex As Long                               '-- Draw Bottom Line Index Flag

    '-- Clear
    Cls
    m_oPaintDC.Cls BackColor
    bRight = CBool(Me.PanelAlignment = EDPA_Right)


    '-- Set The Position On Where To Draw The Text and Icon
    lFontpos = (m_lPanelOffset / 2) - (m_lFontHeight / 2)

    If Not m_Panels Is Nothing Then

        If m_Panels.Count > 0 Then


            '-- loop Through The Panels And Draw Them
            For Each oPanel In m_Panels

                With New cMemDC

                    '-- New Drawing DC
                    '-- initialise To the Panel Size
                    .Init ScaleWidth, m_lPanelOffset, m_oPaintDC.hdc
                    .BackStyle = BS_TRANSPARENT

                    '-- Set The Font
                    Set .Font = UserControl.Font

                    .ForeColor = UserControl.ForeColor

                    '-- Process The Panels If The Selected Item Is Not Nothing
                    If (Not m_oSelectedItem Is Nothing) Then

                        '-- Process The Selected Item
                        If (oPanel.Key = m_oSelectedItem.Key) Then

                            '-- Fill The Rect
                            '-- This Makes The Border Of The Panel
                            .Rectangle 0, 0, ScaleWidth, m_lPanelOffset, m_lColorBorder, , m_lColorBorder

                            '-- Gradient The Icon Part Of The Panel
                            .FillGradient IIf(bRight, ScaleWidth - m_lPanelOffset, 1), 1, IIf(bRight, ScaleWidth - 1, m_lPanelOffset), m_lPanelOffset - 1, IIf(Not bRight, m_lColorTwoNormal, vbWhite), IIf(Not bRight, vbWhite, m_lColorTwoNormal), False

                        Else

                            '-- Process The Hover Item
                            If (Not m_oHoverItem Is Nothing) Then

                                If (oPanel.Key = m_oHoverItem.Key) Then

                                    '-- Fill The Rect
                                    '-- This Makes The Border Of The Panel
                                    .Rectangle 0, 0, ScaleWidth, m_lPanelOffset, m_lColorBorder, , m_lColorBorder


                                    '-- Gradient The Rest
                                    If m_bButtonDown Then
                                        '-- Gradient The Icon Part Of The Panel
                                        .FillGradient IIf(bRight, ScaleWidth - m_lPanelOffset + 1, m_lPanelOffset), 1, IIf(bRight, ScaleWidth - 1, 1), m_lPanelOffset, m_lColorOneSelected, m_lColorTwoSelected, True
                                    Else
                                        '-- Gradient The Icon Part Of The Panel
                                        .FillGradient IIf(bRight, ScaleWidth - m_lPanelOffset + 1, m_lPanelOffset - 1), 1, IIf(bRight, ScaleWidth - 1, 1), m_lPanelOffset, BlendColor(m_lColorOneNormal, vbWhite, 230), BlendColor(m_lColorTwoNormal, vbWhite, 50), True
                                    End If

                                Else
                                    '-- Fill The Rect
                                    '-- This Makes The Border Of The Panel
                                    .Rectangle 0, 0, ScaleWidth, m_lPanelOffset, m_lColorBorder, , m_lColorBorder

                                    '-- Gradient The Icon Part Of The Panel
                                    .FillGradient IIf(bRight, ScaleWidth - m_lPanelOffset + 1, m_lPanelOffset - 1), 1, IIf(bRight, ScaleWidth - 1, 1), m_lPanelOffset, BlendColor(m_lColorOneNormal, vbWhite, 100), BlendColor(m_lColorTwoNormal, vbWhite, 200), True

                                End If
                            Else

                                '-- No Panels Selected
                                '-- A Panel Should Always Be Selected but Just Incase

                                '-- Fill The Rect
                                '-- This Makes The Border Of The Panel
                                .Rectangle 0, 0, ScaleWidth, m_lPanelOffset, m_lColorBorder, , m_lColorBorder

                                '-- Gradient The Icon Part Of The Panel
                                .FillGradient IIf(bRight, ScaleWidth - m_lPanelOffset + 1, m_lPanelOffset - 1), 1, IIf(bRight, ScaleWidth - 1, 1), m_lPanelOffset, BlendColor(m_lColorOneNormal, vbWhite, 100), BlendColor(m_lColorTwoNormal, vbWhite, 200), True

                            End If
                        End If

                    End If

                    '-- Draw The Panel Picture
                    If bRight Then
                        If Enabled Then
                            .PaintPicture oPanel.picIcon, ScaleWidth - ((m_lPanelOffset) / 2) - (Me.PanelIconPictureSize / 2), (m_lPanelOffset / 2) - (Me.PanelIconPictureSize / 2), PanelIconPictureSize, PanelIconPictureSize, , , , oPanel.PicMaskcolor
                        Else
                            .PaintDisabledPicture oPanel.picIcon, ScaleWidth - ((m_lPanelOffset) / 2) - (Me.PanelIconPictureSize / 2), (m_lPanelOffset / 2) - (Me.PanelIconPictureSize / 2), PanelIconPictureSize, PanelIconPictureSize
                        End If
                    Else
                        If Enabled Then
                            .PaintPicture oPanel.picIcon, (m_lPanelOffset / 2) - (Me.PanelIconPictureSize / 2), (m_lPanelOffset / 2) - (Me.PanelIconPictureSize / 2), PanelIconPictureSize, PanelIconPictureSize, , , , oPanel.PicMaskcolor
                        Else
                            .PaintDisabledPicture oPanel.picIcon, (m_lPanelOffset / 2) - (Me.PanelIconPictureSize / 2), (m_lPanelOffset / 2) - (Me.PanelIconPictureSize / 2), PanelIconPictureSize, PanelIconPictureSize
                        End If
                    End If

                    '-- Blt To The Main paintDC
                    .BitBlt m_oPaintDC.hdc, 0, yOffset

                End With

                '-- Set The Panel Rect
                oPanel.RecLeft = IIf(bRight, ScaleWidth - m_lPanelOffset, 0)
                oPanel.RecTop = yOffset
                oPanel.RecRight = IIf(bRight, ScaleWidth, m_lPanelOffset)
                oPanel.RecBottom = yOffset + m_lPanelOffset


                '-- Set The OffSets to Draw From The Bottom
                '-- As We Have Drawn The selected Panel
                If (Not m_oSelectedItem Is Nothing) Then

                    If (oPanel.Key = m_oSelectedItem.Key) Then

                        '-- Selected index
                        lsIndex = lindex + 1

                        '-- Set The Embedded control Boundaries
                        With m_RectControlBounderies
                            .Top = m_lPanelOffset
                            .Left = IIf(bRight, 1, m_lPanelOffset)
                            .Right = IIf(bRight, (ScaleWidth - m_lPanelOffset) - 1, (ScaleWidth - (1 + m_lPanelOffset)))
                            .Bottom = ScaleHeight - (m_lPanelOffset + lsIndex + 1)
                        End With

                        '-- Embebbed control draw Line Start
                        lSelectedstart = yOffset + m_lPanelOffset

                        '-- Set The Next Draw yOffset
                        yOffset = yOffset + m_lPanelOffset - 1
                        m_RectControlBounderies.Bottom = ScaleHeight - m_lPanelOffset - 1

                        '-- Move And Make The Selected Control Visible
                        If Not m_oSelectedItem.EmbededControl Is Nothing Then
                            If m_RectControlBounderies.Bottom < 0 Then
                                m_oSelectedItem.EmbededControl.Visible = False
                                m_oSelectedItem.EmbededControl.Enabled = Enabled
                            Else
                                On Error Resume Next
                                m_oSelectedItem.EmbededControl.Move m_RectControlBounderies.Left * Screen.TwipsPerPixelX, m_RectControlBounderies.Top * Screen.TwipsPerPixelY, m_RectControlBounderies.Right * Screen.TwipsPerPixelX, m_RectControlBounderies.Bottom * Screen.TwipsPerPixelY
                                m_oSelectedItem.EmbededControl.Visible = True
                                m_oSelectedItem.EmbededControl.Enabled = Enabled
                            End If
                        End If

                    Else
                        yOffset = yOffset + m_lPanelOffset - 1
                    End If
                Else
                    yOffset = yOffset + m_lPanelOffset - 1
                End If

                lindex = lindex + 1

            Next

            If lsIndex = m_Panels.Count Then
                m_oPaintDC.DrawLine IIf(bRight, 0, m_lPanelOffset - 1), ScaleHeight - 1, IIf(bRight, ScaleWidth - m_lPanelOffset, ScaleWidth), ScaleHeight - 1, m_lColorBorder
            Else
                m_oPaintDC.DrawLine IIf(bRight, 0, m_lPanelOffset - 1), ScaleHeight - 1, IIf(bRight, ScaleWidth - m_lPanelOffset, ScaleWidth), ScaleHeight - 1, m_lColorBorder
            End If

            '-- Draw The Left Or Right BorderLine Of The Embbeded object
            If bRight Then
                m_oPaintDC.DrawLine 0, 0, 0, ScaleHeight, m_lColorBorder
            Else
                m_oPaintDC.DrawLine ScaleWidth - 1, 0, ScaleWidth - 1, ScaleHeight, m_lColorBorder
            End If

            '-- Draw The Line Between Panel Space
            m_oPaintDC.DrawLine IIf(bRight, ScaleWidth - m_lPanelOffset, m_lPanelOffset - 1), lSelectedstart, IIf(bRight, ScaleWidth - m_lPanelOffset, m_lPanelOffset - 1), ScaleHeight - 1, m_lColorBorder

            '-- Draw the Bottom Line Of The Control
            m_oPaintDC.DrawLine IIf(Not bRight, 0, ScaleWidth), yOffset, IIf(bRight, ScaleWidth - m_lPanelOffset, ScaleWidth), yOffset, m_lColorBorder

            '-- Paint The Selected Item Caption At The Top
            If Not m_oSelectedItem Is Nothing Then
                m_oPaintDC.FillGradient IIf(bRight, 1, m_lPanelOffset), 1, IIf(bRight, ScaleWidth - m_lPanelOffset, ScaleWidth - 1), m_lPanelOffset, IIf(Not bRight, vbWhite, m_lColorOneNormal), IIf(Not bRight, m_lColorOneNormal, vbWhite), False
                m_oPaintDC.FillGradient IIf(bRight, 1, m_lPanelOffset), m_lPanelOffset - 3, IIf(bRight, ScaleWidth - m_lPanelOffset, ScaleWidth - 1), m_lPanelOffset, IIf(bRight, vbWhite, m_lColorOneNormal), IIf(bRight, m_lColorOneNormal, vbWhite), False
                m_oPaintDC.ForeColor = IIf(Enabled, m_SelectedItemColor, &H80000011)
                m_oPaintDC.DrawText m_oSelectedItem.Caption, IIf(bRight, 2, m_lPanelOffset + 5), lFontpos, IIf(bRight, ScaleWidth - m_lPanelOffset - 10, ScaleWidth), m_lPanelOffset - 1, IIf(bRight, DT_RIGHT Or DT_WORD_ELLIPSIS, DT_LEFT Or DT_WORD_ELLIPSIS)
            End If


        Else
            '-- Control Border
            With m_oPaintDC
                .DrawLine 0, 0, ScaleWidth - 1, 0, m_lColorBorder
                .DrawLine 0, ScaleHeight - 1, ScaleWidth, ScaleHeight - 1, m_lColorBorder
                .DrawLine ScaleWidth - 1, 0, ScaleWidth - 1, ScaleHeight, m_lColorBorder
                .DrawLine 0, yOffset, 0, ScaleHeight - 1, m_lColorBorder
            End With

        End If
    Else
        '-- Control Border
        With m_oPaintDC
            .DrawLine 0, 0, ScaleWidth - 1, 0, m_lColorBorder
            .DrawLine 0, ScaleHeight - 1, ScaleWidth, ScaleHeight - 1, m_lColorBorder
            .DrawLine ScaleWidth - 1, 0, ScaleWidth - 1, ScaleHeight, m_lColorBorder
            .DrawLine 0, yOffset, 0, ScaleHeight - 1, m_lColorBorder
        End With
    End If

    m_oPaintDC.BitBlt hdc



End Sub

Public Property Get PanelIconPictureSize() As EPanel_PicSize
Attribute PanelIconPictureSize.VB_Description = "Default Icon Size"
    PanelIconPictureSize = m_PanelIconPictureSize
End Property

Public Property Let PanelIconPictureSize(ByVal New_PanelIconPictureSize As EPanel_PicSize)
    m_PanelIconPictureSize = New_PanelIconPictureSize
    PropertyChanged "PanelIconPictureSize"
    pvRefreshScreen

End Property

Private Sub UserControl_LostFocus()
    m_bInFocus = False
    Set m_oHoverItem = Nothing
    Call Redraw
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)


    RaiseEvent PanelMouseDown(Button)

    If Button = vbLeftButton Then
        m_bButtonDown = True
        Call Redraw
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)


    Dim oHittest As pTab



    If Button = vbLeftButton Then Exit Sub

    Set oHittest = Hittest(x, y)


    If oHittest Is Nothing Then

        Call pvDestroyTooltip

        '-- Clear The Hoveritem
        If Not m_oHoverItem Is Nothing Then
            Set m_oHoverItem = Nothing
            Redraw
        End If
        Exit Sub

    End If

    If Not m_oHoverItem Is Nothing Then

        '-- Set The hovering Item
        If Not m_oHoverItem.Key = oHittest.Key Then
            Set m_oHoverItem = oHittest
            RaiseEvent PanelHovering(oHittest)
            Call Redraw
            pvSetToolTip oHittest.Caption, oHittest.ToolTipText, True
        End If

    Else
        Set m_oHoverItem = oHittest
        RaiseEvent PanelHovering(oHittest)
        pvSetToolTip oHittest.Caption, oHittest.ToolTipText, True
        Call Redraw

    End If


End Sub
Private Sub pvSetToolTip(ByVal sHeader As String, ByVal sToolTip As String, Optional bDestroy As Boolean = False)

    If bDestroy Then pvDestroyTooltip

    If m_ShowTips Then
        With Me
            ToolTipStyle = mvarToolTipStyle
            ToolTipTitle = sHeader
            TipText = sToolTip
            VisibleTime = 4000
            DelayTime = 400
            pvCreateToolTip UserControl.hwnd
        End With
    End If

End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim oHittest As pTab

    m_bButtonDown = False

    If Button = vbLeftButton Then
        Set oHittest = Hittest(x, y)
        If Not oHittest Is Nothing Then
            If Not m_oHoverItem Is Nothing Then
                If oHittest.Key = m_oHoverItem.Key Then
                    If Not SelectedItem.Key = oHittest.Key Then
                        Set Me.SelectedItem = oHittest
                        RaiseEvent PanelSelected(oHittest)
                    End If
                End If
            End If
        End If
        Set m_oHoverItem = Nothing
        Redraw
    End If

    RaiseEvent PanelMouseUp(Button)

End Sub

Private Sub UserControl_Paint()
    pvRefreshScreen
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'-- Tooltip
    mvarDelayTime = 500
    mvarVisibleTime = 5000


    ToolTipStyle = PropBag.ReadProperty("ToolTipStyle", mvarToolTipStyle)
    m_PanelIconPictureSize = PropBag.ReadProperty("PanelIconPictureSize", m_def_PanelIconPictureSize)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)

    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_PanelStyle = PropBag.ReadProperty("PanelStyle", m_def_PanelStyle)

    m_MinEmbebedControlHeight = PropBag.ReadProperty("MinEmbebedControlHeight", m_def_MinEmbebedControlHeight)
    mb_LockUpdate = PropBag.ReadProperty("LockUpdate", m_def_LockUpdate)
    m_PanelAlignment = PropBag.ReadProperty("PanelAlignment", m_def_PanelAlignment)



    If Ambient.UserMode Then                          'If we're not in design mode
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")

        If Not bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
                bTrack = False
            End If
        End If

        If bTrack Then
            'OS supports mouse leave so subclass for it
            With UserControl
                'Start subclassing the UserControl
                Call Subclass_Start(.hwnd)
                Call Subclass_AddMsg(.hwnd, WM_MOUSEMOVE, MSG_AFTER)
                Call Subclass_AddMsg(.hwnd, WM_MOUSELEAVE, MSG_AFTER)
                Call Subclass_AddMsg(.hwnd, WM_SYSCOLORCHANGE, MSG_AFTER)
                Call Subclass_AddMsg(.hwnd, WM_THEMECHANGED, MSG_AFTER)
            End With
        End If
    End If


    m_UseCustomColor = PropBag.ReadProperty("UseCustomColor", m_def_UseCustomColor)
    m_CustomColor = PropBag.ReadProperty("CustomColor", m_def_CustomColor)
    m_SelectedItemColor = PropBag.ReadProperty("SelectedItemColor", m_def_SelectedItemColor)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_ShowTips = PropBag.ReadProperty("ShowTips", m_def_ShowTips)
End Sub

Private Sub UserControl_Resize()

    pvRefreshScreen

End Sub

'---------------------------------------------------------------------------------------
' Name      : pvRefreshScreen
' DateTime  : 08 July 2005 09:55
' Author    : Gary Noble
' Purpose   : Refresh The Actual Control
'---------------------------------------------------------------------------------------
'
Private Sub pvRefreshScreen()


    If m_oPaintDC Is Nothing Then Set m_oPaintDC = New cMemDC

    '-- Get The Windows Theme Name
    Call GetThemeName(hwnd)

    '-- Set The Gradient Colors
    Call pvGetGradientColors


    '-- Set The Painting Params
    With m_oPaintDC

        .Init ScaleWidth, ScaleHeight, UserControl.hdc

        Set .Font = UserControl.Font

        '-- Set The Font Height
        m_lFontHeight = .TextHeight("',")

        '-- set The Panel Height
        m_lPanelOffset = (Me.PanelIconPictureSize + 12)

        '-- ReAdjust The Panel Height If The Panel Height Is Smaller Than The Font Height
        If m_lPanelOffset < m_lFontHeight Then m_lPanelOffset = m_lFontHeight + 12

        '-- Blit The BackGround
        .FillRect 0, 0, ScaleWidth, ScaleHeight, BackColor, m_lColorBorder

    End With

    '-- REDRAW
    Call Redraw


End Sub

Private Sub UserControl_Terminate()
    On Error GoTo Catch
    'Stop all subclassing
    Call Subclass_StopAll
    '-- Tooltip
    Call pvDestroyTooltip
Catch:
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("PanelIconPictureSize", m_PanelIconPictureSize, m_def_PanelIconPictureSize)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("PanelStyle", m_PanelStyle, m_def_PanelStyle)
    Call PropBag.WriteProperty("ToolTipStyle", mvarToolTipStyle, 0)
    Call PropBag.WriteProperty("MinEmbebedControlHeight", m_MinEmbebedControlHeight, m_def_MinEmbebedControlHeight)
    Call PropBag.WriteProperty("LockUpdate", mb_LockUpdate, m_def_LockUpdate)
    Call PropBag.WriteProperty("PanelAlignment", m_PanelAlignment, m_def_PanelAlignment)
    Call PropBag.WriteProperty("UseCustomColor", m_UseCustomColor, m_def_UseCustomColor)
    Call PropBag.WriteProperty("CustomColor", m_CustomColor, m_def_CustomColor)
    Call PropBag.WriteProperty("SelectedItemColor", m_SelectedItemColor, m_def_SelectedItemColor)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ShowTips", m_ShowTips, m_def_ShowTips)
End Sub


Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    Call Redraw
End Property

Public Property Get SelectedItem() As pTab
    Set SelectedItem = m_oSelectedItem
End Property

'---------------------------------------------------------------------------------------
' Name      : SelectedItem
' DateTime  : 08 July 2005 09:40
' Author    : Gary Noble
' Purpose   : Sets The selected item
' Note      : if NO Panel Is Selected Then By Default We Select The First One
'---------------------------------------------------------------------------------------
'
Public Property Set SelectedItem(oPanel As pTab)
    On Error Resume Next

    If Not m_oSelectedItem Is Nothing Then
        If Not m_oSelectedItem.EmbededControl Is Nothing Then m_oSelectedItem.EmbededControl.Visible = False
    End If

    Set m_oSelectedItem = oPanel

    If Not mb_LockUpdate Then Call Redraw

    '-- Make The Embebbed Control visible
    If Not m_oSelectedItem.EmbededControl Is Nothing Then
        m_oSelectedItem.EmbededControl.Visible = True
        m_oSelectedItem.EmbededControl.SetFocus
        m_bInFocus = True
    End If

    On Error GoTo 0
    Exit Property

SelectedItem_Error:

    Err.Raise Err.Number, App.ProductName & " _pTab SelectedItem", Err.Description

End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Call pvRefreshScreen
End Property

'---------------------------------------------------------------------------------------
' Name      : pvGetGradientColors
' DateTime  : 08 July 2005 09:45
' Author    : Gary Noble
' Purpose   : Default System Colour Handler
'---------------------------------------------------------------------------------------
'
Sub pvGetGradientColors()


    m_lColorOneSelected = 1
    m_lColorTwoSelected = 1
    m_lColorHeaderColorOne = 1
    m_lColorHotOne = 1
    m_lColorHotTwo = 1

    If AppThemed And UseCustomColor = False Then
        Select Case m_sCurrentSystemThemename
            Case "HomeStead"
                m_lColorOneNormal = RGB(228, 235, 200)
                m_lColorTwoNormal = RGB(175, 194, 142)
                m_lColorBorder = RGB(100, 144, 88)
                m_lColorHeaderColorOne = RGB(165, 182, 121)
                m_lColorHeaderColorTwo = BlendColor(RGB(99, 122, 68), vbBlack, 200)
            Case "NormalColor"
                m_lColorOneNormal = RGB(197, 221, 250)
                m_lColorTwoNormal = RGB(128, 167, 225)
                m_lColorBorder = RGB(0, 45, 150)
                m_lColorHeaderColorOne = RGB(81, 128, 208)
                m_lColorHeaderColorTwo = BlendColor(RGB(11, 63, 153), vbBlack, 230)
            Case "Metallic"
                m_lColorOneNormal = RGB(219, 220, 232)
                m_lColorTwoNormal = RGB(149, 147, 177)
                m_lColorBorder = RGB(119, 118, 151)
                m_lColorHeaderColorOne = RGB(163, 162, 187)
                m_lColorHeaderColorTwo = BlendColor(RGB(112, 111, 145), vbBlack, 200)
            Case Else
                m_lColorOneNormal = BlendColor(vbButtonFace, vbWhite, 120)
                m_lColorTwoNormal = vbButtonFace
                m_lColorBorder = BlendColor(vbButtonFace, vbBlack, 200)
                m_lColorHeaderColorOne = vbButtonFace
                m_lColorHeaderColorTwo = BlendColor(vbInactiveTitleBar, vbBlack, 200)
                m_lColorBorder = TranslateColor(vbInactiveTitleBar)
        End Select
        m_lColorOneSelectedNormal = RGB(248, 216, 126)
        m_lColorTwoSelectedNormal = RGB(240, 160, 38)
        m_lColorHotOne = BlendColor(vbWindowBackground, vbButtonFace, 220)
        m_lColorHotTwo = RGB(248, 216, 126)
        m_lColorOneSelected = RGB(240, 160, 38)
        m_lColorTwoSelected = RGB(248, 216, 126)
    ElseIf UseCustomColor = False Then
        m_lColorOneNormal = BlendColor(vbButtonFace, vbWhite, 120)
        m_lColorTwoNormal = vbButtonFace
        m_lColorBorder = BlendColor(vbButtonFace, vbBlack, 200)
        m_lColorHeaderColorOne = vbButtonFace
        m_lColorHeaderColorTwo = BlendColor(vbInactiveTitleBar, BlendColor(vbBlack, vbButtonFace, 10), 200)
        m_lColorBorder = TranslateColor(vbInactiveTitleBar)
        m_lColorHotTwo = BlendColor(vbInactiveTitleBar, BlendColor(vbButtonFace, vbWhite, 50), 10)
        m_lColorHotOne = m_lColorHotTwo
        m_lColorOneSelected = BlendColor(vbInactiveTitleBar, BlendColor(vbButtonFace, vbWhite, 150), 100)
        m_lColorTwoSelected = m_lColorOneSelected
        m_lColorOneSelectedNormal = BlendColor(vbInactiveTitleBar, BlendColor(vbButtonFace, vbWhite, 150), 130)
        m_lColorTwoSelectedNormal = m_lColorOneSelectedNormal
    ElseIf UseCustomColor Then

        m_lColorOneNormal = BlendColor(TranslateColor(m_CustomColor), vbWhite, 150)
        m_lColorTwoNormal = TranslateColor(m_CustomColor)
        m_lColorBorder = BlendColor(TranslateColor(m_CustomColor), vbBlack, 150)
        m_lColorHeaderColorOne = BlendColor(TranslateColor(m_CustomColor), vbWhite, 120)
        m_lColorHeaderColorTwo = BlendColor(TranslateColor(m_CustomColor), vbBlack, 200)    'BlendColor(TranslateColor(m_CustomColor), BlendColor(TranslateColor(m_CustomColor), vbWhite), 200)
        m_lColorHotTwo = BlendColor(TranslateColor(m_CustomColor), vbWhite, 70)    ' BlendColor(TranslateColor(m_CustomColor), vbBlack, 230)
        m_lColorHotOne = BlendColor(TranslateColor(m_CustomColor), vbWhite, 30)    ' BlendColor(TranslateColor(m_CustomColor), BlendColor(TranslateColor(m_CustomColor), vbWhite, 150), 130)

        m_lColorTwoSelected = BlendColor(TranslateColor(m_CustomColor), BlendColor(TranslateColor(m_CustomColor), vbWhite, 150), 130)
        m_lColorOneSelected = BlendColor(TranslateColor(m_CustomColor), vbBlack, 230)

        m_lColorOneSelectedNormal = BlendColor(TranslateColor(m_CustomColor), vbWhite, 30)
        m_lColorTwoSelectedNormal = BlendColor(TranslateColor(m_CustomColor), vbWhite, 70)

    End If

End Sub

'---------------------------------------------------------------------------------------
' Name      : Clear
' DateTime  : 08 July 2005 09:44
' Author    : Gary Noble
' Purpose   : Does Exactly What It Says On The Tin
'---------------------------------------------------------------------------------------
'
Public Sub Clear()

    On Error GoTo Clear_Error

    Set m_Panels = Nothing
    Set m_oHoverItem = Nothing
    ResetSelected
    pvRefreshScreen

    On Error GoTo 0
    Exit Sub

Clear_Error:

    Err.Raise Err.Number, App.ProductName & " _pTab Clear", Err.Description

End Sub

Public Property Get PanelStyle() As EPanel_DrawStyle
    PanelStyle = m_PanelStyle
End Property

Public Property Let PanelStyle(ByVal New_PanelStyle As EPanel_DrawStyle)
    m_PanelStyle = New_PanelStyle
    PropertyChanged "PanelStyle"
    pvGetGradientColors
    Call Redraw
End Property


'---------------------------------------------------------------------------------------
' Name      : Hittest
' DateTime  : 08 July 2005 09:42
' Author    : Gary Noble
' Purpose   : Determines If We Are in An Item or Hovering an Item
'---------------------------------------------------------------------------------------
'
Private Function Hittest(ByVal x As Single, ByVal y As Single) As pTab

    Dim oPanel As pTab
    Dim RC As RECT

    If Not m_Panels Is Nothing Then
        For Each oPanel In m_Panels

            '-- Get the Panel Boundaries
            With RC
                .Top = oPanel.RecTop
                .Bottom = oPanel.RecBottom
                .Left = oPanel.RecLeft
                .Right = oPanel.RecRight
            End With

            '-- Are We In the Panel Boundaries
            If PtInRect(RC, x, y) Then
                Set Hittest = oPanel
                Exit For
            End If


        Next
    End If


End Function

'---------------------------------------------------------------------------------------
' Name      : ResetSelected
' DateTime  : 08 July 2005 09:43
' Author    : Gary Noble
' Purpose   : Resets The Selected Panel Incase The Selected Panel Is Nothing
'---------------------------------------------------------------------------------------
'
Friend Function ResetSelected()

    Set Me.SelectedItem = Nothing
    Set m_oHoverItem = Nothing

    If Not m_Panels Is Nothing Then
        If Me.SelectedItem Is Nothing Then
            If m_Panels.Count > 0 Then
                Set Me.SelectedItem = m_Panels(1)
            End If
        End If
    End If

End Function

Public Property Get hwnd() As Long

    hwnd = UserControl.hwnd
End Property

Public Property Get MinEmbebedControlHeight() As Long
    MinEmbebedControlHeight = m_MinEmbebedControlHeight
End Property

Public Property Let MinEmbebedControlHeight(ByVal New_MinEmbebedControlHeight As Long)

    m_MinEmbebedControlHeight = New_MinEmbebedControlHeight
    PropertyChanged "MinEmbebedControlHeight"
End Property


Public Property Get LockUpdate() As Boolean
    LockUpdate = mb_LockUpdate
End Property

Public Property Let LockUpdate(ByVal New_LockUpdate As Boolean)
    mb_LockUpdate = New_LockUpdate
    PropertyChanged "LockUpdate"
    If Not mb_LockUpdate Then Redraw
End Property

Public Property Get PanelAlignment() As EPanel_Alignment
    PanelAlignment = m_PanelAlignment
End Property

Public Property Let PanelAlignment(ByVal New_PanelAlignment As EPanel_Alignment)
    m_PanelAlignment = New_PanelAlignment
    PropertyChanged "PanelAlignment"
    Redraw
End Property

Public Property Get UseCustomColor() As Boolean
    UseCustomColor = m_UseCustomColor
End Property

Public Property Let UseCustomColor(ByVal New_UseCustomColor As Boolean)
    m_UseCustomColor = New_UseCustomColor
    PropertyChanged "UseCustomColor"
    GetThemeName hwnd
    pvGetGradientColors
    Call Redraw
End Property

Public Property Get CustomColor() As OLE_COLOR
    CustomColor = m_CustomColor
End Property

Public Property Let CustomColor(ByVal New_CustomColor As OLE_COLOR)
    m_CustomColor = New_CustomColor
    PropertyChanged "CustomColor"
    Call Redraw
End Property

'---------------------------------------------------------------------------------------
' Name      : PanelColours
' DateTime  : 08 July 2005 15:30
' Author    : Gary Noble
' Purpose   : Returns The Colors Used To Paint the Panels
' notes     : This Could Be Good, Just Incase You Want To Set Other Colors In
'             Your App To The Same Colors Used To Paint The Tabs
'---------------------------------------------------------------------------------------
'
Public Sub PanelColours(ByVal lColorOneNormal As Long, ByVal lColorTwoNormal As Long, _
                        ByVal lColorOneSelected As Long, ByVal lColorTwoSelected As Long, _
                        ByVal lColorOneSelectedNormal As Long, ByVal lColorTwoSelectedNormal As Long, _
                        ByVal lColorBorder As Long, ByVal lColorOneHot As Long, ByVal lColorTwoHot As Long)

    lColorOneNormal = m_lColorOneNormal
    lColorTwoNormal = m_lColorTwoNormal
    lColorOneSelected = m_lColorOneSelected
    lColorTwoSelected = m_lColorTwoSelected
    lColorOneSelectedNormal = m_lColorOneSelectedNormal
    lColorTwoSelectedNormal = m_lColorTwoSelectedNormal
    lColorOneHot = m_lColorHotOne
    lColorTwoHot = m_lColorHotTwo
    lColorBorder = m_lColorBorder


End Sub

Public Property Get SelectedItemColor() As OLE_COLOR
    SelectedItemColor = m_SelectedItemColor
End Property

Public Property Let SelectedItemColor(ByVal New_SelectedItemColor As OLE_COLOR)
    m_SelectedItemColor = New_SelectedItemColor
    PropertyChanged "SelectedItemColor"
    Redraw
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    Call Redraw
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    Redraw
End Property

Public Property Get ShowTips() As Boolean
    ShowTips = m_ShowTips
End Property

Public Property Let ShowTips(ByVal New_ShowTips As Boolean)
    m_ShowTips = New_ShowTips
    PropertyChanged "ShowTips"
End Property



'---------------------------------------------------------------------------------------
' Name      : pvDrawMessengerPanel
' DateTime  : 12 July 2005 11:05
' Author    : Gary Noble
' Purpose   : Paints A Messenger Style Panel
'---------------------------------------------------------------------------------------
'
Private Sub pvDrawMessengerPanel(oItem As pTab, yOffset As Long, Optional bSelected As Boolean, Optional bHovering As Boolean)

    Dim lFontpos As Long                              '-- Font Draw Pos
    Dim PanelPoly() As POINTAPI                       '-- Poly angles
    Dim oType As Long                                 '-- Selected/hovering Flag
    Dim oDraw As cMemDC                               '-- Drawing DC
    Dim bRight As Boolean                             '-- Alignment
    Dim lPolyCount As Long                            '-- Poly count
    Dim lRgn As Long                                  '-- poly Region Ptr

'-- New Draw DC
    Set oDraw = New cMemDC

    '-- Is Right To left
    bRight = CBool(Me.PanelAlignment = EDPA_Right)

    '-- Set The Poly Angles
    '-- This Also Takes Into Consideration The Right To Left Factor
    ReDim PanelPoly(1 To 6) As POINTAPI

    PanelPoly(1).x = IIf(Not bRight, 0, ScaleWidth)
    PanelPoly(1).y = IIf(Not bRight, 5, 5)
    PanelPoly(2).x = IIf(Not bRight, 5, ScaleWidth - 5)
    PanelPoly(2).y = IIf(Not bRight, 0, 0)
    PanelPoly(3).x = IIf(Not bRight, m_lPanelOffset + 25, ScaleWidth - m_lPanelOffset - 25)
    PanelPoly(3).y = 0
    PanelPoly(4).x = IIf(Not bRight, m_lPanelOffset + 15, ScaleWidth - m_lPanelOffset - 15)
    PanelPoly(4).y = IIf(Not bRight, m_lPanelOffset + 25, m_lPanelOffset + 25)
    PanelPoly(5).x = IIf(Not bRight, m_lPanelOffset, ScaleWidth - m_lPanelOffset)
    PanelPoly(5).y = m_lPanelOffset + 22
    PanelPoly(6).x = IIf(Not bRight, 0, ScaleWidth)
    PanelPoly(6).y = IIf(Not bRight, m_lPanelOffset + 2, m_lPanelOffset + 2)
    lPolyCount = 6


    '-- Set Type Flag
    If bSelected Then
        oType = 1
    ElseIf bHovering Then
        oType = 2
    Else
        oType = 3
    End If

    '-- Draw The Panel
    With oDraw

        '-- New Drawing DC
        '-- initialise To the Panel Size
        .Init ScaleWidth, m_lPanelOffset + 25, m_oPaintDC.hdc
        .Cls IIf(bSelected, &H400040, BackColor)
               
        

        Select Case oType

                '-- Selected
            Case Is = 1

                ' Draw tab gradient background
                lRgn = CreatePolygonRgn(PanelPoly(1), lPolyCount, ALTERNATE)
                SelectClipRgn .hdc, lRgn

                .FillGradient IIf(bRight, ScaleWidth - m_lPanelOffset, 1), 1, IIf(bRight, ScaleWidth - 1, m_lPanelOffset), m_lPanelOffset + 25 - 1, vbWhite, vbWhite, True

                SelectClipRgn .hdc, 0
                DeleteObject lRgn

            Case Is = 2
                '-- Hovering

                '-- This Makes The Border Of The Panel
                lRgn = CreatePolygonRgn(PanelPoly(1), lPolyCount, ALTERNATE)
                SelectClipRgn .hdc, lRgn

                If m_bButtonDown Then
                    .FillGradient IIf(bRight, ScaleWidth - m_lPanelOffset, m_lPanelOffset), 1, IIf(bRight, ScaleWidth - 1, 1), m_lPanelOffset + 25, BlendColor(m_lColorTwoNormal, vbWhite, 50), BlendColor(m_lColorTwoNormal, vbBlack, 210), True
                Else
                    .FillGradient IIf(bRight, ScaleWidth - m_lPanelOffset, m_lPanelOffset), 1, IIf(bRight, ScaleWidth - 1, 1), m_lPanelOffset + 25, BlendColor(m_lColorBorder, vbWhite, 200), BlendColor(m_lColorTwoNormal, vbWhite, 50), True
                End If

                SelectClipRgn .hdc, 0
                DeleteObject lRgn

            Case Is = 3

                '-- No Panels Selected
                '-- A Panel Should Always Be Selected but Just Incase

                '-- This Makes The Border Of The Panel
                lRgn = CreatePolygonRgn(PanelPoly(1), lPolyCount, ALTERNATE)
                SelectClipRgn .hdc, lRgn

                .FillGradient IIf(bRight, ScaleWidth - m_lPanelOffset, m_lPanelOffset), 1, IIf(bRight, ScaleWidth - 1, 1), m_lPanelOffset + 25, BlendColor(m_lColorOneNormal, vbWhite, 100), BlendColor(m_lColorTwoNormal, vbWhite, 200), True

                SelectClipRgn .hdc, 0
                DeleteObject lRgn


        End Select


        '-- Draw The Border line Round The polygon
        If Not bRight Then
            .DrawLine 5, 0, ScaleWidth, 0, m_lColorBorder
            .DrawLine 0, 5, 0, m_lPanelOffset, m_lColorBorder
            .DrawLine 5, 0, 0, 5, m_lColorBorder
            .DrawLine 0, m_lPanelOffset, 3, m_lPanelOffset + 5, m_lColorBorder
            .DrawLine 2, m_lPanelOffset + 4, m_lPanelOffset + 5, m_lPanelOffset + 25, m_lColorBorder
        Else
            .DrawLine 0, 0, ScaleWidth - 5, 0, m_lColorBorder
            .DrawLine ScaleWidth - 1, 5, ScaleWidth - 1, m_lPanelOffset + 1, m_lColorBorder
            .DrawLine ScaleWidth - 5, 0, ScaleWidth, 5, m_lColorBorder
            .DrawLine ScaleWidth, m_lPanelOffset, ScaleWidth - 3, m_lPanelOffset + 5, m_lColorBorder
            .DrawLine ScaleWidth - 2, m_lPanelOffset + 4, ScaleWidth - (m_lPanelOffset + 4), m_lPanelOffset + 25, m_lColorBorder
        End If

        '        MsgBox oItem.PicMaskcolor & vbCrLf & oItem.Caption

        '-- Draw The Panel Picture
        If bRight Then

            If Enabled Then
                .PaintPicture oItem.picIcon, ScaleWidth - ((m_lPanelOffset) / 2) - (Me.PanelIconPictureSize / 2), (m_lPanelOffset / 2) - (Me.PanelIconPictureSize / 2), PanelIconPictureSize, PanelIconPictureSize, , , , oItem.PicMaskcolor
            Else
                .PaintDisabledPicture oItem.picIcon, ScaleWidth - ((m_lPanelOffset) / 2) - (Me.PanelIconPictureSize / 2), (m_lPanelOffset / 2) - (Me.PanelIconPictureSize / 2), PanelIconPictureSize, PanelIconPictureSize
            End If

        Else

            If Enabled Then
                .PaintPicture oItem.picIcon, (m_lPanelOffset / 2) - (Me.PanelIconPictureSize / 2), (m_lPanelOffset / 2) - (Me.PanelIconPictureSize / 2), PanelIconPictureSize, PanelIconPictureSize, , , , oItem.PicMaskcolor
            Else
                .PaintDisabledPicture oItem.picIcon, (m_lPanelOffset / 2) - (Me.PanelIconPictureSize / 2), (m_lPanelOffset / 2) - (Me.PanelIconPictureSize / 2), PanelIconPictureSize, PanelIconPictureSize
            End If

        End If

        .DrawLine IIf(bRight, ScaleWidth - m_lPanelOffset, m_lPanelOffset - 1), 1, IIf(bRight, ScaleWidth - m_lPanelOffset, m_lPanelOffset - 1), m_lPanelOffset + 21, vbWhite
    
        If bSelected Then
            .TransBlt m_oPaintDC.hdc, 0, yOffset, , , , , &H400040
        Else
            .BitBlt m_oPaintDC.hdc, 0, yOffset
        End If

    End With

    Set oDraw = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Name      : pvDrawPanelsMessengerStyle
' DateTime  : 08 July 2005 12:31
' Author    : Gary Noble
' Purpose   : Draws The Panels similar To MSN Messenger
'---------------------------------------------------------------------------------------
'
Private Sub pvDrawPanelsMessengerStyle()

    Dim yOffset As Long                               '-- Top Draw Offset
    Dim lFontpos As Long                              '-- Font Draw Pos
    Dim lindex As Long                                '-- Current Index Drawn
    Dim oPanel As pTab                                '-- Panel
    Dim bRight As Boolean                             '-- Alignment
    Dim xTop As Long

    '-- Clear
    Cls
    m_oPaintDC.Cls BackColor
    bRight = CBool(Me.PanelAlignment = EDPA_Right)

    '-- Set The Embedded control Boundaries
    With m_RectControlBounderies
        .Top = m_lPanelOffset + 1
        .Left = IIf(bRight, 1, m_lPanelOffset)
        .Right = IIf(bRight, (ScaleWidth - m_lPanelOffset) - 1, (ScaleWidth - (1 + m_lPanelOffset)))
        .Bottom = ScaleHeight - (m_lPanelOffset + 1) - 1
    End With


    '-- Set The Position On Where To Draw The Text and Icon
    lFontpos = (m_lPanelOffset / 2) - (m_lFontHeight / 2)

    If Not m_Panels Is Nothing Then

        If m_Panels.Count > 0 Then

            '-- loop Through The Panels And Draw Them
            For Each oPanel In m_Panels

                lindex = lindex + 1

                '-- Process The Panels If The Selected Item Is Not Nothing
                If (Not m_oSelectedItem Is Nothing) Then

                    '-- Process The Selected Item
                    If (oPanel.Key = m_oSelectedItem.Key) Then

                        '-- We Redraw The selected item At the bottom Of this Module

                    Else

                        '-- Process The Hover Item
                        If (Not m_oHoverItem Is Nothing) Then

                            If (oPanel.Key = m_oHoverItem.Key) Then
                                pvDrawMessengerPanel oPanel, yOffset, , True
                            Else
                                pvDrawMessengerPanel oPanel, yOffset
                            End If

                        Else

                            pvDrawMessengerPanel oPanel, yOffset

                        End If
                    End If

                End If

                '-- Set The Panel Rect
                oPanel.RecLeft = IIf(bRight, ScaleWidth - m_lPanelOffset, 0)
                oPanel.RecTop = yOffset
                oPanel.RecRight = IIf(bRight, ScaleWidth, m_lPanelOffset)
                oPanel.RecBottom = yOffset + m_lPanelOffset + IIf(lindex = m_Panels.Count, 25, 8)

                '-- As We Have Drawn The selected Panel
                If (Not m_oSelectedItem Is Nothing) Then

                    If (oPanel.Key = m_oSelectedItem.Key) Then

                        '-- Selected item yoffset
                        xTop = yOffset

                        '-- Increment The Offset
                        yOffset = yOffset + m_lPanelOffset + 10

                        ' m_RectControlBounderies.Bottom = ScaleHeight - m_RectControlBounderies.Top - 1    'yOffset - m_RectControlBounderies.Top

                        '-- Move And Make The Selected Control Visible
                        If Not m_oSelectedItem.EmbededControl Is Nothing Then

                            If m_RectControlBounderies.Bottom < 0 Then

                                m_oSelectedItem.EmbededControl.Enabled = Enabled
                                m_oSelectedItem.EmbededControl.Visible = False

                            Else

                                On Error Resume Next

                                m_oSelectedItem.EmbededControl.Move m_RectControlBounderies.Left * Screen.TwipsPerPixelX, m_RectControlBounderies.Top * Screen.TwipsPerPixelY, m_RectControlBounderies.Right * Screen.TwipsPerPixelX, m_RectControlBounderies.Bottom * Screen.TwipsPerPixelY
                                m_oSelectedItem.EmbededControl.Visible = True
                                m_oSelectedItem.EmbededControl.Enabled = Enabled

                            End If

                        End If

                    Else

                        '-- Increment The Offset
                        yOffset = yOffset + m_lPanelOffset + 7

                    End If

                Else

                    '-- Increment The Offset
                    yOffset = yOffset + m_lPanelOffset + 7

                End If

            Next


            '-- Left/Right Hand Border Line
            m_oPaintDC.DrawLine IIf(Not bRight, m_lPanelOffset - 1, ScaleWidth - m_lPanelOffset), 0, IIf(bRight, ScaleWidth - m_lPanelOffset, m_lPanelOffset - 1), ScaleHeight - 1, m_lColorBorder

            'Redraw The selected item
            pvDrawMessengerPanel m_oSelectedItem, xTop, True

            If Enabled Then
                m_oPaintDC.ForeColor = Me.SelectedItemColor
            Else
                m_oPaintDC.ForeColor = &H80000011
            End If

            '-- Draw The selected item Caption and gradient
            m_oPaintDC.FillGradient IIf(bRight, 1, m_lPanelOffset), 1, IIf(bRight, ScaleWidth - m_lPanelOffset, ScaleWidth), m_lPanelOffset + 1, IIf(Not bRight, vbWhite, m_lColorOneNormal), IIf(Not bRight, m_lColorOneNormal, vbWhite), False
            m_oPaintDC.DrawText m_oSelectedItem.Caption, IIf(bRight, 5, m_lPanelOffset + 10), lFontpos, IIf(bRight, ScaleWidth - m_lPanelOffset - 5, ScaleWidth), m_lPanelOffset, IIf(bRight, DT_RIGHT Or DT_WORD_ELLIPSIS, DT_LEFT Or DT_WORD_ELLIPSIS)
            m_oPaintDC.FillGradient IIf(bRight, 1, m_lPanelOffset), m_lPanelOffset, IIf(bRight, ScaleWidth - m_lPanelOffset, ScaleWidth), m_lPanelOffset + 1, IIf(Not bRight, vbWhite, m_lColorBorder), IIf(Not bRight, m_lColorBorder, vbWhite), False


            '-- Draw the Bottom Line Of The Control
            m_oPaintDC.DrawLine IIf(bRight, 0, m_lPanelOffset - 1), ScaleHeight - 1, IIf(bRight, ScaleWidth - m_lPanelOffset + 1, ScaleWidth - 1), ScaleHeight - 1, m_lColorBorder

            '-- Draw The Other Left Or Right BorderLine Of The Embbeded object
            m_oPaintDC.DrawLine IIf(bRight, 0, ScaleWidth - 1), 0, IIf(bRight, 0, ScaleWidth - 1), ScaleHeight, m_lColorBorder

        Else

            '-- Control Border
            With m_oPaintDC
                .DrawLine 0, 0, ScaleWidth - 1, 0, m_lColorBorder
                .DrawLine 0, ScaleHeight - 1, ScaleWidth, ScaleHeight - 1, m_lColorBorder
                .DrawLine ScaleWidth - 1, 0, ScaleWidth - 1, ScaleHeight, m_lColorBorder
                .DrawLine 0, yOffset, 0, ScaleHeight - 1, m_lColorBorder
            End With

        End If

    Else

        '-- Control Border
        With m_oPaintDC
            .DrawLine 0, 0, ScaleWidth - 1, 0, m_lColorBorder
            .DrawLine 0, ScaleHeight - 1, ScaleWidth, ScaleHeight - 1, m_lColorBorder
            .DrawLine ScaleWidth - 1, 0, ScaleWidth - 1, ScaleHeight, m_lColorBorder
            .DrawLine 0, yOffset, 0, ScaleHeight - 1, m_lColorBorder
        End With

    End If

    '-- draw The control
    m_oPaintDC.BitBlt hdc


End Sub


'----------------------------------
'-- Tooltip
'----------------------------------

Public Property Let ToolTipStyle(ByVal vData As ttToolTipStyleEnum)
    mvarToolTipStyle = vData
End Property

Public Property Get ToolTipStyle() As ttToolTipStyleEnum
    ToolTipStyle = mvarToolTipStyle
End Property


'---------------------------------------------------------------------------------------
' Name      : pvCreateToolTip
' DateTime  : 8 July 2005 17:01
' Author    : Gary Noble
' Purpose   : Create A Windows Tooltip
'---------------------------------------------------------------------------------------
'
Private Function pvCreateToolTip(ByVal ParentHwnd As Long) As Boolean
'Dim lpRect As RECT
    Dim lWinToolTipStyle As Long

    If m_lTTHwnd <> 0 Then
        DestroyWindow m_lTTHwnd
    End If

    m_lParentHwnd = ParentHwnd

    lWinToolTipStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX

    '-- create baloon ToolTipStyle if desired
    If mvarToolTipStyle = TTBalloon Then
        lWinToolTipStyle = lWinToolTipStyle Or TTS_BALLOON
    End If

    m_lTTHwnd = CreateWindowEx(0&, _
                               TOOLTIPS_CLASSA, _
                               vbNullString, _
                               lWinToolTipStyle, _
                               CW_USEDEFAULT, _
                               CW_USEDEFAULT, _
                               CW_USEDEFAULT, _
                               CW_USEDEFAULT, _
                               0&, _
                               0&, _
                               App.hInstance, _
                               0&)

    '-- now set our tooltip info structure
    With ti
        '-- if we want it centered, then set that flag
        If mvarCentered Then
            .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP Or TTF_IDISHWND
        Else
            .lFlags = TTF_SUBCLASS Or TTF_IDISHWND
        End If

        '-- set the hwnd prop to our parent control's hwnd
        .hwnd = m_lParentHwnd
        .lId = m_lParentHwnd                          '0
        .hInstance = App.hInstance
        .lSize = Len(ti)
    End With

    '-- add the tooltip structure
    SendMessage m_lTTHwnd, TTM_ADDTOOLA, 0&, ti

    '-- if we want a ToolTipTitle or we want an icon
    If mvarToolTipTitle <> vbNullString Or mvarIcon <> TTNoIcon Then
        SendMessage m_lTTHwnd, TTM_SETToolTipTitle, CLng(mvarIcon), ByVal mvarToolTipTitle
    End If

    If mvarForeColor <> Empty Then
        SendMessage m_lTTHwnd, TTM_SETTIPTEXTCOLOR, mvarForeColor, 0&
    End If

    If mvarBackColor <> Empty Then
        SendMessage m_lTTHwnd, TTM_SETTIPBKCOLOR, mvarBackColor, 0&
    End If

    SendMessageLong m_lTTHwnd, TTM_SETDELAYTIME, TTDT_AUTOPOP, mvarVisibleTime
    SendMessageLong m_lTTHwnd, TTM_SETDELAYTIME, TTDT_INITIAL, mvarDelayTime
End Function

Private Property Let ToolTipTitle(ByVal vData As String)
    mvarToolTipTitle = vData
    If m_lTTHwnd <> 0 And mvarToolTipTitle <> Empty And mvarIcon <> TTNoIcon Then
        SendMessage m_lTTHwnd, TTM_SETToolTipTitle, CLng(mvarIcon), ByVal mvarToolTipTitle
    End If
End Property

Private Property Get ToolTipTitle() As String
    ToolTipTitle = ti.lpStr
End Property



Private Property Let TipText(ByVal vData As String)
    mvarTipText = vData
    ti.lpStr = vData
    If m_lTTHwnd <> 0 Then
        SendMessage m_lTTHwnd, TTM_UPDATETIPTEXTA, 0&, ti
    End If
End Property

Private Property Get TipText() As String
    TipText = mvarTipText
End Property


Private Sub pvDestroyTooltip()
    If m_lTTHwnd <> 0 Then
        DestroyWindow m_lTTHwnd
        m_lTTHwnd = 0
    End If
End Sub

Private Property Get VisibleTime() As Long
    VisibleTime = mvarVisibleTime
End Property

Private Property Let VisibleTime(ByVal lData As Long)
    mvarVisibleTime = lData
End Property

Private Property Get DelayTime() As Long
    DelayTime = mvarDelayTime
End Property

Private Property Let DelayTime(ByVal lData As Long)
    mvarDelayTime = lData
End Property

