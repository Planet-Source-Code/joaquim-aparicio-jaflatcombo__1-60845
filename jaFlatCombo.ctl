VERSION 5.00
Begin VB.UserControl jaFlatCombo 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3045
   ScaleHeight     =   87
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   203
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "jaFlatCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Dim m_BorderColor As OLE_COLOR
'Const m_def_BorderColor = &H80000009


Private Enum eMsgWhen
  MSG_after = 1                                                                         'Message calls back after the original (previous) WndProc
  MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
  MSG_BEFORE_AND_AFTER = MSG_after Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset

Private Type tSubData                                                                   'Subclass data type
  hWnd                               As Long                                            'Handle of the window being subclassed
  nAddrSub                           As Long                                            'The address of our new WndProc (allocated memory).
  nAddrOrig                          As Long                                            'The address of the pre-existing WndProc
  nMsgCntA                           As Long                                            'Msg after table entry count
  nMsgCntB                           As Long                                            'Msg before table entry count
  aMsgTblA()                         As Long                                            'Msg after table array
  aMsgTblB()                         As Long                                            'Msg Before table array
End Type

Private sc_aSubData()                As tSubData                                        'Subclass data array

Private Const WM_COMMAND = &H111
Private Const WM_PAINT = &HF
Private Const WM_TIMER = &H113
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const SM_CXHTHUMB = 10

Private Const WM_SETFOCUS = &H7
Private Const WM_KILLFOCUS = &H8
Private Const WM_MOUSEACTIVATE = &H21

Private Type POINTAPI
   X As Long
   Y As Long
End Type
Private Type RECT
   Left     As Long
   Top      As Long
   Right    As Long
   Bottom   As Long
End Type

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Const PS_SOLID = 0
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, lpsz2 As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function RedrawWindow Lib "user32" ( _
   ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Enum EDrawStyle
   FC_DRAWNORMAL = &H1
   FC_DRAWRAISED = &H2
   FC_DRAWPRESSED = &H4
End Enum

Private m_hWnd             As Long
Private m_bMouseOver       As Boolean

Private m_bLBtnDown As Boolean

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'==================================================================================================

'UserControl events

Event Click() 'MappingInfo=Combo1,Combo1,-1,Click
Event DblClick() 'MappingInfo=Combo1,Combo1,-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=Combo1,Combo1,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=Combo1,Combo1,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=Combo1,Combo1,-1,KeyUp
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Default Property Values:
Const m_def_BorderColor = 0
'Property Variables:
Dim m_BorderColor As OLE_COLOR
'Event Declarations:




'Read the properties from the property bag - also, a good place to start the subclassing (if we're running)

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim Index As Integer
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)

    If Ambient.UserMode Then                                                              'If we're not in design mode
        Call Subclass_Start(Combo1.hWnd)                                               'Start subclassing
    
        Call Subclass_AddMsg(Combo1.hWnd, WM_PAINT, MSG_BEFORE_AND_AFTER)
        Call Subclass_AddMsg(Combo1.hWnd, WM_COMMAND, MSG_BEFORE)
        Call Subclass_AddMsg(Combo1.hWnd, WM_MOUSEMOVE, MSG_BEFORE)
        Call Subclass_AddMsg(Combo1.hWnd, WM_TIMER, MSG_BEFORE)
        Call Subclass_AddMsg(Combo1.hWnd, WM_SETFOCUS, MSG_BEFORE)
        Call Subclass_AddMsg(Combo1.hWnd, WM_KILLFOCUS, MSG_BEFORE)
        Call Subclass_AddMsg(Combo1.hWnd, WM_MOUSEHOVER, MSG_BEFORE)
        m_hWnd = Combo1.hWnd
    End If
    Combo1.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Combo1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Combo1.FontName = PropBag.ReadProperty("FontName", "Tahoma")
    Combo1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Combo1.Locked = PropBag.ReadProperty("Locked", False)
    Combo1.SelLength = PropBag.ReadProperty("SelLength", 0)
    Combo1.SelStart = PropBag.ReadProperty("SelStart", 0)
    Combo1.SelText = PropBag.ReadProperty("SelText", "")
    Combo1.Text = PropBag.ReadProperty("Text", "Combo1")
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
End Sub

Private Sub UserControl_Resize()
    Combo1.Move 0, 0, UserControl.ScaleWidth
    UserControl.ScaleHeight = Combo1.Height
End Sub

'The control is terminating - a good place to stop the subclasser
Private Sub UserControl_Terminate()
  On Error GoTo Catch
  'Stop all subclassing
  Call Subclass_StopAll
Catch:
End Sub

'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
Attribute zSubclass_Proc.VB_MemberFlags = "40"
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
Select Case uMsg
   Case WM_COMMAND
      If (m_hWnd = lParam) Then
         ' Type of notification is in the hiword of wParam:
         Select Case wParam \ &H10000
         Case CBN_CLOSEUP
            OnPaint (m_hWnd = GetFocus() Or bDown), bDown
         End Select
         OnTimer False
      End If
      
   Case WM_PAINT
      bDown = DroppedDown()
      bFocus = (m_hWnd = GetFocus() Or bDown)
      OnPaint (bFocus), bDown
      If (bFocus) Then
         OnTimer False
      End If
      
   Case WM_SETFOCUS
      OnPaint True, False
      OnTimer False
      
   Case WM_KILLFOCUS
      OnPaint False, False

   Case WM_MOUSEMOVE
      If Not (m_bMouseOver) Then
         bDown = DroppedDown()
         If Not (m_hWnd = GetFocus() Or bDown) Then
            OnPaint True, False
            m_bMouseOver = True
            ' Start checking to see if mouse is no longer over.
            SetTimer m_hWnd, 1, 10, 0
         End If
      End If
      
   Case WM_TIMER
      OnTimer True
      If Not (m_bMouseOver) Then
         OnPaint False, False
      End If
      
   End Select
End Sub

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines

'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_after)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_after Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_after, .nAddrSub)
    End If
  End With
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_after)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
  'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to be removed from the before, after or both callback tables
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_after Then
      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_after, .nAddrSub)
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
  Const CODE_LEN              As Long = 200                                             'Length of the machine code in bytes
  Const FUNC_CWP              As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM              As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL              As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER              As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5              As String = "vba5"                                        'Location of the EbMode function if running VB5
  Const MOD_VBA6              As String = "vba6"                                        'Location of the EbMode function if running VB6
  Const PATCH_01              As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02              As Long = 68                                              'Address of the previous WndProc
  Const PATCH_03              As Long = 78                                              'Relative address of SetWindowsLong
  Const PATCH_06              As Long = 116                                             'Address of the previous WndProc
  Const PATCH_07              As Long = 121                                             'Relative address of CallWindowProc
  Const PATCH_0A              As Long = 186                                             'Address of the owner object
  Static aBuf(1 To CODE_LEN)  As Byte                                                   'Static code buffer byte array
  Static pCWP                 As Long                                                   'Address of the CallWindowsProc
  Static pEbMode              As Long                                                   'Address of the EbMode IDE break/stop/running function
  Static pSWL                 As Long                                                   'Address of the SetWindowsLong function
  Dim i                       As Long                                                   'Loop index
  Dim j                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sHex                    As String                                                 'Hex code string
  
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
      aBuf(j) = Val("&H" & Mid$(sHex, i, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      i = i + 2
    Loop                                                                                'Next pair of hex characters
    
'Get API function addresses
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
      aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
      If pEbMode = 0 Then                                                               'Found?
        pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                         'VB5 perhaps
      End If
    End If
    
    pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
    pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lng_hWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Subclass_Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .hWnd = lng_hWnd                                                                    'Store the hWnd
    .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                              'Copy the machine code from the static byte array to the code array in sc_aSubData
    Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                           'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                           'Patch the relative address of the CallWindowProc api function
    Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                                     'Patch the address of this object instance into the static machine code buffer
  End With
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()
  Dim i As Long
  
  i = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While i >= 0                                                                       'Iterate through each element
    With sc_aSubData(i)
      If .hWnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        Call Subclass_Stop(.hWnd)                                                       'Subclass_Stop
      End If
    End With
    
    i = i - 1                                                                           'Next element
  Loop
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
'Parameters:
  'lng_hWnd  - The handle of the window to stop being subclassed
  With sc_aSubData(zIdx(lng_hWnd))
    Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hWnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
End Sub

'======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
    nMsgCnt = 0                                                                         'Message count is now zero
    If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
      nEntry = PATCH_05                                                                 'Patch the before table message count location
    Else                                                                                'Else after
      nEntry = PATCH_09                                                                 'Patch the after table message count location
    End If
    Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
  Else                                                                                  'Else deleteting a specific message
    Do While nEntry < nMsgCnt                                                           'For each table entry
      nEntry = nEntry + 1
      If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
        aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
        Exit Do                                                                         'Bail
      End If
    Loop                                                                                'Next entry
  End If
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hWnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
        If Not bAdd Then                                                                'If we're searching not adding
          Exit Function                                                                 'Found
        End If
      ElseIf .hWnd = 0 Then                                                             'If this an element marked for reuse.
        If bAdd Then                                                                    'If we're adding
          Exit Function                                                                 'Re-use it
        End If
      End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
  Loop
  
  If Not bAdd Then
    Debug.Assert False                                                                  'hWnd not found, programmer error
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

Private Sub Draw(ByVal dwStyle As EDrawStyle, clrTopLeft As OLE_COLOR, clrBottomRight As OLE_COLOR)
   If m_hWnd = 0 Then Exit Sub
    DrawCombo dwStyle, clrTopLeft, clrBottomRight
End Sub


Private Function Draw3DRect( _
      ByVal hdc As Long, _
      ByRef rcItem As RECT, _
      ByVal oTopLeftColor As OLE_COLOR, _
      ByVal oBottomRightColor As OLE_COLOR _
   )
Dim hPen As Long
Dim hPenOld As Long
Dim tP As POINTAPI
   hPen = CreatePen(PS_SOLID, 1, TranslateColor(oTopLeftColor))
   hPenOld = SelectObject(hdc, hPen)
   MoveToEx hdc, rcItem.Left, rcItem.Bottom - 1, tP
   LineTo hdc, rcItem.Left, rcItem.Top
   LineTo hdc, rcItem.Right - 1, rcItem.Top
   SelectObject hdc, hPenOld
   DeleteObject hPen
   If (rcItem.Left <> rcItem.Right) Then
      hPen = CreatePen(PS_SOLID, 1, TranslateColor(oBottomRightColor))
      hPenOld = SelectObject(hdc, hPen)
      LineTo hdc, rcItem.Right - 1, rcItem.Bottom - 1
      LineTo hdc, rcItem.Left, rcItem.Bottom - 1
      SelectObject hdc, hPenOld
      DeleteObject hPen
   End If
End Function

Private Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If
End Function

Private Sub DrawCombo(ByVal dwStyle As EDrawStyle, clrTopLeft As OLE_COLOR, clrBottomRight As OLE_COLOR)
    Dim rcItem As RECT
    Dim rcWork As RECT
    Dim rcButton As RECT
    Dim hdc As Long
    Dim hWndFocus As Long
    Dim tP As POINTAPI
    Dim hBr As Long
    Dim bRightToLeft As Long
   
   GetClientRect m_hWnd, rcItem
   hdc = GetDC(m_hWnd)
   
      Draw3DRect hdc, rcItem, clrTopLeft, clrBottomRight
      InflateRect rcItem, -1, -1
      Draw3DRect hdc, rcItem, vbWindowBackground, vbWindowBackground
       
       LSet rcButton = rcItem
        rcButton.Left = rcButton.Right - GetSystemMetrics(SM_CXHTHUMB) - 2
       
       If (dwStyle = FC_DRAWNORMAL) And (clrTopLeft <> vbHighlight) Then
          hBr = CreateSolidBrush(TranslateColor(vbButtonFace))
       ElseIf (dwStyle = FC_DRAWPRESSED) Then
          hBr = CreateSolidBrush(VSNetPressedColor)
       Else
          hBr = CreateSolidBrush(VSNetSelectionColor)
       End If
       FillRect hdc, rcButton, hBr
       DeleteObject hBr
       
       LSet rcWork = rcButton
        rcWork.Left = rcButton.Left
        rcWork.Right = rcWork.Left
       If (dwStyle = FC_DRAWNORMAL) And (clrTopLeft <> vbHighlight) Then
          Draw3DRect hdc, rcWork, vbWindowBackground, vbWindowBackground
       Else
          Draw3DRect hdc, rcWork, vbHighlight, vbHighlight
       End If
       If (bRightToLeft) Then
          rcWork.Right = rcWork.Right + 1
          rcWork.Left = rcWork.Right
       Else
          rcWork.Left = rcWork.Left - 1
          rcWork.Right = rcWork.Left
       End If
       DrawComboDropDownGlyph hdc, rcButton, vbWindowText
    '   Draw3DRect hdc, rcWork, vbWindowBackground, vbWindowBackground
   
   ReleaseDC m_hWnd, hdc

End Sub

Private Sub DrawComboDropDownGlyph( _
      ByVal hdc As Long, _
      rcButton As RECT, _
      ByVal oColor As OLE_COLOR _
   )
Dim hPen As Long
Dim hPenOld As Long
Dim xC As Long
Dim yC As Long
Dim tJ As POINTAPI
   
   xC = rcButton.Left + (rcButton.Right - rcButton.Left) \ 2 + 1
   yC = rcButton.Top + (rcButton.Bottom - rcButton.Top) \ 2
   
   hPen = CreatePen(PS_SOLID, 1, TranslateColor(oColor))
   hPenOld = SelectObject(hdc, hPen)
   MoveToEx hdc, xC - 3, yC - 2, tJ
   LineTo hdc, xC + 4, yC - 2
   MoveToEx hdc, xC - 2, yC - 1, tJ
   LineTo hdc, xC + 3, yC - 1
   MoveToEx hdc, xC - 1, yC, tJ
   LineTo hdc, xC + 2, yC
   MoveToEx hdc, xC, yC - 1, tJ
   LineTo hdc, xC, yC + 2
   
   SelectObject hdc, hPenOld
   DeleteObject hPen
   
End Sub
'
Public Property Get DroppedDown() As Boolean
      DroppedDown = (SendMessageLong(m_hWnd, CB_GETDROPPEDSTATE, 0, 0) <> 0)
End Property

Private Sub OnPaint(ByVal bFocus As Boolean, ByVal bDropped As Boolean)
 'used for paint
   If bFocus Then
      Dim clrTopLeft As Long
      Dim clrBottomRight As Long
      clrTopLeft = vbHighlight
      clrBottomRight = vbHighlight
      
      If (bDropped) Then
         Draw FC_DRAWPRESSED, clrTopLeft, clrBottomRight
      Else
         Draw FC_DRAWRAISED, clrTopLeft, clrBottomRight
      End If
   Else
         Draw FC_DRAWNORMAL, BorderColor, BorderColor
   End If
   
End Sub


Private Sub OnTimer(ByVal bCheckMouse As Boolean)
Dim bOver As Boolean
Dim rcItem As RECT
Dim tP As POINTAPI
   
   If (bCheckMouse) Then
      bOver = True
      GetCursorPos tP
      GetWindowRect m_hWnd, rcItem
      If (PtInRect(rcItem, tP.X, tP.Y) = 0) Then
         bOver = False
      End If
   End If
   
   If Not (bOver) Then
      KillTimer m_hWnd, 1
      m_bMouseOver = False
   End If

End Sub


Private Property Get BlendColor(ByVal oColorFrom As OLE_COLOR, ByVal oColorTo As OLE_COLOR, Optional ByVal alpha As Long = 128) As Long
    Dim lCFrom As Long
    Dim lCTo As Long
    lCFrom = TranslateColor(oColorFrom)
    lCTo = TranslateColor(oColorTo)
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
     
   
   BlendColor = RGB( _
      ((lSrcR * alpha) / 255) + ((lDstR * (255 - alpha)) / 255), _
      ((lSrcG * alpha) / 255) + ((lDstG * (255 - alpha)) / 255), _
      ((lSrcB * alpha) / 255) + ((lDstB * (255 - alpha)) / 255) _
      )
      
End Property

Private Property Get VSNetControlColor() As Long
   VSNetControlColor = BlendColor(vbButtonFace, VSNetBackgroundColor, 195)
End Property

Private Property Get VSNetBackgroundColor() As Long
   VSNetBackgroundColor = BlendColor(vbWindowBackground, vbButtonFace, 220)
End Property
Private Property Get VSNetCheckedColor() As Long
   VSNetCheckedColor = BlendColor(vbHighlight, vbWindowBackground, 30)
End Property
Private Property Get VSNetBorderColor() As Long
   VSNetBorderColor = TranslateColor(vbHighlight)
End Property
Private Property Get VSNetSelectionColor() As Long
   VSNetSelectionColor = BlendColor(vbHighlight, vbWindowBackground, 70)
End Property
Private Property Get VSNetPressedColor() As Long
   VSNetPressedColor = BlendColor(vbHighlight, VSNetSelectionColor, 70)
End Property


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim Index As Integer
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("Enabled", Combo1.Enabled, True)
    Call PropBag.WriteProperty("Font", Combo1.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontName", Combo1.FontName, "")
    Call PropBag.WriteProperty("ForeColor", Combo1.ForeColor, &H80000008)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
'    Call PropBag.WriteProperty("ItemData" & Index, Combo1.ItemData(Index), 0)
'    Call PropBag.WriteProperty("ListIndex", Combo1.ListIndex, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
'    Call PropBag.WriteProperty("List" & Index, Combo1.List(Index), "")
    Call PropBag.WriteProperty("Locked", Combo1.Locked, False)
    Call PropBag.WriteProperty("SelLength", Combo1.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", Combo1.SelStart, 0)
    Call PropBag.WriteProperty("SelText", Combo1.SelText, "")
    Call PropBag.WriteProperty("Text", Combo1.Text, "Combo1")
    Call PropBag.WriteProperty("DroppedDown", m_DroppedDown, m_def_DroppedDown)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,AddItem
Public Sub AddItem(ByVal Item As String, Optional Index As Variant)
Attribute AddItem.VB_Description = "Adds an item to a Listbox or ComboBox control or a row to a Grid control."
    If Not IsMissing(Index) Then
        Combo1.AddItem Item, Index
    Else
        Combo1.AddItem Item
    End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Clear
Public Sub Clear()
Attribute Clear.VB_Description = "Clears the contents of a control or the system Clipboard."
    Combo1.Clear
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = Combo1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Combo1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Combo1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Combo1.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = Combo1.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    Combo1.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Combo1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Combo1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,ItemData
Public Property Get ItemData(ByVal Index As Integer) As Long
Attribute ItemData.VB_Description = "Returns/sets a specific number for each item in a ComboBox or ListBox control."
    ItemData = Combo1.ItemData(Index)
End Property

Public Property Let ItemData(ByVal Index As Integer, ByVal New_ItemData As Long)
    Combo1.ItemData(Index) = New_ItemData
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,ListIndex
Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
    ListIndex = Combo1.ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    Combo1.ListIndex() = New_ListIndex
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,ListCount
Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control."
    ListCount = Combo1.ListCount
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,List
Public Property Get List(ByVal Index As Integer) As String
Attribute List.VB_Description = "Returns/sets the items contained in a control's list portion."
    List = Combo1.List(Index)
End Property

Public Property Let List(ByVal Index As Integer, ByVal New_List As String)
    Combo1.List(Index) = New_List
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = Combo1.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    Combo1.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,NewIndex
Public Property Get NewIndex() As Integer
Attribute NewIndex.VB_Description = "Returns the index of the item most recently added to a control."
    NewIndex = Combo1.NewIndex
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    Combo1.Refresh
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
    SelLength = Combo1.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    Combo1.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    SelStart = Combo1.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    Combo1.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
    SelText = Combo1.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    Combo1.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = Combo1.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    Combo1.Text() = New_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BorderColor = m_def_BorderColor
End Sub

