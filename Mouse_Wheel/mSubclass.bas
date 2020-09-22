Attribute VB_Name = "mSubclass"
Option Explicit

Public VBInstance                   As vbide.VBE 'this has the instantiated application object

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const IDX_WINDOWPROC  As Long = -4

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Const PropName              As String = "Hooked"

Public Declare Function GetCursorPos Lib "user32" (lpPoint As Point) As Long
Private Type Point
    x As Long
    y As Long
End Type
Public CursorPos                    As Point

Private Const WM_KILLFOCUS          As Long = 8
Private Const WM_MOUSEWHEEL         As Long = &H20A

Private hWndActiveCodePane          As Long
Private CodePaneOriginalProcPtr     As Long
Private oldHwndCodePane             As Long

'Send Mail
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL         As Long = 1
Private Const SE_NO_ERROR           As Long = 33 'Values below 33 are error returns

'Registry
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType, lpData As Any, lpcbData As Long) As Long
Private Declare Sub RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long)
Private Const KEY_QUERY_VALUE       As Long = 1
Private Const REG_OPTION_RESERVED   As Long = 0
Private Const ERROR_NONE            As Long = 0

'Registry scroll setting access keys...
Private Const HKEY_CURRENT_USER     As Long = &H80000001
Private Const DesktopSettings       As String = "Control Panel\Desktop"
Private Const SmoothScroll          As String = "SmoothScroll"
Private Const WheelScrollLines      As String = "WheelScrollLines"

'...and our own settings...
Public Const Scroll                 As String = "Scroll"
Public Const Lines                  As String = "Lines"
Public Const Mode                   As String = "Mode"
Public Const sSmooth                As String = "Smooth"
Public Const sInstant               As String = "Instant"

'..and finally what we got (or didn't get) from the Registry or from our own Options
Public LinesToScroll                As String
Public Smooth                       As Long

'- - - - - - - - - - - - - - - - - - - - - - - - modify both values to correspond- - - - -
Public Const opHpCapt               As String = "Half a &Page"
Private Const ScrollFraction        As Single = 1 / 2 'fraction of page to scroll
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private Function CodePaneProc(ByVal hwnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
  'Window procedure for all IDE MDI Child Windows
  'Intercepts all messages to code pane windows

  Dim TopLn     As Long
  Dim NumLines  As Long
  Dim ScrollTo  As Long

    CodePaneOriginalProcPtr = GetProp(hwnd, PropName) 'get original winproc pointer from this window's property
    If CodePaneOriginalProcPtr Then 'if we got one then this window is subclassed:
        CodePaneProc = CallWindowProc(CodePaneOriginalProcPtr, hwnd, nMsg, wParam, lParam) 'call the original winproc to do what has to be done
        Select Case nMsg 'and now split on message type
          Case WM_KILLFOCUS 'this window just lost the focus (remember - the original procedure has already been performed)
            RemoveProp hwnd, PropName 'so remove the property
            SetWindowLong hwnd, IDX_WINDOWPROC, CodePaneOriginalProcPtr 'and re-install the original winproc pointer
            HookActiveCodePane 'and finally hook the code pane which is now active
          Case WM_MOUSEWHEEL 'hah! there it is, that's what it's all about - the user fingers the mouse wheel
            If wParam And 1 Then 'left mouse button is down while scrolling
                fSetOpts.Show vbModal 'show options dialog
                Unload fSetOpts
              Else 'NOT WPARAM...
                If Not VBInstance.ActiveCodePane Is Nothing Then 'bug fix - we have a codepane to scroll in
                    With VBInstance.ActiveCodePane
                        'translate mousewheel and pressed key (Shift or Cntl)
                        TopLn = .TopLine
                        Select Case LinesToScroll
                          Case "-2"
                            NumLines = .CountOfVisibleLines * ScrollFraction
                          Case "-1"
                            NumLines = .CountOfVisibleLines - 1 'so that the bottom line is at the top after scrolling
                          Case Else
                            NumLines = Abs(Val(LinesToScroll))
                            If NumLines >= .CountOfVisibleLines Then 'not more than a page
                                NumLines = .CountOfVisibleLines - 1
                            End If
                        End Select
                        If NumLines < 1 Then 'at least one line
                            NumLines = 1
                        End If
                        ScrollTo = TopLn - Sgn(wParam) * NumLines / ((wParam And &HFFFF&) \ 4 + 1) 'compute new top line
                        If ScrollTo = TopLn Then
                            ScrollTo = TopLn - Sgn(wParam)
                        End If
                        With .CodeModule
                            Select Case ScrollTo 'correct it if it is out of range
                              Case Is < 1
                                ScrollTo = 1
                              Case Is > .CountOfLines
                                ScrollTo = .CountOfLines
                            End Select
                        End With '.CODEMODULE
                        If Smooth Then
                            Do
                                TopLn = TopLn + Sgn(ScrollTo - TopLn)
                                Sleep 1
                                .TopLine = TopLn
                            Loop Until TopLn = ScrollTo
                          Else 'SMOOTH = FALSE/0
                            .TopLine = ScrollTo
                        End If
                    End With 'VBINSTANCE.ACTIVECODEPANE
                End If
            End If
        End Select
    End If

End Function

Public Sub GetScrollSettings()

  Dim RegHandle     As Long
  Dim DataType      As Long
  Dim DataLength    As Long

    If RegOpenKeyEx(HKEY_CURRENT_USER, DesktopSettings, REG_OPTION_RESERVED, KEY_QUERY_VALUE, RegHandle) = ERROR_NONE Then
        LinesToScroll = String$(4, 0)
        DataLength = Len(LinesToScroll)
        If RegQueryValueEx(RegHandle, WheelScrollLines, REG_OPTION_RESERVED, DataType, ByVal LinesToScroll, DataLength) = ERROR_NONE Then
            LinesToScroll = Left$(LinesToScroll, DataLength + (Asc(Mid$(LinesToScroll, DataLength, 1)) = 0))
            If Not IsNumeric(LinesToScroll) Then 'default
                LinesToScroll = "-2"
            End If
          Else 'default'NOT REGQUERYVALUEEX(REGHANDLE,...
            LinesToScroll = "-2"
        End If
        DataLength = Len(Smooth)
        If RegQueryValueEx(RegHandle, SmoothScroll, REG_OPTION_RESERVED, DataType, Smooth, DataLength) = ERROR_NONE Then
            Smooth = CBool(Smooth)
          Else 'default'NOT REGQUERYVALUEEX(REGHANDLE,...
            Smooth = True
        End If
        RegCloseKey RegHandle
      Else 'default'NOT REGOPENKEYEX(HKEY_CURRENT_USER,...
        LinesToScroll = "-2"
        Smooth = True
    End If
    LinesToScroll = GetSetting(App.Title, Scroll, Lines, LinesToScroll)
    Smooth = (GetSetting(App.Title, Scroll, Mode, IIf(Smooth, sSmooth, sInstant)) = sSmooth)

End Sub

Public Sub HookActiveCodePane()

    'hWndActiveCodePane = FindWindowEx(VBInstance.MainWindow.hWnd, 0, "MDIClient", vbNullString) 'find topmost (active) child window of class "MDIClient" in VB's main MDI window
    If oldHwndCodePane <> hWndActiveCodePane Then
        UnhookPreviousCodePane
    End If
    
    hWndActiveCodePane = FindWindow("VbaWindow", vbNullString)
    If hWndActiveCodePane = 0 Then
        hWndActiveCodePane = FindWindowEx(VBInstance.MainWindow.hwnd, 0, "MDIClient", vbNullString) 'find topmost (active) child window of class "MDIClient" in VB's main MDI window
    End If
    If hWndActiveCodePane <> 0 And oldHwndCodePane <> hWndActiveCodePane Then 'found one - should be a code pane window
        SetProp hWndActiveCodePane, PropName, GetWindowLong(hWndActiveCodePane, IDX_WINDOWPROC) 'store the winproc pointer of this window in a property
        SetWindowLong hWndActiveCodePane, IDX_WINDOWPROC, AddressOf CodePaneProc 'and now point to our CodePaneProc so that we can see the messages arriving at this code pane
        oldHwndCodePane = hWndActiveCodePane
    End If

End Sub

Public Sub UnhookPreviousCodePane()

    CodePaneOriginalProcPtr = GetProp(oldHwndCodePane, PropName) 'get the original code pane winproc pointer from property
    If CodePaneOriginalProcPtr Then 'if there is one then we unhook this window:
        RemoveProp oldHwndCodePane, PropName 'remove the property
        SetWindowLong oldHwndCodePane, IDX_WINDOWPROC, CodePaneOriginalProcPtr 'and restore the original winproc pointer
    End If

End Sub

Public Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
On Error Resume Next
    HookActiveCodePane
End Sub
