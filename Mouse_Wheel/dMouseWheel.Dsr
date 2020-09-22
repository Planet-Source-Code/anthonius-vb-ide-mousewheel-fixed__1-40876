VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} dMouseWheel 
   ClientHeight    =   6795
   ClientLeft      =   1800
   ClientTop       =   1935
   ClientWidth     =   11250
   _ExtentX        =   19844
   _ExtentY        =   11986
   _Version        =   393216
   Description     =   "Mousewheel support"
   DisplayName     =   "SOFTPAE VBIDE MouseWheel"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "dMouseWheel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'© 2002     UMGEDV GmbH  (umgedv@aol.com)
'
'Author     UMG (Ulli K. Muehlenweg)
'
'Title      VB6 IDE Mouse Wheel Support
'
'           Adds mouse wheel support to the VB IDE Code Panes
'           Simply compile the .DLL into your VB folder
'
'**********************************************************************************
'Development History
'**********************************************************************************
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'08Sep2002 Version 1.2.4      UMG
'
'Get scroll options from Registry, or form our own settings; compile time option
'no longer exists.
'
'Hold down left mouse button and rotate wheel to open the Scroll Settings Dialog Box.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'06Sep2002 Version 1.1.5      UMG
'
'"Exit" bug fixed - wasn't an Exit bug really: this happened when the user tried to
'scroll AND all codepanes were closed AND at least one codepane had been open before.
'       ===                           ===
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'05Sep2002 Version 1.1.4      UMG
'
'Scrolling method changed - no sending keystrokes anymore
'
'You can now slow down scrolling by factors 2, 3 or 4 by holding down the Shift key,
'the Cntl key or both respectively, while scrolling the mouse wheel.
'
'You also have the choice between two alternative scrolling modes at compile time
'by altering a single #Const
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'05Sep2002 Version 1.0.1      UMG
'
'Now has a "fraction of page to scroll" - constant, currently set to 1/2, modify that
'as you like. If you feel like storing/getting this value from/in Settings: the only
'limit is your imagination.
'
'A little code cosmetic and plenty of comments
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'04Sep2002 Version 1.0.0      UMG
'
'Prototype
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private WithEvents ProjEvent    As VBProjectsEvents
Attribute ProjEvent.VB_VarHelpID = -1
Private WithEvents CompEvent    As VBComponentsEvents
Attribute CompEvent.VB_VarHelpID = -1

Private Sub AddinInstance_OnBeginShutdown(custom() As Variant)

    UnhookPreviousCodePane 'VB's gonna go now so unhook any hooked window

End Sub

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)

    DoEvents
    Set VBInstance = Application 'save object variable pointing to the application instance
    With VBInstance
        Set ProjEvent = .Events.VBProjectsEvents 'ensure that we're kept up to date about project events
        ProjEvent_ItemActivated .ActiveVBProject 'and components events of this project also (VB should fire this event initially also, but doesn't - so we fake it)
    End With 'VBINSTANCE
    GetScrollSettings 'from registry and get our own settings too if there are any
    
    Load frmWnd: frmWnd.Hide
    SetTimer frmWnd.hwnd, 0, 1, AddressOf TimerProc

End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)

    UnhookPreviousCodePane 'the user is unloading the addin so unhook any hooked window
    KillTimer frmWnd.hwnd, 0
    Unload frmWnd

End Sub

Private Sub CompEvent_ItemActivated(ByVal VBComponent As vbide.VBComponent)
    UnhookPreviousCodePane 'so unhook the previous code panel (this does nothing if there is no previous hook)
    HookActiveCodePane 'and hook the newly selected one
End Sub

Private Sub CompEvent_ItemSelected(ByVal VBComponent As vbide.VBComponent)

  'A new component was selected

    UnhookPreviousCodePane 'so unhook the previous code panel (this does nothing if there is no previous hook)
    HookActiveCodePane 'and hook the newly selected one

End Sub

Private Sub ProjEvent_ItemActivated(ByVal VBProject As vbide.VBProject)

  'The one and only project (or a new project in a multi-project app) was selected

    Set CompEvent = VBInstance.Events.VBComponentsEvents(VBProject) 'so point to the components events of this newly selected project

End Sub
