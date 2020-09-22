VERSION 5.00
Begin VB.Form fSetOpts 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "x"
   ClientHeight    =   3885
   ClientLeft      =   2310
   ClientTop       =   2310
   ClientWidth     =   5955
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   Icon            =   "fSetOpts.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   5955
      TabIndex        =   12
      Top             =   0
      Width           =   5955
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse Wheel Support Add-In"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   345
         TabIndex        =   13
         Top             =   135
         Width           =   4140
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   15
         X2              =   5955
         Y1              =   795
         Y2              =   795
      End
   End
   Begin VB.CommandButton btCAO 
      Caption         =   "&Save"
      Height          =   405
      Index           =   3
      Left            =   345
      TabIndex        =   11
      ToolTipText     =   "Apply and save settings"
      Top             =   3330
      Width           =   915
   End
   Begin VB.CommandButton btCAO 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   420
      Index           =   2
      Left            =   3555
      TabIndex        =   10
      ToolTipText     =   "Apply settings and close box"
      Top             =   3330
      Width           =   915
   End
   Begin VB.CommandButton btCAO 
      Caption         =   "&Apply"
      Height          =   405
      Index           =   1
      Left            =   1395
      TabIndex        =   9
      ToolTipText     =   "Apply settings"
      Top             =   3330
      Width           =   915
   End
   Begin VB.CommandButton btCAO 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   0
      Left            =   4680
      TabIndex        =   8
      Top             =   3330
      Width           =   915
   End
   Begin VB.Frame frScroll 
      Caption         =   "&Lines to scroll"
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   345
      TabIndex        =   0
      Top             =   1005
      Width           =   5250
      Begin VB.OptionButton opPage 
         Caption         =   "&Whole Page"
         CausesValidation=   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2175
         TabIndex        =   3
         Top             =   645
         Width           =   1200
      End
      Begin VB.OptionButton opHalfPage 
         Caption         =   "[set at runtime]"
         CausesValidation=   0   'False
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   4
         Top             =   645
         Width           =   2370
      End
      Begin VB.OptionButton opAbsValue 
         Caption         =   "&Enter number of lines for wheel:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   195
         TabIndex        =   1
         Top             =   300
         Width           =   2700
      End
      Begin VB.TextBox txLines 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   4500
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "3"
         ToolTipText     =   "1 thru 99"
         Top             =   255
         Width           =   480
      End
   End
   Begin VB.Frame frSmooth 
      Caption         =   "S&mooth scrolling"
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   345
      TabIndex        =   5
      Top             =   2280
      Width           =   5250
      Begin VB.OptionButton opOff 
         Caption         =   "&Off"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2085
         TabIndex        =   7
         Top             =   300
         Width           =   630
      End
      Begin VB.OptionButton opOn 
         Caption         =   "&On"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   195
         TabIndex        =   6
         Top             =   300
         Width           =   525
      End
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   5955
      Y1              =   3150
      Y2              =   3150
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   5955
      Y1              =   3165
      Y2              =   3165
   End
End
Attribute VB_Name = "fSetOpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum Idx
    idxCancel = 0
    idxApply = 1
    idxOK = 2
    idxSave = 3
End Enum

Private Sub btCAO_Click(Index As Integer)

    Select Case Index
      Case idxCancel
        Unload Me
      Case idxApply
        Select Case True
          Case opAbsValue
            LinesToScroll = txLines
          Case opPage
            LinesToScroll = "-1"
          Case opHalfPage
            LinesToScroll = "-2"
        End Select
        Smooth = opOn
      Case idxOK
        btCAO_Click idxApply
        btCAO_Click idxCancel
      Case idxSave
        btCAO_Click idxApply
        SaveSetting App.Title, Scroll, Lines, LinesToScroll
        SaveSetting App.Title, Scroll, Mode, IIf(Smooth, sSmooth, sInstant)
    End Select

End Sub

Private Sub Form_Load()

  Const Margin  As Long = 5 'pixels - prevent Me from being placed directly at the screen borders
  Dim MarginX   As Long
  Dim MarginY   As Long

    GetCursorPos CursorPos 'get mouse cursor posn
    With CursorPos
        .x = .x * Screen.TwipsPerPixelX - Width / 2 'adjust to twips and also reflect my dimensions
        .y = .y * Screen.TwipsPerPixelY - Height / 2
        MarginX = Margin * Screen.TwipsPerPixelX
        MarginY = Margin * Screen.TwipsPerPixelY
        Select Case True 'limit x to be within screen
          Case .x < MarginX
            .x = MarginX
          Case .x + Width > Screen.Width - MarginX
            .x = Screen.Width - Width - MarginX
        End Select
        Select Case True 'limit y to be within screen
          Case .y < MarginY
            .y = MarginY
          Case .y + Height > Screen.Height - MarginY
            .y = Screen.Height - Height - MarginY
        End Select
        Move .x, .y 'move Me to that position
    End With 'CURSORPOS

    'preset initial captions and values
    Caption = App.Title
    opHalfPage.Caption = opHpCapt
    opAbsValue = True
    opPage = (LinesToScroll = "-1")
    opHalfPage = (LinesToScroll = "-2")
    opOn = Smooth
    opOff = (Smooth = False)
    If opAbsValue Then
        txLines = LinesToScroll
    End If

End Sub


Private Sub opAbsValue_Click()

    With txLines
        .Enabled = opAbsValue
        .TabStop = opAbsValue
        If opAbsValue Then
            .SelStart = 0
            .SelLength = 2
            On Error Resume Next 'this may be called during form load when we cannot set focus
                .SetFocus
            On Error GoTo 0
        End If
    End With 'TXLINES

End Sub

Private Sub opHalfPage_Click()

    opAbsValue_Click

End Sub

Private Sub opPage_Click()

    opAbsValue_Click

End Sub

Private Sub txLines_KeyPress(KeyAscii As Integer)

    If InStr("0123456789" & Chr$(vbKeyBack), Chr$(KeyAscii)) = 0 Then
        KeyAscii = 0
        Beep
    End If

End Sub

Private Sub txLines_Validate(Cancel As Boolean)

    Cancel = (Val(txLines) = 0)
    If Cancel Then
        Beep
    End If

End Sub
