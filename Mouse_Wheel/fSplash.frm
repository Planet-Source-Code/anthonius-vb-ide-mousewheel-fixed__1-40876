VERSION 5.00
Begin VB.Form fSplash 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1140
   ClientLeft      =   2115
   ClientTop       =   2280
   ClientWidth     =   4620
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   Icon            =   "fSplash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   Begin VB.Image img 
      BorderStyle     =   1  'Fixed Single
      Height          =   765
      Left            =   195
      Picture         =   "fSplash.frx":000C
      Top             =   188
      Width           =   825
   End
   Begin VB.Label lblAbout 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Mousewheel Support..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1230
      TabIndex        =   0
      Top             =   450
      Width           =   3255
   End
End
Attribute VB_Name = "fSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'This form has no code

':) Ulli's VB Code Formatter V2.15.4 (10.09.2002 14:17:14) 4 + 0 = 4 Lines
