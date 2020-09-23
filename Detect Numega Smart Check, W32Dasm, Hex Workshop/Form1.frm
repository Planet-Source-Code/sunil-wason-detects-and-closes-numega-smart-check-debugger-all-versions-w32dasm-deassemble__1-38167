VERSION 5.00
Begin VB.Form frmCheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anti Crack module"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkFound 
      Caption         =   "Raise system error if Target Application / Debugger / Deassembler detected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   3600
      Width           =   4215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3960
      Top             =   2280
   End
   Begin VB.Label Label5 
      Caption         =   $"Form1.frx":0000
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   4215
   End
   Begin VB.Label Label6 
      Caption         =   $"Form1.frx":0093
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label4 
      Caption         =   "5.  Notepad"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "4.  Winamp"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "3.  Hex Workshop v3.1"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "2.  WinDasm v8.93 (may work for other versions)"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "1.  Numega Smart Check"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "In this sample code, the Applications / Debuggers / Deassemblers considered are:-"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   4215
   End
End
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Name : Detect Numega Smart Check, W32Dasm, Hex Workshop
' Author : Sunil Wason (sunilwason@yahoo.com)
' Purpose : Monitor any debugger/deassembler or
' application etc (whose classname either in
' full or truncated stored in the AppClassName
' array) is running in the windows environment.
' If any of these applications is found, close
' them.

' You may use this code freely and mail me any
' improvements if done. However, kindly donot
' forget to give due credit tothe author if you are using
' this code or its modified version in any
' application.

' Brief Description of the program.
' An array is filled by the class names
' (either in truncated form or full) whose
' applications are required to be monitored.
' We store truncated classnames for applications
' such as WinAmp and Numega Smart Check or any
' other similar applications due to the reasons
' enumerated below.
' "Winamp" is the truncated classname used is
' this code for various Winamp windows & their
' versions. The actual Classnames for all the
' WinAmp windows that run in the foreground or
' background concurrently are:-
' Winamp PE (for Winamp Playlist editor)
' Winamp MB (for Winamp Minibrowser)
' Winamp v1.x (for various vesrions of Winamp)
' Winamp EQ (for Winamp Equalizer)

' Similarly, Numega Smart Check has its actual
' (or full) classnames ranging as NMSCMW0,
' NMSCMW1, NMSCMW2, NMSCMW3, NMSCMW4 and so on.
' I persume they have crossed NMSCMW50.

' In this application, by simply passing
' Winamp or NMSCMW which are considered as
' the truncated classnames, we are able to
' close any Winamp application of
' Numega Smart Check debugger tool.

' Although Winamp & Notepad are not debuggers or
' deassemblers however they have been also added
' only to show that this program can be
' extended to close any application

Private Sub Timer1_Timer()

Dim dummyval As Long
Dim HwndDasm As Long
Dim hwnd As Long
Dim i As Integer

'Fill the class name (truncated or full)
'in an array
FillClassName

'Monitors all the applications running in
'the Windows environment and closes those
'whose classnames have been stored in the
'AppClassName array
For i = 0 To NoOfAppClassMonitored
    hwnd = AppPresent(Trim(AppClassName(i)), frmCheck)
    If hwnd <> 0 Then
        KillWin hwnd
    End If
Next i

End Sub 'Timer1_Timer()

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Unload Me
Set frmCheck = Nothing

End Sub 'Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

