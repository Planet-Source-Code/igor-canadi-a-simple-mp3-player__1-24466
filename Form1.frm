VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mp3 Player"
   ClientHeight    =   5730
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   7800
   FillStyle       =   0  'Solid
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      BackColor       =   &H80000012&
      Caption         =   "<<<"
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000012&
      Caption         =   ">>>"
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000012&
      Caption         =   "To begining"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   4080
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   600
      Top             =   3960
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000012&
      Caption         =   "Exit"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000012&
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   3600
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000012&
      Caption         =   "Pause "
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   3120
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000012&
      Caption         =   "Play"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   2640
      Width           =   3255
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H80000012&
      ForeColor       =   &H80000018&
      Height          =   4185
      Left            =   120
      Pattern         =   "*.mp3"
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
   Begin MSComctlLib.Slider slider 
      Height          =   315
      Left            =   4320
      TabIndex        =   7
      Top             =   2040
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   315
      Left            =   240
      TabIndex        =   13
      Top             =   5280
      Visible         =   0   'False
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   1
      Max             =   5000
      TickStyle       =   3
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      Caption         =   "Volume:"
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   4920
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   1440
      Width           =   3135
   End
   Begin MediaPlayerCtl.MediaPlayer Mp1 
      Height          =   855
      Left            =   2760
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   2295
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   0
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   0   'False
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   0   'False
      SendWarningEvents=   0   'False
      SendErrorEvents =   0   'False
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   0   'False
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      ForeColor       =   &H80000018&
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.Menu sa 
      Caption         =   "&Mp3 Player"
      Begin VB.Menu volumen 
         Caption         =   "&Volume"
         Shortcut        =   ^G
      End
      Begin VB.Menu ads 
         Caption         =   "&Volume Control"
      End
      Begin VB.Menu oprog 
         Caption         =   "&About"
      End
      Begin VB.Menu izlaz 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ads_Click()
On Error GoTo errorhandler
       Dim lngresult As Long
       lngresult = Shell("c:\windows\Sndvol32.exe", vbNormalFocus)
       Exit Sub
errorhandler:
    lngresult = Shell("c:\winnt\system32\Sndvol32.exe", vbNormalFocus)
End Sub

Private Sub Command1_Click()
On Error GoTo Err
slider.Max = Mp1.Duration
Mp1.Play
Command2.Enabled = True
Command3.Enabled = True
Command6.Enabled = True
Command1.Enabled = False
Err:
If Err.Number = -2147467259 Then
MsgBox "You didn't select a file, or you are stupid!", vbExclamation, "Error"
End If
End Sub

Private Sub Command2_Click()
On Error GoTo Err
Mp1.Pause
Command2.Enabled = True
Command3.Enabled = True
Command1.Enabled = True
Err:
If Err.Number = -2147467259 Then
MsgBox "You didn't select a file, or you are stupid!", vbExclamation, "Error"
End If
End Sub

Private Sub Command3_Click()
On Error GoTo Err
Mp1.Stop
slider.Value = 0
Mp1.CurrentPosition = 0
Command2.Enabled = False
Command3.Enabled = False
Command1.Enabled = True
Command6.Enabled = False
Err:
If Err.Number = -2147467259 Then
MsgBox "You didn't select a file, or speaker is busy!", vbExclamation, "Error"
End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
On Error GoTo hell
slider.Value = slider.Value + 5
Mp1.CurrentPosition = Mp1.CurrentPosition + 5
hell:
If Err.Number = 380 Then MsgBox "You can't go further!", vbExclamation, "error"
End Sub

Private Sub Command6_Click()
slider.Value = 0
Mp1.CurrentPosition = 0
Command2.Enabled = True
Command3.Enabled = True
Command6.Enabled = True
End Sub

Private Sub Command7_Click()
On Error GoTo hell
slider.Value = slider.Value - 5
Mp1.CurrentPosition = Mp1.CurrentPosition - 5
hell:
If Err.Number = 380 Then MsgBox "You can't go further!", vbExclamation, "error"
End Sub

Private Sub File1_Click()
Dim Path2
Label1.Caption = File1.FileName
Path2 = File1.Path
If Not Right(Path2, 1) = "\" Then
Path2 = Path2 & "\"
End If
Path2 = Path2 & File1.FileName
Mp1.FileName = Path2
Label4.Caption = Int(Mp1.Duration) & " sec durration"
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command6.Enabled = False
slider.Max = Mp1.Duration
End Sub

Private Sub File1_DblClick()
Dim Path2
Label1.Caption = File1.FileName
Path2 = File1.Path
If Not Right(Path2, 1) = "\" Then
Path2 = Path2 & "\"
End If
Path2 = Path2 & File1.FileName
Mp1.FileName = Path2
Label4.Caption = Int(Mp1.Duration) & " sec durration"
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = True
Command6.Enabled = True
Mp1.Play
slider.Max = Mp1.Duration
End Sub

Private Sub Slider3_Click()
sha = Slider3.Value - 5000
Mp1.Volume = sha
poo = Slider3.Min
foo = Slider3.Value
End Sub

Private Sub Slider3_Scroll()
sha = Slider3.Value - 5000
Mp1.Volume = sha
poo = Slider3.Min
foo = Slider3.Value
End Sub
Private Sub Form_Load()
Dim pim, sha
Dim foo As Integer, poo As Integer
Slider3.Value = Form1.Mp1.Volume + 5000
Dim d, mi, X, Y As Integer
Y = 0
X = 60
File1.Path = App.Path
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MsgBox("Do you realy want to exit?", vbQuestion + vbYesNo, "Exit") = vbYes Then
End
Else:
Cancel = True
End If
End Sub

Private Sub slider_Scroll()
Mp1.CurrentPosition = slider.Value
End Sub

Private Sub izlaz_Click()
End
End Sub

Private Sub oprog_Click()
Form3.Show
End Sub


Private Sub Timer1_Timer()
slider.Value = Mp1.CurrentPosition
d = Mp1.CurrentPosition
Label2.Caption = Int(d) & " sec played"
Label3.Caption = (Int(Mp1.Duration) - Int(d)) & " sec left to play"
If Int(d) = Int(Mp1.Duration) Then
If Mp1.Duration = 0 Then
Exit Sub
Else
Command2.Enabled = False
Command3.Enabled = False
Command1.Enabled = True
Command6.Enabled = False
End If
End If
End Sub

Private Sub volumen_Click()
If Slider3.Visible = False Then
Slider3.Visible = True
Label5.Visible = True
Else
If Slider3.Visible = True Then
Slider3.Visible = False
Label5.Visible = False
End If
End If
End Sub
