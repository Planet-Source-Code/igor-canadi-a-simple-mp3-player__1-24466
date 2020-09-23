VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6390
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   4560
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   840
      Picture         =   "Form3.frx":1CFA
      ScaleHeight     =   3615
      ScaleWidth      =   4815
      TabIndex        =   2
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Napomena: Glasnoæa se može mjenjati samo mp3 fileovima."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1320
      TabIndex        =   6
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "This program is freeware."
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   840
      MouseIcon       =   "Form3.frx":B626
      TabIndex        =   5
      Top             =   5520
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "Mail to: Petar Palasek"
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   840
      MouseIcon       =   "Form3.frx":B930
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Mail to: Igor Èanadi "
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   840
      MouseIcon       =   "Form3.frx":BC3A
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Authors : Igor Èanadi i Petar Palašek"
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   4440
      Width           =   3015
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
Form3.Hide
End Sub

Private Sub Label2_Click()
Shell ("start mailto:igor_canadi@hotmail.com")
End Sub

Private Sub Label3_Click()
Shell ("start mailto:pero_palasek@hotmail.com")
End Sub

