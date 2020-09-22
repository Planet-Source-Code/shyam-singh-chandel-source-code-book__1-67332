VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H0000C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2655
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000C0C0&
      Caption         =   "Cancel"
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C0C0&
      Caption         =   "Login"
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      Height          =   2655
      Left            =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim User As String, Pass As String, ina As String, ifa As String

Private Sub Command1_Click()
If Text1.Text = User And Text2.Text = Pass Then
Form1.Show
Unload Me
Else
MsgBox "Wrong User IDs"
Text1 = ""
Text2 = ""
Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
User = GetSetting("Windows", "Info", "WinDelta")
Pass = GetSetting("Windows", "Info", "WinDeltaHook")
If User = "" Then
ina = InputBox("Enter User Name")
Call SaveSetting("Windows", "Info", "WinDelta", ina)
If Pass = "" Then
ifa = InputBox("Enter Password")
Call SaveSetting("Windows", "Info", "WinDeltaHook", ifa)
User = GetSetting("Windows", "Info", "WinDelta")
Pass = GetSetting("Windows", "Info", "WinDeltaHook")
End If
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text2.SetFocus
End If
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Command1_Click
End If
End Sub

