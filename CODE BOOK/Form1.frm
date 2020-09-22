VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000C0C0&
   Caption         =   "SOURCE CODES"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   24
      Top             =   3840
      Width           =   3495
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   23
      Top             =   6720
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7080
      TabIndex        =   0
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7080
      TabIndex        =   1
      Top             =   2280
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7080
      TabIndex        =   2
      Top             =   2760
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11640
      TabIndex        =   3
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11640
      TabIndex        =   4
      Top             =   2280
      Width           =   3135
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11640
      TabIndex        =   5
      Top             =   2760
      Width           =   3135
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   7080
      TabIndex        =   14
      Top             =   960
      Width           =   3135
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0CCA
      Left            =   11640
      List            =   "Form1.frx":0CD7
      TabIndex        =   13
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFF00&
      Caption         =   "Edit"
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   10080
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   5655
      Left            =   4440
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   3840
      Width           =   10335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "Blank"
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   10080
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Exit"
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
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   10080
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Restore Deleted"
      Enabled         =   0   'False
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
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   10080
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Delete"
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
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   10080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Save"
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
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10080
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   4
      Left            =   6120
      Top             =   9840
      Width           =   8895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Source Codes"
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
      Index           =   10
      Left            =   4440
      TabIndex        =   27
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   6255
      Index           =   3
      Left            =   4200
      Top             =   3480
      Width           =   10815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "All Codes"
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
      Index           =   6
      Left            =   360
      TabIndex        =   26
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Founded"
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
      Index           =   7
      Left            =   360
      TabIndex        =   25
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   6255
      Index           =   2
      Left            =   120
      Top             =   3480
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
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
      Index           =   0
      Left            =   5880
      TabIndex        =   22
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Function"
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
      Index           =   1
      Left            =   5880
      TabIndex        =   21
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type of Code"
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
      Index           =   2
      Left            =   5880
      TabIndex        =   20
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Coding By"
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
      Index           =   3
      Left            =   10440
      TabIndex        =   19
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Side Effects"
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
      Index           =   4
      Left            =   10440
      TabIndex        =   18
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Writing Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   10440
      TabIndex        =   17
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   1815
      Index           =   1
      Left            =   5640
      Top             =   1560
      Width           =   9375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search by"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   10440
      TabIndex        =   16
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Word"
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
      Index           =   9
      Left            =   5880
      TabIndex        =   15
      Top             =   960
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   0
      Left            =   5640
      Top             =   720
      Width           =   9375
   End
   Begin VB.Image Image1 
      Height          =   3645
      Left            =   8040
      Picture         =   "Form1.frx":0CFC
      Top             =   10920
      Visible         =   0   'False
      Width           =   6000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'developer: Shyam Singh Chandel < shyamschandel@rediffmail.com >

Private Sub Command1_Click()
On Error Resume Next
If Text1.Text = "" Then
MsgBox "Enter Subject Please" & Chr(13) & "Thanks" & Chr(13) & "SSC"
Text1.SetFocus
Exit Sub
End If

If Text2.Text = "" Then
MsgBox "Enter Function Please" & Chr(13) & "Thanks" & Chr(13) & "SSC"
Text2.SetFocus
Exit Sub
End If

If Text3.Text = "" Then
MsgBox "Enter Type of Code Please" & Chr(13) & "Thanks" & Chr(13) & "SSC"
Text3.SetFocus
Exit Sub
End If

If Text4.Text = "" Then
MsgBox "Enter Programer Name Please" & Chr(13) & "Thanks" & Chr(13) & "SSC"
Text4.SetFocus
Exit Sub
End If

If Text5.Text = "" Then
MsgBox "Enter Side Effects Please" & Chr(13) & "Thanks" & Chr(13) & "SSC"
Text5.SetFocus
Exit Sub
End If

If Text6.Text = "" Then
MsgBox "Enter Date of Writing Code Please" & Chr(13) & "Thanks" & Chr(13) & "SSC"
Text6.SetFocus
Exit Sub
End If

If Text7.Text = "" Then
MsgBox "Enter Main Code Please" & Chr(13) & "Thanks" & Chr(13) & "SSC"
Text7.SetFocus
Exit Sub
End If

SQL = "SELECT * FROM CODEVIEW"
Set RS = New ADODB.Recordset
 RS.Open SQL, CN, adOpenStatic, adLockOptimistic
  RS.AddNew
    RS!SUBJECT = Text1.Text
    RS!Function = Text2.Text
    RS!TypeOFCode = Text3.Text
    RS!CodingBy = Text4.Text
    RS!SideEffects = Text5.Text
    RS!DateofWriting = Text6.Text
    RS!code = Text7.Text
  RS.Update
  MsgBox "Record has been saved"
  Command5_Click
   RS.Close
   List1.Clear
   load
End Sub

Private Sub Command2_Click()
On Error Resume Next
SQL = "SELECT * FROM CODEVIEW where Subject='" & List1.Text & "'"
Set RS = New ADODB.Recordset
RS.Open SQL, CN, adOpenStatic, adLockOptimistic
 Backup
RS.Delete
MsgBox "Record has been deleted"
Command5_Click
RS.Close
List1.Clear
load
End Sub

Private Sub Command3_Click()
On Error Resume Next
SQL = "SELECT * FROM CODEVIEW where Subject='" & List1.Text & "'"
Set RS = New ADODB.Recordset
RS.Open SQL, CN, adOpenStatic, adLockOptimistic
RestoreBackup
If Text1.Text = "" Then
MsgBox "NO RECORD FOR RESTORE"
Exit Sub
End If
RS.AddNew
    RS!SUBJECT = Text1.Text
    RS!Function = Text2.Text
    RS!TypeOFCode = Text3.Text
    RS!CodingBy = Text4.Text
    RS!SideEffects = Text5.Text
    RS!DateofWriting = Text6.Text
  RS.Update
  MsgBox "Record has been Restored"
  Command5_Click
  RS.Close
  List1.Clear
  load
  DeleteBackup
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Command5_Click()
BlankFields
End Sub

Private Sub Command6_Click()
On Error Resume Next
SQL = "SELECT * FROM CODEVIEW where Subject='" & Text1.Text & "'"
Set RS = New ADODB.Recordset
RS.Open SQL, CN, adOpenStatic, adLockOptimistic
If Text1.Text = "" Then
MsgBox "NO RECORD FOR EDITING"
Exit Sub
End If
currentMode = EditMode
    RS!Function = Text2.Text
    RS!TypeOFCode = Text3.Text
    RS!CodingBy = Text4.Text
    RS!SideEffects = Text5.Text
    RS!DateofWriting = Text6.Text
    RS!code = Text7.Text
RS.Update
RS.Close
MsgBox "Record is edited"

End Sub

Private Sub Form_Load()
Me.Picture = Image1.Picture
connectDB
load
Combo1.ListIndex = 0
End Sub
Private Sub load()
SQL = "SELECT * FROM CODEVIEW"
Set RS = New ADODB.Recordset
 RS.Open SQL, CN, adOpenStatic, adLockOptimistic
   Do While Not RS.EOF
      List1.AddItem RS!SUBJECT
    RS.MoveNext
    Loop
  RS.Close
End Sub

Private Sub List1_Click()
SQL = "SELECT * FROM CODEVIEW where Subject='" & List1.Text & "'"
Set RS = New ADODB.Recordset
 RS.Open SQL, CN, adOpenStatic, adLockOptimistic
    AddFields
 RS.Close
End Sub
Private Sub DeleteBackup()
Call SaveSetting("Restore", "Restore Data", "Subject", "")
Call SaveSetting("Restore", "Restore Data", "Function", "")
Call SaveSetting("Restore", "Restore Data", "TypeOFCode", "")
Call SaveSetting("Restore", "Restore Data", "CodingBy", "")
Call SaveSetting("Restore", "Restore Data", "SideEffects", "")
Call SaveSetting("Restore", "Restore Data", "DateofWriting", "")
Call SaveSetting("Restore", "Restore Data", "CODE", "")
End Sub
Sub Backup()
Call SaveSetting("Restore", "Restore Data", "Subject", Text1.Text)
Call SaveSetting("Restore", "Restore Data", "Function", Text2.Text)
Call SaveSetting("Restore", "Restore Data", "TypeOFCode", Text3.Text)
Call SaveSetting("Restore", "Restore Data", "CodingBy", Text4.Text)
Call SaveSetting("Restore", "Restore Data", "SideEffects", Text5.Text)
Call SaveSetting("Restore", "Restore Data", "DateofWriting", Text6.Text)
Call SaveSetting("Restore", "Restore Data", "CODE", Text7.Text)

End Sub
Sub RestoreBackup()
Text1.Text = GetSetting("Restore", "Restore Data", "Subject")
Text2.Text = GetSetting("Restore", "Restore Data", "Function")
Text3.Text = GetSetting("Restore", "Restore Data", "TypeOFCode")
Text4.Text = GetSetting("Restore", "Restore Data", "CodingBy")
Text5.Text = GetSetting("Restore", "Restore Data", "SideEffects")
Text6.Text = GetSetting("Restore", "Restore Data", "DateofWriting")
Text7.Text = GetSetting("Restore", "Restore Data", "CODE")

End Sub
Sub AddFields()
On Error Resume Next
    Text1 = RS!SUBJECT
    Text2 = RS!Function
    Text3 = RS!TypeOFCode
    Text4 = RS!CodingBy
    Text5 = RS!SideEffects
    Text6 = RS!DateofWriting
    Text7 = RS!code
End Sub
Sub BlankFields()
  Text1 = ""
  Text2 = ""
  Text3 = ""
  Text4 = ""
  Text5 = ""
  Text6 = ""
  Text7 = ""
  Text1.SetFocus
End Sub

Private Sub List2_Click()
SQL = "SELECT * FROM CODEVIEW where Subject='" & List2.Text & "'"
Set RS = New ADODB.Recordset
 RS.Open SQL, CN, adOpenStatic, adLockOptimistic
    AddFields
 RS.Close
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text3.SetFocus
End If
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text4.SetFocus
End If
End Sub
Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text5.SetFocus
End If
End Sub
Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text6.SetFocus
End If
End Sub
Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text7.SetFocus
End If
End Sub
Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
'Command1_Click
End If
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
On Error Resume Next
If Combo1.Text = "Subject" Then
SQL = "SELECT * FROM [" & "CODEVIEW" & "]  where [" & "SUBJECT" & "]" & "LIKE '%" & Text8.Text & "%'"
ElseIf Combo1.Text = "Function" Then
SQL = "SELECT * FROM [" & "CODEVIEW" & "]  where [" & "FUNCTION" & "]" & "LIKE '%" & Text8.Text & "%'"
ElseIf Combo1.Text = "Type of Code" Then
SQL = "SELECT * FROM [" & "CODEVIEW" & "]  where [" & "TypeOFCode" & "]" & "LIKE '%" & Text8.Text & "%'"
End If
Set RS = New ADODB.Recordset
 RS.Open SQL, CN, adOpenStatic, adLockOptimistic
    'AddFields
    Do While Not RS.EOF
        List2.AddItem RS!SUBJECT
    RS.MoveNext
    Loop
    
 RS.Close
 End If
 
End Sub
