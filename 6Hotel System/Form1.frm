VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hotel Management"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4200
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "db\hotel2.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   -40000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   2  'Snapshot
         RecordSource    =   "user"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton Command3 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   "db\hotel2.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -40000
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "admin"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00404040&
         Caption         =   "User Login"
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00404040&
         Caption         =   "Admin Login"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   2880
         Width           =   1335
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000009&
         Height          =   2340
         Left            =   120
         Picture         =   "Form1.frx":08CA
         ScaleHeight     =   2280
         ScaleWidth      =   3690
         TabIndex        =   6
         Top             =   240
         Width           =   3750
      End
      Begin VB.Label Label3 
         Caption         =   "Password"
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "User Name"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   2880
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset

Private Sub Command1_Click()
With Data2
.Recordset.MoveFirst
Do Until .Recordset.EOF
If (.Recordset.Fields(0) = Text1.Text) And (.Recordset.Fields(1) = Text2.Text) Then
'If Text1.Text = Data1.Recordset("userid") Then '& Text2.Text = Data1.Recordset("password") Then
Unload Me
Form6.Show
Exit Sub
Else
.Recordset.MoveNext
End If
Loop
'Label4.Caption = "Invalid username or ID.Try again.. "
MsgBox "Invalid Username or password.Try again...", vbCritical, "HMS"
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End With
End Sub

Private Sub Command2_Click()
With Data1
.Recordset.MoveFirst
Do Until .Recordset.EOF
If (.Recordset.Fields("userid") = Text1.Text) And (.Recordset.Fields("password") = Text2.Text) Then
'If Text1.Text = Data1.Recordset("userid") Then '& Text2.Text = Data1.Recordset("password") Then
Unload Me
Form2.Show
Form2.Caption = "Welcome user..."
Exit Sub
Else
.Recordset.MoveNext
End If
Loop
'Label4.Caption = "Invalid username or ID.Try again.. "
MsgBox "Invalid Username or password.Try again...", vbCritical, "HMS"
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End With
End Sub

Private Sub Command3_Click()
Close Databases
End
End Sub

Private Sub Form_Load()
'Set db = OpenDatabase("hotel.mdb")
'Set rs = db.OpenRecordset("admin")
Data1.DatabaseName = App.Path & "\db\hotel2.mdb"
Data1.RecordSource = "select * from user"
Data2.DatabaseName = App.Path & "\db\hotel2.mdb"
Data2.RecordSource = "select * from admin"
End Sub

