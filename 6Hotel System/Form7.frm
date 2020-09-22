VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "About Hotel Management"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5565
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   ScaleHeight     =   3150
   ScaleWidth      =   5565
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   345
      Left            =   4155
      TabIndex        =   0
      Top             =   2745
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   15
      X2              =   5564
      Y1              =   2580
      Y2              =   2580
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   5564
      Y1              =   2565
      Y2              =   2565
   End
   Begin VB.Label Label8 
      Caption         =   "All rights reserved. "
      Height          =   255
      Left            =   990
      TabIndex        =   8
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Copyright © 2003-2005 Chetan Ghorpade. "
      Height          =   255
      Left            =   990
      TabIndex        =   7
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Easy Hotel Management v6.0.25"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   990
      TabIndex        =   6
      Top             =   0
      Width           =   3210
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Report bugs and comments to chetan (at) chetan (dot) tk"
      Height          =   255
      Left            =   990
      TabIndex        =   5
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "GalleOn-line India Networks 50, CGC-2, V.M.V. Road, Amravati 444604 (MS) IN"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   990
      TabIndex        =   4
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "This product is Licensed under the terms of GNU/GPL License Agreement."
      Height          =   495
      Left            =   990
      TabIndex        =   3
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label Label6 
      Caption         =   "For further information in details visit:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   990
      TabIndex        =   2
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "http://www.chetan.tk"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3630
      MouseIcon       =   "Form7.frx":08CA
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   270
      Picture         =   "Form7.frx":0BD4
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbRed
Label4.FontUnderline = False
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbRed
Label4.FontUnderline = False
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbRed
Label4.FontUnderline = False
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbRed
Label4.FontUnderline = False
End Sub

Private Sub Label4_Click()
ShellExecute Me.hWnd, "open", "http://www.chetan.tk", vbNullString, "", 0
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbBlue
Label4.FontUnderline = True
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbRed
Label4.FontUnderline = False
End Sub


