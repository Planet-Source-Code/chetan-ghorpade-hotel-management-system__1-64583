VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6810
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9405
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   9405
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab2 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "MAIN"
      TabPicture(0)   =   "Form2.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label45"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label46"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Picture2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "ISSUE of ROOMS"
      TabPicture(1)   =   "Form2.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "Data3"
      Tab(1).Control(3)=   "Data2"
      Tab(1).Control(4)=   "Data1"
      Tab(1).Control(5)=   "Frame7"
      Tab(1).Control(6)=   "Label37"
      Tab(1).Control(7)=   "Label1(0)"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "BOOKING"
      TabPicture(2)   =   "Form2.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(2)=   "Frame2"
      Tab(2).Control(3)=   "Label18"
      Tab(2).Control(4)=   "Label24"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Checkout Details"
      TabPicture(3)   =   "Form2.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(1)=   "Frame5"
      Tab(3).Control(2)=   "Label15"
      Tab(3).Control(3)=   "Label43"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "About Author"
      TabPicture(4)   =   "Form2.frx":093A
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Image1(0)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Image1(1)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Line1(0)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Line1(1)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Label8(1)"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Label7(1)"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Label2(1)"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Label3(1)"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Label5(1)"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Label6(1)"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "Label4(1)"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "Image1(2)"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "Image2"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "Label13"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).ControlCount=   14
      Begin VB.Frame Frame8 
         Height          =   5055
         Left            =   -74760
         TabIndex        =   104
         Top             =   1320
         Width           =   2415
         Begin VB.CommandButton Command20 
            Caption         =   "Room status"
            Height          =   375
            Left            =   240
            TabIndex        =   107
            Top             =   3720
            Width           =   1815
         End
         Begin VB.ListBox List4 
            Height          =   2985
            Index           =   0
            ItemData        =   "Form2.frx":0956
            Left            =   1200
            List            =   "Form2.frx":0958
            TabIndex        =   106
            Top             =   480
            Width           =   1095
         End
         Begin VB.ListBox List4 
            Height          =   2985
            Index           =   1
            ItemData        =   "Form2.frx":095A
            Left            =   120
            List            =   "Form2.frx":095C
            TabIndex        =   105
            Top             =   480
            Width           =   1095
         End
         Begin VB.Timer Timer1 
            Interval        =   100
            Left            =   2280
            Top             =   120
         End
         Begin VB.Label Label42 
            Caption         =   "Vacant"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   1440
            TabIndex        =   109
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label41 
            Caption         =   "Occupied"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   108
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   -72240
         TabIndex        =   100
         Top             =   5520
         Width           =   6135
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   840
            TabIndex        =   102
            Top             =   360
            Width           =   3135
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Search"
            Height          =   375
            Left            =   4080
            TabIndex        =   101
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "Name"
            Height          =   255
            Left            =   240
            TabIndex        =   103
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Data Data3 
         Caption         =   "Data3"
         Connect         =   "Access"
         DatabaseName    =   "db\hotel2.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -73440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   360
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   "db\hotel2.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -71520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "reservation"
         Top             =   360
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "db\hotel2.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -69600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "room"
         Top             =   360
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.Frame Frame7 
         Height          =   4215
         Left            =   -72240
         TabIndex        =   70
         Top             =   1320
         Width           =   6135
         Begin VB.CommandButton Command4 
            Caption         =   "Update"
            Height          =   375
            Left            =   3840
            TabIndex        =   89
            Top             =   2880
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Edit"
            Height          =   375
            Left            =   4920
            TabIndex        =   88
            Top             =   2280
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Add"
            Height          =   375
            Left            =   3840
            TabIndex        =   87
            Top             =   2280
            Width           =   855
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "Form2.frx":095E
            Left            =   4800
            List            =   "Form2.frx":099B
            TabIndex        =   86
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   4800
            TabIndex        =   85
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   1200
            TabIndex        =   84
            Top             =   3240
            Width           =   1815
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   1200
            TabIndex        =   83
            Top             =   2760
            Width           =   1815
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   1200
            TabIndex        =   82
            Top             =   2280
            Width           =   1815
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   1200
            TabIndex        =   81
            Top             =   1800
            Width           =   2295
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1200
            TabIndex        =   80
            Top             =   1320
            Width           =   495
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "Form2.frx":09FE
            Left            =   4800
            List            =   "Form2.frx":0A08
            TabIndex        =   79
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1200
            TabIndex        =   78
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1200
            TabIndex        =   77
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Remove"
            Height          =   375
            Left            =   4920
            TabIndex        =   76
            Top             =   2880
            Width           =   855
         End
         Begin VB.CommandButton Command15 
            Caption         =   "first"
            Height          =   255
            Left            =   600
            TabIndex        =   75
            Top             =   3840
            Width           =   615
         End
         Begin VB.CommandButton Command16 
            Caption         =   "next"
            Height          =   255
            Left            =   1440
            TabIndex        =   74
            Top             =   3840
            Width           =   615
         End
         Begin VB.CommandButton Command17 
            Caption         =   "prev"
            Height          =   255
            Left            =   2280
            TabIndex        =   73
            Top             =   3840
            Width           =   615
         End
         Begin VB.CommandButton Command18 
            Caption         =   "last"
            Height          =   255
            Left            =   3120
            TabIndex        =   72
            Top             =   3840
            Width           =   615
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Clear"
            Height          =   375
            Left            =   4920
            TabIndex        =   71
            Top             =   3480
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Pincode"
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   2760
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "Room no"
            Height          =   255
            Left            =   3840
            TabIndex        =   98
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "City"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Phone"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   96
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Arrival Time"
            Height          =   255
            Index           =   0
            Left            =   3840
            TabIndex        =   95
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "Date of arrival"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   94
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Address"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   93
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Sex"
            Height          =   255
            Index           =   0
            Left            =   3840
            TabIndex        =   92
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "Age"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   91
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Name"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   90
            Top             =   840
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   5955
         Left            =   -74880
         Picture         =   "Form2.frx":0A1A
         ScaleHeight     =   5895
         ScaleWidth      =   8910
         TabIndex        =   69
         Top             =   480
         Width           =   8970
      End
      Begin VB.Frame Frame3 
         Height          =   4935
         Left            =   -74760
         TabIndex        =   57
         Top             =   1320
         Width           =   3375
         Begin VB.CommandButton Command8 
            Caption         =   "Booking list"
            Height          =   375
            Left            =   480
            TabIndex        =   62
            Top             =   3360
            Width           =   2175
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Booking fix"
            Height          =   375
            Left            =   480
            TabIndex        =   61
            Top             =   3840
            Width           =   2175
         End
         Begin VB.ListBox List1 
            Height          =   2400
            ItemData        =   "Form2.frx":10E91
            Left            =   240
            List            =   "Form2.frx":10E93
            TabIndex        =   60
            Top             =   720
            Width           =   1455
         End
         Begin VB.ListBox List2 
            Height          =   2400
            ItemData        =   "Form2.frx":10E95
            Left            =   1680
            List            =   "Form2.frx":10E97
            TabIndex        =   59
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton Command23 
            Caption         =   "Delete Booking"
            Height          =   375
            Left            =   480
            TabIndex        =   58
            Top             =   4320
            Width           =   2175
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   2235
            Left            =   240
            TabIndex        =   63
            Top             =   720
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   3942
            _Version        =   393217
            Enabled         =   -1  'True
            TextRTF         =   $"Form2.frx":10E99
         End
         Begin VB.Label Label39 
            Caption         =   "Name"
            Height          =   255
            Left            =   600
            TabIndex        =   65
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label40 
            Caption         =   "Booking date"
            Height          =   255
            Left            =   1800
            TabIndex        =   64
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1095
         Left            =   -71280
         TabIndex        =   52
         Top             =   5160
         Width           =   4935
         Begin VB.TextBox Text15 
            Height          =   285
            Left            =   2160
            TabIndex        =   54
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Search"
            Height          =   375
            Left            =   3720
            TabIndex        =   53
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label25 
            Caption         =   "Search for resevation"
            Height          =   255
            Left            =   360
            TabIndex        =   55
            Top             =   480
            Width           =   1575
         End
      End
      Begin VB.Frame Frame6 
         Height          =   4695
         Left            =   -74760
         TabIndex        =   46
         Top             =   1320
         Width           =   3015
         Begin VB.TextBox Text26 
            Height          =   285
            Left            =   960
            TabIndex        =   50
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Search"
            Height          =   375
            Left            =   840
            TabIndex        =   49
            Top             =   1800
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Room No"
            Height          =   255
            Left            =   360
            TabIndex        =   48
            Top             =   1320
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Name"
            Height          =   255
            Left            =   1560
            TabIndex        =   47
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label36 
            Caption         =   "Search"
            Height          =   255
            Left            =   240
            TabIndex        =   51
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Frame Frame5 
         Height          =   4695
         Left            =   -71640
         TabIndex        =   19
         Top             =   1320
         Width           =   5535
         Begin VB.TextBox Text16 
            Height          =   285
            Left            =   1440
            TabIndex        =   33
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox Text17 
            Height          =   285
            Left            =   1440
            TabIndex        =   32
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox Text18 
            Height          =   285
            Left            =   1440
            TabIndex        =   31
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox Text19 
            Height          =   285
            Left            =   1440
            TabIndex        =   30
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox Text20 
            Height          =   285
            Left            =   1440
            TabIndex        =   29
            Top             =   2760
            Width           =   1215
         End
         Begin VB.TextBox Text21 
            Height          =   285
            Left            =   1440
            TabIndex        =   28
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text22 
            Height          =   285
            Left            =   4200
            TabIndex        =   27
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox Text23 
            Height          =   285
            Left            =   4200
            TabIndex        =   26
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox Text24 
            Height          =   285
            Left            =   4200
            TabIndex        =   25
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox Text25 
            Height          =   285
            Left            =   4200
            TabIndex        =   24
            Top             =   1800
            Width           =   855
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Save"
            Height          =   375
            Left            =   3960
            TabIndex        =   23
            Top             =   2640
            Width           =   1095
         End
         Begin VB.TextBox Text27 
            Height          =   285
            Left            =   1440
            TabIndex        =   22
            Top             =   3240
            Width           =   1095
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Clear"
            Height          =   375
            Left            =   3960
            TabIndex        =   21
            Top             =   3120
            Width           =   1095
         End
         Begin VB.TextBox Text28 
            Height          =   285
            Left            =   1440
            TabIndex        =   20
            Top             =   3720
            Width           =   615
         End
         Begin VB.Label Label26 
            Caption         =   "Name"
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label27 
            Caption         =   "Address"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label28 
            Caption         =   "Phone"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label29 
            Caption         =   "Checkin Date "
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label Label30 
            Caption         =   "Checkout date"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label Label31 
            Caption         =   "Room Number"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label32 
            Caption         =   "Amount"
            Height          =   255
            Left            =   3120
            TabIndex        =   39
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label33 
            Caption         =   "Taxes"
            Height          =   255
            Left            =   3120
            TabIndex        =   38
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label34 
            Caption         =   "Services"
            Height          =   255
            Left            =   3120
            TabIndex        =   37
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label35 
            Caption         =   "Total Amount"
            Height          =   255
            Left            =   3120
            TabIndex        =   36
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label38 
            Caption         =   "Checkout time"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   3240
            Width           =   1095
         End
         Begin VB.Label Label44 
            Caption         =   "Duration"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   3720
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3735
         Left            =   -71280
         TabIndex        =   3
         Top             =   1320
         Width           =   4935
         Begin VB.CommandButton Command19 
            Caption         =   "Update"
            Height          =   375
            Left            =   2280
            TabIndex        =   13
            Top             =   3240
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Confirmed Arrival"
            Height          =   195
            Left            =   1680
            TabIndex        =   12
            Top             =   2760
            Width           =   1815
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Clear"
            Height          =   375
            Left            =   3240
            TabIndex        =   11
            Top             =   3240
            Width           =   735
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Edit"
            Height          =   375
            Left            =   1320
            TabIndex        =   10
            Top             =   3240
            Width           =   735
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Add"
            Height          =   375
            Left            =   360
            TabIndex        =   9
            Top             =   3240
            Width           =   735
         End
         Begin VB.TextBox Text14 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "M/dd/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Left            =   1680
            TabIndex        =   8
            Top             =   2280
            Width           =   1695
         End
         Begin VB.TextBox Text13 
            Height          =   285
            Left            =   1680
            TabIndex        =   7
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   1680
            TabIndex        =   6
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   1680
            TabIndex        =   5
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   1680
            TabIndex        =   4
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label23 
            Caption         =   "Date"
            Height          =   255
            Left            =   360
            TabIndex        =   18
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label22 
            Caption         =   "Estimated  Arrival"
            Height          =   255
            Left            =   360
            TabIndex        =   17
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label21 
            Caption         =   "Phone"
            Height          =   255
            Left            =   360
            TabIndex        =   16
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label20 
            Caption         =   "Address"
            Height          =   255
            Left            =   360
            TabIndex        =   15
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label19 
            Caption         =   "Name"
            Height          =   255
            Left            =   360
            TabIndex        =   14
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No. +91 9890 415715"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3030
         TabIndex        =   119
         Top             =   3840
         Width           =   2895
      End
      Begin VB.Image Image2 
         Height          =   990
         Left            =   1080
         Picture         =   "Form2.frx":10F1B
         Top             =   840
         Width           =   3885
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   630
         Picture         =   "Form2.frx":12B66
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.chetan.tk"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   6390
         MouseIcon       =   "Form2.frx":13430
         MousePointer    =   99  'Custom
         TabIndex        =   118
         Top             =   5400
         Width           =   2055
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "For other softwares, products and Services visit:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3030
         TabIndex        =   117
         Top             =   5160
         Width           =   4935
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "This product is Licensed under the terms of GNU/GPL License Agreement."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   3030
         TabIndex        =   116
         Top             =   4200
         Width           =   3615
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "GalleOn-line India Networks 50, Cotton Green Colony-2, V.M.V. Road, Amravati 444604 (Maharashtra) IN"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   1
         Left            =   3030
         TabIndex        =   115
         Top             =   2880
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Report bugs and comments to chetan (at) chetan (dot) tk"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3030
         TabIndex        =   114
         Top             =   4800
         Width           =   5655
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2003-2005 Mr. Chetan Ghorpade. "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3030
         TabIndex        =   113
         Top             =   2280
         Width           =   4455
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "All rights reserved. "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3030
         TabIndex        =   112
         Top             =   2520
         Width           =   3255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   240
         X2              =   8880
         Y1              =   1965
         Y2              =   1965
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0080C0FF&
         BorderWidth     =   3
         Index           =   0
         X1              =   240
         X2              =   8880
         Y1              =   1980
         Y2              =   1980
      End
      Begin VB.Image Image1 
         Height          =   5265
         Index           =   1
         Left            =   3480
         Picture         =   "Form2.frx":1373A
         Top             =   840
         Width           =   5580
      End
      Begin VB.Image Image1 
         Height          =   5265
         Index           =   0
         Left            =   120
         Picture         =   "Form2.frx":1A774
         Top             =   840
         Width           =   5580
      End
      Begin VB.Label Label37 
         Caption         =   "Room Status"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   -74520
         TabIndex        =   111
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Checkin Information"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   -71280
         TabIndex        =   110
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "New Reservation"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   -70680
         TabIndex        =   68
         Top             =   965
         Width           =   3735
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Caption         =   "Booking  List"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   -74160
         TabIndex        =   67
         Top             =   965
         Width           =   2175
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Search Customer"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   66
         Top             =   965
         Width           =   2775
      End
      Begin VB.Label Label43 
         Caption         =   "Check Out Details"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   -69960
         TabIndex        =   56
         Top             =   965
         Width           =   2655
      End
      Begin VB.Label Label46 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68520
         TabIndex        =   1
         Top             =   2580
         Width           =   1935
      End
      Begin VB.Label Label45 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72600
         TabIndex        =   2
         Top             =   2580
         Width           =   2295
      End
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu logoff 
         Caption         =   "&Log off"
         Shortcut        =   ^L
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu util 
      Caption         =   "utilities"
      Begin VB.Menu rmst 
         Caption         =   "&Room status"
      End
      Begin VB.Menu rep 
         Caption         =   "&View Reports"
      End
      Begin VB.Menu stat 
         Caption         =   "&Statistics"
      End
      Begin VB.Menu seperater 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu dbms 
         Caption         =   "&Database Manager"
      End
   End
   Begin VB.Menu hlp 
      Caption         =   "Help"
      Begin VB.Menu hlphms 
         Caption         =   "Help on HMS"
      End
      Begin VB.Menu about 
         Caption         =   "About Us"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim db1 As Database
Dim rs1 As Recordset
Dim db2 As Database
Dim rs2 As Recordset
Dim val1 As Integer
Dim val2 As Integer
Dim val3 As Integer
Dim val As Integer


Private Sub about_Click()
Form7.Show 1
End Sub

Private Sub Command1_Click()
If Text2.Text = "" Then
MsgBox "Please enter name", vbInformation, "HMS"
Text2.SetFocus
Else
If Text3.Text = "" Then
MsgBox "Please enter age", vbInformation, "HMS"
Text3.SetFocus
Else
If Text4.Text = "" Then
MsgBox "Please enter address", vbInformation, "HMS"
Text4.SetFocus
Else
If Text5.Text = "" Then
MsgBox "please enter city", vbInformation, "HMS"
Text5.SetFocus
Else
If Text6.Text = "" Then
MsgBox "please enter pin", vbInformation, "HMS"
Text6.SetFocus
Else
If Text7.Text = "" Then
MsgBox "please enter phone", vbInformation, "HMS"
Text7.SetFocus
Else
If Combo1.Text = "" Then
MsgBox "please enter sex", vbInformation, "HMS"
Combo1.SetFocus
Else
If Combo2.Text = "" Then
MsgBox "please enter roomno", vbInformation, "HMS"
Combo2.SetFocus
Else
rs.AddNew
rs.Fields(0) = Text1.Text
rs.Fields(1) = Text9.Text
rs.Fields(2) = Text2.Text
rs.Fields(3) = Combo1.Text
rs.Fields(4) = Text3.Text
rs.Fields(5) = Text4.Text
rs.Fields(6) = Text5.Text
rs.Fields(7) = Text6.Text
rs.Fields(8) = Text7.Text
rs.Fields(9) = Combo2.Text
rs.Update
Data1.Recordset.MoveFirst
Do Until Data1.Recordset.EOF
If Data1.Recordset.Fields(0) = Combo2.Text Then
Data1.Recordset.Edit
Data1.Recordset.Fields(1) = True
Data1.Recordset.Update
MsgBox ("Data added. Room alloted for visitor") + Combo2.Text, vbInformation, "HMS"
Exit Sub
Else
Data1.Recordset.MoveNext
End If
Loop
End If
End If
End If
End If
End If
End If
End If
End If

End Sub

Private Sub Command10_Click()
rs1.MoveFirst
Do Until rs1.EOF
If rs1.Fields(1) = Text15.Text And rs1.Fields(5) = True Then
Call rescheck
Exit Sub
Else
If (rs1.Fields(1) = Text15.Text) And rs1.Fields(5) = False Then
Text11.Text = rs1.Fields(1)
Text12.Text = rs1.Fields(2)
Text13.Text = rs1.Fields(3)
Text14.Text = rs1.Fields(4)
Check1.Value = 0
'Check1.Value = rs1.Fields(5)
Exit Sub
Else
rs1.MoveNext
End If
End If
Loop
MsgBox "No data found.Try again..", vbInformation, "HMS"
End Sub

Private Sub Command11_Click()
rs2.AddNew
rs2.Fields(0) = Text16.Text
rs2.Fields(1) = Text17.Text
rs2.Fields(2) = Text18.Text
rs2.Fields(3) = Text21.Text
rs2.Fields(4) = Text19.Text
rs2.Fields(5) = Text20.Text
rs2.Fields(6) = DateValue(Format(Now, "Short Date")) - rs.Fields(0)
rs2.Fields(7) = Text25.Text
rs2.Update
Call chkoutroom
rs.MoveFirst
Do Until rs.EOF
If rs.Fields(2) = Text16.Text Then
rs.Delete
Exit Sub
Else
rs.MoveNext
End If
Loop
MsgBox "Guest Checked out sucessfuly...", vbInformation, "HMS"

End Sub

Private Sub Command12_Click()
Form2.PrintForm

End Sub

Private Sub Command13_Click()
rs.MoveFirst
If Option1.Value = True Then
Call optionsearch
Else
Do Until rs.EOF
If rs.Fields(2) = Text26.Text Then
Text16.Text = rs.Fields(2)
Text17.Text = rs.Fields(5)
Text18.Text = rs.Fields(8)
Text21.Text = rs.Fields(9)
Text19.Text = rs.Fields(0)
val = DateValue(Format(Now, "Short Date")) - rs.Fields(0)
Text28.Text = DateValue(Format(Now, "Short Date")) - rs.Fields(0)
Text22.Text = val + 300
Text23.Text = (10 / 100) * Text22.Text
Text24.Text = (20 / 100) * Text23.Text
val1 = Int(Text22.Text)
val2 = Int(Text23.Text)
val3 = Int(Text24.Text)

Text25.Text = val1 + val2 + val3
Text16.Enabled = True
Text17.Enabled = True
Text18.Enabled = True
Text19.Enabled = True
Text21.Enabled = True
Text22.Enabled = True
Text23.Enabled = True
Text24.Enabled = True
Text25.Enabled = True
Command11.Enabled = True
'Command12.Enabled = True
Command21.Enabled = True
Exit Sub
Else
rs.MoveNext
End If
Loop
MsgBox "no datas found.", vbInformation, "HMS"
Text26.Text = ""
Text26.SetFocus

End If
End Sub

Private Sub Command14_Click()
If rs.BOF Or rs.EOF = True Then
MsgBox "  END OF FILE", vbOKOnly, "HMS"
Else
rs.Delete
End If
End Sub

Private Sub Command15_Click()
If rs.BOF = True Then
MsgBox "beginning of record", vbOKOnly, "HMS"
Else
rs.MoveFirst
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(2)
Combo1.Text = rs.Fields(3)
Text3.Text = rs.Fields(4)
Text4.Text = rs.Fields(5)
Text5.Text = rs.Fields(6)
Text6.Text = rs.Fields(7)
Text7.Text = rs.Fields(8)
Combo2.Text = rs.Fields(9)
End If
End Sub

Private Sub Command16_Click()
If rs.EOF <> True Then
rs.MoveNext
If rs.EOF = True Then
rs.MoveLast
MsgBox "End of record", vbInformation, "HMS"
Else
'rs.MoveNext
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(2)
Combo1.Text = rs.Fields(3)
Text3.Text = rs.Fields(4)
Text4.Text = rs.Fields(5)
Text5.Text = rs.Fields(6)
Text6.Text = rs.Fields(7)
Text7.Text = rs.Fields(8)
Combo2.Text = rs.Fields(9)
End If
End If
End Sub

Private Sub Command17_Click()
If rs.BOF <> True Then
rs.MovePrevious
If rs.BOF = True Then
rs.MoveFirst
MsgBox "begining of record", vbInformation, "HMS"
Else
'rs.MovePrevious
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(2)
Combo1.Text = rs.Fields(3)
Text3.Text = rs.Fields(4)
Text4.Text = rs.Fields(5)
Text5.Text = rs.Fields(6)
Text6.Text = rs.Fields(7)
Text7.Text = rs.Fields(8)
Combo2.Text = rs.Fields(9)
End If
End If
End Sub

Private Sub Command18_Click()
If rs.EOF = True Then
MsgBox ("End of record")
Else
rs.MoveLast
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(2)
Combo1.Text = rs.Fields(3)
Text3.Text = rs.Fields(4)
Text4.Text = rs.Fields(5)
Text5.Text = rs.Fields(6)
Text6.Text = rs.Fields(7)
Text7.Text = rs.Fields(8)
Combo2.Text = rs.Fields(9)
End If
End Sub

Private Sub Command19_Click()
rs1.Edit
rs1.Fields(1) = Text11.Text
rs1.Fields(2) = Text12.Text
rs1.Fields(3) = Text13.Text
rs1.Fields(4) = Text14.Text
rs1.Fields(5) = Check1.Value
rs1.Update
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True

MsgBox "Reservation for Guest is updated", vbInformation, "HMS"
Command19.Enabled = False
End Sub

Private Sub Command2_Click()
NameQuery = InputBox("Enter A Name To Search For", "Name Query")
rs.MoveFirst
Do Until rs.EOF
If rs.Fields("name") = NameQuery Then
Text1.Text = rs.Fields(0)
Text9.Text = rs.Fields(1)
Text2.Text = rs.Fields(2)
Combo1.Text = rs.Fields(3)
Text3.Text = rs.Fields(4)
Text4.Text = rs.Fields(5)
Text5.Text = rs.Fields(6)
Text6.Text = rs.Fields(7)
Text7.Text = rs.Fields(8)
Combo2.Text = rs.Fields(9)
MsgBox "Match Found.Edit the record", vbInformation, "HMS"
Command1.Enabled = False
Command2.Enabled = False
Command14.Enabled = False
Command22.Enabled = True
Command4.Enabled = True

Exit Sub
Else
rs.MoveNext
End If
Loop
MsgBox ("No matches found.Please try again.."), vbCritical, "HMS"
End Sub

Private Sub Command20_Click()
List4(0).Clear
List4(1).Clear
Call roomstatus
End Sub

Private Sub Command21_Click()
Text16 = ""
Text17 = ""
Text18 = ""
Text19 = ""
Text21 = ""
Text22 = ""
Text23 = ""
Text24 = ""
Text25 = ""
Text28 = ""
Text26 = ""
Text26.SetFocus
End Sub

Private Sub Command22_Click()
Text2.SetFocus
Text2.Text = ""
Combo1.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8 = ""
Combo2.Text = ""
Command1.Enabled = True
Command2.Enabled = True
Command14.Enabled = False

End Sub

Private Sub Command23_Click()
Call expireconfirmation
End Sub

Private Sub Command3_Click()
rs.MoveFirst
Do Until rs.EOF
If rs.Fields("name") = Text8.Text Then
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(2)
Combo1.Text = rs.Fields(3)
Text3.Text = rs.Fields(4)
Text4.Text = rs.Fields(5)
Text5.Text = rs.Fields(6)
Text6.Text = rs.Fields(7)
Text7.Text = rs.Fields(8)
Combo2.Text = rs.Fields(9)
Command1.Enabled = False
Exit Sub
Else
rs.MoveNext
End If
Loop
MsgBox "No matches found.Please try again..", vbInformation, "HMS"
Text8.Text = ""

Text8.SetFocus
End Sub

Private Sub Command4_Click()
rs.Edit
rs.Fields(0) = Text1.Text
rs.Fields(1) = Text9.Text
rs.Fields(2) = Text2.Text
rs.Fields(3) = Combo1.Text
rs.Fields(4) = Text3.Text
rs.Fields(5) = Text4.Text
rs.Fields(6) = Text5.Text
rs.Fields(7) = Text6.Text
rs.Fields(8) = Text7.Text
rs.Fields(9) = Combo2.Text
rs.Update
MsgBox "current record is updated", vbInformation, "HMS"
Command1.Enabled = True
Command22.Enabled = True
Command2.Enabled = True
Command4.Enabled = False
Command22.SetFocus
End Sub

Private Sub Command5_Click()
If Text11 = "" Then
MsgBox "Please enter name", vbInformation, "HMS"
Text11.SetFocus
Else
If Text12 = "" Then
MsgBox "Please enter address", vbInformation, "HMS"
Text12.SetFocus
Else
If Text13 = "" Then
MsgBox "Please enter phone", vbInformation, "HMS"
Text13.SetFocus
Else
If Text14 = "" Then
MsgBox "Please enter estimated arrival", vbInformation, "HMS"
Text14.SetFocus
Else
rs1.AddNew
rs1.Fields(0) = Text10.Text
rs1.Fields(1) = Text11.Text
rs1.Fields(2) = Text12.Text
rs1.Fields(3) = Text13.Text
rs1.Fields(4) = Text14.Text
rs1.Fields(5) = Check1.Value
rs1.Update
MsgBox "Reservation for new visitor added", vbOKOnly, "HMS"
End If
End If
End If
End If

End Sub

Private Sub Command6_Click()
resinput = InputBox("Enter the name to be edited", "resinput")
rs1.MoveFirst
Do Until rs1.EOF
If rs1.Fields(1) = resinput And rs1.Fields(5) = True Then
Call rescheck1
Exit Sub
Else
If (rs1.Fields(1) = resinput) And rs1.Fields(5) = False Then
Text11.Text = rs1.Fields(1)
Text12.Text = rs1.Fields(2)
Text13.Text = rs1.Fields(3)
Text14.Text = rs1.Fields(4)
Check1.Value = 0
Command5.Enabled = False
Command7.Enabled = True
Command19.Enabled = True
Command6.Enabled = False
Exit Sub
Else
rs1.MoveNext
End If
End If
Loop
MsgBox "No data found.Try again..", vbOKOnly, "HMS"
End Sub

Private Sub Command7_Click()
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15 = ""
Command5.Enabled = True
Command6.Enabled = True
Command19.Enabled = False
Check1.Value = 0
End Sub

Private Sub Command8_Click()
List1.Clear
List2.Clear
Call ResList
End Sub

Private Sub Command9_Click()
List1.Clear
List2.Clear
Call Resconfirmed
End Sub

Private Sub exit_Click()
Close Databases
End
End Sub

Private Sub Form_Load()

Set db = OpenDatabase(App.Path + "/db/hotel2.mdb")
Set rs = db.OpenRecordset("checkin")

Set db1 = OpenDatabase(App.Path + "/db/hotel2.mdb")
Set rs1 = db1.OpenRecordset("reservation")

Set db2 = OpenDatabase(App.Path + "/db/hotel2.mdb")
Set rs2 = db2.OpenRecordset("checkout")

Text10.Text = Date
Text20.Text = Date
Text1.Text = Date
Command4.Enabled = False
Command14.Enabled = False
Command19.Enabled = False

Text16.Enabled = False
Text17.Enabled = False
Text18.Enabled = False
Text19.Enabled = False
Text21.Enabled = False
Text22.Enabled = False
Text23.Enabled = False
Text24.Enabled = False
Text25.Enabled = False
Command11.Enabled = False

Command21.Enabled = False
Text28.Enabled = False
Label45.Caption = Format(Date, "Long Date")
Label46.Caption = Time
End Sub


Private Sub logoff_Click()
Close Databases
Unload Form2
Unload Form3
Unload Form4
Unload Form5
Form1.Show
End Sub

Private Sub rep_Click()
Form4.Show 1
End Sub

Private Sub rmst_Click()
Form5.Show vbModal
End Sub

Private Sub stat_Click()
Form3.Show
End Sub

Private Sub Timer1_Timer()
Text9.Text = Time()
Text27.Text = Time()
Label46.Caption = Time()
End Sub

Private Sub ResList()
Dim strSQL As String

On Error Resume Next
strSQL = "Select * from Reservation"
List1.Clear
List2.Clear

        
With Data2
   .RecordSource = strSQL
   .Refresh
   .Recordset.MoveFirst

    Do Until .Recordset.EOF
    If .Recordset.Fields("confirmed") = 0 Then
        List1.AddItem .Recordset("name")
        List2.AddItem .Recordset("arrivaldate")
        End If
        .Recordset.MoveNext
    Loop
    
End With
End Sub


Private Sub Resconfirmed()
Dim sql As String

On Error Resume Next
sql = "Select * from Reservation"
List1.Clear
List2.Clear
   
With Data2
   .RecordSource = sql
   .Refresh
   .Recordset.MoveFirst

    Do Until .Recordset.EOF
    If .Recordset.Fields(5) = True Then
        List1.AddItem .Recordset("name")
        List2.AddItem .Recordset("arrivaldate")
        End If
        .Recordset.MoveNext
    Loop
    
End With
End Sub

Private Sub roomstatus()
Dim sql As String

sql = "Select * from room"

With Data1
   .RecordSource = sql
   .Refresh
   .Recordset.MoveFirst
    Do Until .Recordset.EOF
    On Error Resume Next
    If .Recordset.Fields(1) = True Then
        List4(1).AddItem .Recordset("roomno")
        'List3.AddItem.Index (1)
        Else: List4(0).AddItem .Recordset("roomno") 'Fill listbox for Rooms Tab
        End If
        
        .Recordset.MoveNext
        Loop
        End With
End Sub
Private Sub chkoutroom()


Data1.Recordset.MoveFirst
Do Until Data1.Recordset.EOF
If Data1.Recordset.Fields(0) = Text21.Text Then
Data1.Recordset.Edit
Data1.Recordset.Fields(1) = False
Data1.Recordset.Update
MsgBox ("Visitor sucessfully checked out..") + Combo2.Text, vbOKOnly, "HMS"
Exit Sub
Else
Data1.Recordset.MoveNext
End If
Loop
End Sub


Private Sub checkinvalidate(checkin_form_error)
Let checkin_form_error = False

If Text2 = "" Then
Text2.SetFocus
MsgBox "Please enter the name", vbExclamation, Error
Let checkin_form_error = True
Exit Sub
ElseIf Text3.Text = "" Then
Text3.SetFocus
MsgBox "Please enter the age", vbExclamation, Error
Let checkin_form_error = True
ElseIf Text4.Text = "" Then
Text4.SetFocus
MsgBox "Please enter the address", vbExclamation, Error
Let checkin_form_error = True
ElseIf Text5.Text = "" Then
Text5.SetFocus
MsgBox "Please enter the city", vbExclamation, Error
Let checkin_form_error = True
ElseIf Text6.Text = "" Then
Text6.SetFocus
MsgBox "Please enter the pincode", vbExclamation, Error
Let checkin_form_error = True
ElseIf Text7.Text = "" Then
Text7.SetFocus
MsgBox "Please enter the phone", vbExclamation, Error
Let checkin_form_error = True
ElseIf Combo1.Text = "" Then
Combo1.SetFocus
MsgBox "Please enter the Sex", vbExclamation, Error
Let checkin_form_error = True
ElseIf Combo2.Text = "" Then
Combo2.SetFocus
MsgBox "Please enter the Room Number", vbExclamation, Error
Let checkin_form_error = True
End If
End Sub
Private Sub expireconfirmation()
With Data2
   .RecordSource = "select * from reservation"
   .Refresh
   .Recordset.MoveFirst
    Do Until .Recordset.EOF
    On Error Resume Next
    If .Recordset.Fields(4) < Date Then
        .Recordset.Delete
        End If
        .Recordset.MoveNext
        Loop
        End With
        MsgBox "Expired reservation deleted sucessfuly...", vbInformation, "HMS"
        List1.Text = ""
        List2.Text = ""
        Call ResList
       End Sub
Private Sub optionsearch()
Do Until rs.EOF
If rs.Fields(9) = Text26.Text Then
Text16.Text = rs.Fields(2)
Text17.Text = rs.Fields(5)
Text18.Text = rs.Fields(8)
Text21.Text = rs.Fields(9)
Text19.Text = rs.Fields(0)
val = DateValue(Format(Now, "Short Date")) - rs.Fields(0)
Text28.Text = DateValue(Format(Now, "Short Date")) - rs.Fields(0)
Text22.Text = val + 300
Text23.Text = (10 / 100) * Text22.Text
Text24.Text = (20 / 100) * Text23.Text
val1 = Int(Text22.Text)
val2 = Int(Text23.Text)
val3 = Int(Text24.Text)

Text25.Text = val1 + val2 + val3
Text16.Enabled = True
Text17.Enabled = True
Text18.Enabled = True
Text19.Enabled = True
Text21.Enabled = True
Text22.Enabled = True
Text23.Enabled = True
Text24.Enabled = True
Text25.Enabled = True
Command11.Enabled = True
'Command12.Enabled = True
Command21.Enabled = True
Exit Sub
Else
rs.MoveNext
End If
Loop
MsgBox "no datas found.", vbInformation, "HMS"
Text26.Text = ""
Text26.SetFocus
End Sub
Private Sub rescheck()
Text11 = ""
Text12 = ""
Text13 = ""
Text14 = ""
Check1.Value = 0
rs1.MoveFirst
Do Until rs1.EOF
If rs1.Fields(1) = Text15.Text And rs1.Fields(5) = True Then
Text11.Text = rs1.Fields(1)
Text12.Text = rs1.Fields(2)
Text13.Text = rs1.Fields(3)
Text14.Text = rs1.Fields(4)
Check1.Value = 1
Exit Sub
Else
rs1.MoveNext
End If
Loop
MsgBox "No data found.Try again..", vbInformation, "HMS"
End Sub
Private Sub rescheck1()
rs1.MoveFirst
Do Until rs1.EOF
'If (rs1.Fields(1) = resinput) And rs1.Fields(5) = True Then
If rs1.Fields(5) = True Then
Text11.Text = rs1.Fields(1)
Text12.Text = rs1.Fields(2)
Text13.Text = rs1.Fields(3)
Text14.Text = rs1.Fields(4)
Check1.Value = 1
Command5.Enabled = False
Command7.Enabled = True
Command19.Enabled = True
Command6.Enabled = False
Exit Sub
Else
rs1.MoveNext
End If
Loop
MsgBox "No data found.Try again..", vbOKOnly, "HMS"
End Sub
