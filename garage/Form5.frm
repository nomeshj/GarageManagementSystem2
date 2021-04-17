VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H80000002&
   Caption         =   "Form5"
   ClientHeight    =   4740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11040
   LinkTopic       =   "Form5"
   ScaleHeight     =   4740
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   5040
      TabIndex        =   5
      Top             =   3360
      Width           =   2500
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\SE\Desktop\garage\Database4.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Table1"
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove Car"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   1320
      TabIndex        =   3
      Top             =   3360
      Width           =   2500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<- Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   5040
      TabIndex        =   2
      Top             =   1680
      Width           =   2500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next ->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   1440
      TabIndex        =   1
      Top             =   1680
      Width           =   2500
   End
   Begin VB.TextBox Text1 
      DataField       =   "car_name"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Car name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Data1.Recordset.MoveNext
    
End Sub

Private Sub Command2_Click()
    Data1.Recordset.MovePrevious
    
End Sub

Private Sub Command3_Click()
    Data1.Recordset.Delete
    MsgBox ("Car remove successfully")
End Sub
    
Private Sub Command4_Click()
    Form5.Hide
    Form2.Show
    
End Sub
