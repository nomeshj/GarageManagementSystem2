VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000002&
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13650
   LinkTopic       =   "Form3"
   ScaleHeight     =   5535
   ScaleWidth      =   13650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   9840
      TabIndex        =   10
      Top             =   4200
      Width           =   2500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   1200
      TabIndex        =   9
      Top             =   4200
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
      Left            =   9840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Table1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   5640
      TabIndex        =   8
      Top             =   4200
      Width           =   2500
   End
   Begin VB.TextBox TextBox4 
      DataField       =   "mobile_no"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox TextBox3 
      DataField       =   "car_no"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox TextBox2 
      DataField       =   "car_name"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox TextBox1 
      DataField       =   "cust_name"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Mobile No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Car no"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Car name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Data1.Recordset.Update
 MsgBox ("Saved Successfully")
End Sub


Private Sub Command2_Click()
    Data1.Recordset.AddNew
    TextBox1.SetFocus
    
End Sub

Private Sub Command3_Click()
    Form3.Show
    Form2.Show
    
End Sub
