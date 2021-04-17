VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H8000000D&
   Caption         =   "Form4"
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18180
   FillColor       =   &H80000003&
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   9420
   ScaleWidth      =   18180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
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
      Left            =   15000
      TabIndex        =   32
      Top             =   7920
      Width           =   2500
   End
   Begin VB.CheckBox Check10 
      BackColor       =   &H80000018&
      Caption         =   "Roof tracks"
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
      Left            =   10200
      TabIndex        =   26
      Top             =   6360
      Width           =   2295
   End
   Begin VB.CheckBox Check9 
      BackColor       =   &H80000018&
      Caption         =   "New car lights"
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
      Left            =   10200
      TabIndex        =   25
      Top             =   5400
      Width           =   2415
   End
   Begin VB.CheckBox Check8 
      BackColor       =   &H80000018&
      Caption         =   "Parking sensor"
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
      Left            =   10200
      TabIndex        =   24
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H80000018&
      Caption         =   "Navigation system"
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
      Left            =   10320
      TabIndex        =   23
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H80000018&
      Caption         =   "Air conditioning"
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
      Left            =   10320
      TabIndex        =   22
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H80000018&
      Caption         =   "Free flow exhaust system"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   14
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H80000018&
      Caption         =   "Air filter change"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   13
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H80000018&
      Caption         =   "Transmission upgrade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   12
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "edit"
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
      Left            =   7920
      TabIndex        =   11
      Top             =   7920
      Width           =   2500
   End
   Begin VB.TextBox Text2 
      DataField       =   "price"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   10
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
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
      Left            =   11520
      TabIndex        =   9
      Top             =   7920
      Width           =   2500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   8
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   720
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\SE\Desktop\garage\Database4.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   12240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Table1"
      Top             =   720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total"
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
      Left            =   720
      TabIndex        =   6
      Top             =   7920
      Width           =   2500
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H80000018&
      Caption         =   "Supercharging the engine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   4
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000018&
      Caption         =   "Brakes upgrade"
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
      Left            =   960
      TabIndex        =   2
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      DataField       =   "cust_name"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000018&
      Caption         =   "250"
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
      Left            =   13800
      TabIndex        =   31
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000018&
      Caption         =   "500"
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
      Left            =   13920
      TabIndex        =   30
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000018&
      Caption         =   "400"
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
      Left            =   13920
      TabIndex        =   29
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000018&
      Caption         =   "150"
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
      Left            =   13920
      TabIndex        =   28
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000018&
      Caption         =   "220"
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
      Left            =   13920
      TabIndex        =   27
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000002&
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13800
      TabIndex        =   21
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000002&
      Caption         =   "Functional Modification"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9840
      TabIndex        =   20
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000002&
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   19
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000002&
      Caption         =   "Performance Modification"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      TabIndex        =   18
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000018&
      Caption         =   "250"
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
      Left            =   4560
      TabIndex        =   17
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000018&
      Caption         =   "160"
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
      Left            =   4560
      TabIndex        =   16
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000018&
      Caption         =   "300"
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
      Left            =   4560
      TabIndex        =   15
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000018&
      Caption         =   "500"
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
      Left            =   4560
      TabIndex        =   5
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000018&
      Caption         =   "200"
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
      Left            =   4560
      TabIndex        =   3
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Customer Name"
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
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Dim Total  As Integer

Private Sub Command1_Click()
  
    
    If Check1.Value = 1 Then
        Total = Total + 200
    End If
    If Check2.Value = 1 Then
        Total = Total + 500
    End If
    If Check3.Value = 1 Then
        Total = Total + 300
    End If
    If Check4.Value = 1 Then
        Total = Total + 160
    End If
    If Check5.Value = 1 Then
        Total = Total + 250
    End If
    If Check6.Value = 1 Then
        Total = Total + 220
    End If
    If Check7.Value = 1 Then
        Total = Total + 150
    End If
    If Check8.Value = 1 Then
        Total = Total + 400
    End If
    If Check9.Value = 1 Then
        Total = Total + 500
    End If
    If Check10.Value = 1 Then
        Total = Total + 250
    End If
    
    Text2.Text = Total
End Sub

Private Sub Command2_Click()
    Data1.Recordset.MoveNext
    
End Sub

Private Sub Command3_Click()
    Data1.Recordset.MovePrevious
    
End Sub

Private Sub Command4_Click()
    Data1.Recordset.Update
    
End Sub

Private Sub Command5_Click()
    If Total = 0 Then
    MsgBox ("Calculate Total first")
    Else
    Data1.Recordset.Edit
    End If
    
End Sub

Private Sub Command6_Click()
    Form4.Hide
    Form2.Show
    
End Sub
