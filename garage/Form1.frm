VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16125
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8160
   ScaleWidth      =   16125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
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
      Left            =   8880
      TabIndex        =   5
      Top             =   5880
      Width           =   2500
   End
   Begin VB.TextBox TextBox2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   7800
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   4440
      Width           =   2500
   End
   Begin VB.TextBox TextBox1 
      Height          =   375
      Left            =   7800
      TabIndex        =   3
      Top             =   3360
      Width           =   2500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
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
      Left            =   4200
      TabIndex        =   2
      Top             =   5880
      Width           =   2500
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Garage Management System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   6
      Top             =   1440
      Width           =   9855
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
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
      Left            =   5040
      TabIndex        =   1
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
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
      Left            =   5040
      TabIndex        =   0
      Top             =   3360
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If TextBox1.Text = "ninad" And TextBox2.Text = "ninad123" Then
            Form1.Hide
            Form2.Show
        ElseIf TextBox1.Text = "mayur" And TextBox2.Text = "mayur123" Then
            Form1.Hide
            Form2.Show
        Else
            MsgBox ("Invalid Login details")
        End If
End Sub

Private Sub Command2_Click()
    End
    
End Sub
