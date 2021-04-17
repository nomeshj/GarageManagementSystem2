VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17715
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   8655
   ScaleWidth      =   17715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
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
      Left            =   10320
      TabIndex        =   4
      Top             =   5520
      Width           =   2500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove Car from garage"
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
      TabIndex        =   3
      Top             =   5400
      Width           =   2500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modify Car"
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
      Left            =   10200
      TabIndex        =   2
      Top             =   3240
      Width           =   2500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add car to garage"
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
      Left            =   5520
      TabIndex        =   1
      Top             =   3240
      Width           =   2500
   End
   Begin VB.Label Label1 
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
      Height          =   1095
      Left            =   4560
      TabIndex        =   0
      Top             =   840
      Width           =   9135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form2.Hide
    Form3.Show
    
End Sub

Private Sub Command2_Click()
    Form2.Hide
    Form4.Show
    
End Sub

Private Sub Command3_Click()
    Form2.Hide
    Form5.Show
    
End Sub

Private Sub Command4_Click()
    End
    
End Sub
