VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12915
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   12915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton VIEW_ALL 
      Caption         =   "VIEW ALL"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   4
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton GET_STARTED 
      Caption         =   "GET STARTED"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      DataMember      =   "mobile-phone.jpg"
      Height          =   3615
      Left            =   960
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":B25B
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6480
      TabIndex        =   2
      Top             =   840
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "M-SHOPPING APPLICATION"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Hide
frmSplash.Show

End Sub

