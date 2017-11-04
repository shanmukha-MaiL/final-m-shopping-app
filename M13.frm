VERSION 5.00
Begin VB.Form REDMI_Y1 
   Caption         =   "Form13"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9315
   LinkTopic       =   "Form13"
   ScaleHeight     =   7755
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\shanmukha\Downloads\m-shopping\mobileData.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "MOBILE_DATA"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      TabIndex        =   3
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6360
      TabIndex        =   2
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BUY"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      TabIndex        =   1
      Top             =   6600
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   2900
      Left            =   120
      Picture         =   "M13.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label Label12 
      Caption         =   "AVAILABLE MOBILES:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   6480
      Width           =   3255
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      DataField       =   "MobileAvailable"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   14
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "Best Under 10000 Budget"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7080
      TabIndex        =   13
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "QUANTITY"
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
      Left            =   5160
      TabIndex        =   12
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "8000 RS ONLY"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   5880
      Width           =   3015
   End
   Begin VB.Label Label7 
      Caption         =   "GOLD COLOR"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Label Label6 
      Caption         =   "2GB RAM"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label Label5 
      Caption         =   "35 Hrs Battery Life"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Real Time Beautification"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "MIUI 9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "WHY BUY?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "REDMI Y1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "REDMI_Y1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim quantity As Integer
quantity = Val(Text1.Text)
Dim available As Integer
 available = Val(Label11.Caption)
Dim DISPLAY As Integer
If quantity <= available Then
    DISPLAY = MsgBox("Click ok to confirm the purchase ?", vbOKCancel + vbQuestion)
        If DISPLAY = vbCancel Then
            Form1.Show
        End If
        If DISPLAY = vbOK Then
            available = available - quantity
            Label11.Caption = Val(available)
            Form1.Hide
            Form2.Show
            End If
            
End If
If quantity > available Then
    DISPLAY = MsgBox("Entered quantity more than available.", vbOK)
        If DISPLAY = vbOK Then
             Text1.Text = ""
             Form1.Show
             
        End If
End If
End Sub

