VERSION 5.00
Begin VB.Form Combo 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   5
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   4
      Top             =   3600
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Text            =   "Select"
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "SELECT MOBILE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Please Purchase a mobile phone.THANKYOU"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "MOBILE PURCHASING PAGE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "Combo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1.Text = "IPHONE 7" Then
 Combo.Hide
 IPHONE_7.Show
End If
If Combo1.Text = "IPHONE 8" Then
 Combo.Hide
 IPHONE_8.Show
End If
If Combo1.Text = "IPHONE X BLACK" Then
 Combo.Hide
 IPHONE_X.Show
End If
If Combo1.Text = "IPHONE X WHITE" Then
 Combo.Hide
 IPHONE_X_WH.Show
End If
If Combo1.Text = "JIO PHONE" Then
 Combo.Hide
 JIO.Show
End If
If Combo1.Text = "MI A1" Then
 Combo.Hide
 MI_A1.Show
End If
If Combo1.Text = "MI 5I" Then
 Combo.Hide
 MI_5I.Show
End If
If Combo1.Text = "OPPO F1" Then
 Combo.Hide
 OPPO_F1.Show
End If
If Combo1.Text = "OPPO F5" Then
 Combo.Hide
 OPPO_F5.Show
End If
If Combo1.Text = "REDMI 3S PRIME" Then
 Combo.Hide
 REDMI_PRIME.Show
End If
If Combo1.Text = "REDMI NOTE 4 GOLD" Then
 Combo.Hide
 REDMI_NOTE4.Show
End If
If Combo1.Text = "REDMI NOTE 4 BLUE" Then
 Combo.Hide
 REDMI_NOTE4_B.Show
End If
If Combo1.Text = "REDMI Y1" Then
 Combo.Hide
 REDMI_Y1.Show
End If
If Combo1.Text = "SAMSUNG S8 WHITE" Then
 Combo.Hide
 SAMSUNG_S8.Show
End If
If Combo1.Text = "SAMSUNG S8 BLACK" Then
 Combo.Hide
 SAMSUNG_S8_B.Show
End If
End Sub

Private Sub Form_Load()
Combo1.AddItem "IPHONE 7"
Combo1.AddItem "IPHONE 8"
Combo1.AddItem "IPHONE X BLACK"
Combo1.AddItem "IPHONE X WHITE"
Combo1.AddItem "JIO PHONE"
Combo1.AddItem "MI A1"
Combo1.AddItem "MI 5I"
Combo1.AddItem "OPPO F1"
Combo1.AddItem "OPPO F5"
Combo1.AddItem "REDMI 3S PRIME"
Combo1.AddItem "REDMI NOTE 4 GOLD"
Combo1.AddItem "REDMI NOTE 4 BLUE"
Combo1.AddItem "REDMI Y1"
Combo1.AddItem "SAMSUNG S8 WHITE"
Combo1.AddItem "SAMSUNG S8 BLACK"

Dim DISPLAY As Integer
DISPLAY = MsgBox("This is Mobile Purchasing Page , Do you wish to Continue ??", vbYesNo)
If DISPLAY = vbNo Then
    Combo.Hide
    LOGIN.Show
    End If
    
End Sub
