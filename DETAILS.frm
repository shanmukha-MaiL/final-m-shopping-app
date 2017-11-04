VERSION 5.00
Begin VB.Form DETAILS 
   Caption         =   "Form1"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10095
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "BACK"
      Height          =   615
      Left            =   360
      TabIndex        =   13
      Top             =   6840
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NEXT"
      Height          =   645
      Left            =   5880
      TabIndex        =   12
      Top             =   6840
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLEAR"
      Height          =   615
      Left            =   3120
      TabIndex        =   11
      Top             =   6840
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Height          =   765
      Left            =   3480
      TabIndex        =   9
      Top             =   5640
      Width           =   3975
   End
   Begin VB.TextBox PIN 
      Height          =   765
      Left            =   3480
      TabIndex        =   7
      Top             =   4560
      Width           =   3975
   End
   Begin VB.TextBox ADD2 
      Height          =   765
      Left            =   3480
      TabIndex        =   5
      Top             =   3600
      Width           =   3975
   End
   Begin VB.TextBox ADD1 
      Height          =   765
      Left            =   3480
      TabIndex        =   3
      Top             =   2760
      Width           =   3975
   End
   Begin VB.TextBox NAME11 
      Height          =   765
      Left            =   3480
      TabIndex        =   1
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Label Label6 
      Caption         =   "CONTACT NO."
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "PINCODE"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   4680
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "CITY"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "ADRESS"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "NAME"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PERSONAL DETAILS"
      Height          =   735
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "DETAILS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NAME1 As String
Dim ADD11 As String
Dim ADD21 As String
Dim PIN1 As String
Dim PNO1 As String
Dim ADDRESS As String


Private Sub Command1_Click()
NAME11.Text = ""
ADD1.Text = ""
ADD2.Text = ""
PIN.Text = ""
PNO.Text = ""
End Sub

Private Sub Command2_Click()
Dim MSG As Integer

NAME1 = NAME11.Text
ADD11 = ADD1.Text
ADD21 = ADD2.Text
PIN1 = PIN.Text

ADDRESS = ADD11 + "," + ADD21 + "," + PIN1 + "."

If NAME11.Text = "" Then
MSG = MsgBox("OOPS!! Name Field Cant be Empty", vbOKOnly)


ElseIf ADD1.Text = "" Then
MSG = MsgBox("OOPS!! Address Field Cant be Empty", vbOKOnly)


ElseIf ADD2.Text = "" Then
    MSG = MsgBox("OOPS!! City Field Cant be Empty", vbOKOnly)

ElseIf PIN.Text = "" Then
    MSG = MsgBox("OOPS!! Pincode Field Cant be Empty", vbOKOnly)
       
Else


    
DETAILS.Hide
BILLING.Show
BILLING.Label12.Caption = NAME1
BILLING.Label13.Caption = ADDRESS
End If

End Sub

Private Sub Command3_Click()
DETAILS.Hide
Combo.Show
End Sub
