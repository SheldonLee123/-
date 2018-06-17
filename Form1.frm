VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   7035
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "不好"
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "好"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   3270
      Left            =   3960
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "做我女朋友好不好?"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "郭峰硕,我观察你很久了"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_Click()

End Sub

Private Sub Command1_Click()

Form1.Hide
Form2.Show

End Sub

Private Sub Command2_Click()

If a Mod 3 = 0 Then
Form1.Hide
Form3.Show
ElseIf a Mod 3 = 1 Then
Form1.Hide
Form4.Show
Else
Form1.Hide
Form5.Show
End If

End Sub

