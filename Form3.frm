VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "abc "
   ClientHeight    =   990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1995
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   990
   ScaleWidth      =   1995
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "&Exit"
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "/"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "X"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   495
      Left            =   2520
      TabIndex        =   13
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   495
      Left            =   1320
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   600
      X2              =   960
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   360
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
On Error Resume Next
Label1 = ""
Label2 = ""
Label3 = ""
Label4 = ""
Label5 = ""
Label1 = Text2 * Text4
Label3 = Label1 / Text4 * Text3
Label2 = Label1 / Text2 * Text1
Label4 = Text2 * Text4
Label5 = 0 + Label2 + Label3
MsgBox Label5 + "/" + Label1
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
End Sub

Private Sub Command2_Click()
On Error Resume Next
Label1 = ""
Label2 = ""
Label3 = ""
Label4 = ""
Label5 = ""
Label2 = Text1 * Text3
Label3 = Text2 * Text4
Label5 = 0 + Label2
MsgBox Label5 + "/" + Label3
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
End Sub

Private Sub Command3_Click()
On Error Resume Next
Label1 = ""
Label2 = ""
Label3 = ""
Label4 = ""
Label5 = ""
Label1 = Text2 * Text4
Label3 = Label1 / Text4 * Text3
Label2 = Label1 / Text2 * Text1
Label4 = Text2 * Text4
Label5 = 0 + Label2 - Label3
MsgBox Label5 + "/" + Label1
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
End Sub

Private Sub Command4_Click()
On Error Resume Next
Label1 = ""
Label2 = ""
Label3 = ""
Label4 = ""
Label5 = ""
Label6 = ""
Label6 = Text4
Text4 = Text3
Text3 = Label6
Label1 = Text1 * Text3
Label2 = Text2 * Text4
MsgBox Label1 + "/" + Label2
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
End Sub
