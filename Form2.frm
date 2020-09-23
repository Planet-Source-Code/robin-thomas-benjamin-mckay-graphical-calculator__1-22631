VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "About Calculator"
   ClientHeight    =   1275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5445
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1275
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Feedback"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Fax"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&E-Mail"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Many thanks to: Chris Seelbach"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Calculator 1.4 - by Robin McKay"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Shell ("start mailto:ian@imckay.fsnet.co.uk")
End Sub

Private Sub Command3_Click()
MsgBox "My fax number is: 0870-136-4618"
End Sub

Private Sub Command4_Click()
MsgBox "Any feedback you have is greatly appreciated. Please leave comments at the Planet Source Code website. Thank you for using Calculator 1.2." + vbCrLf + "Division By Zero bug has been fixed. Thank you to Chris Selbach for pointing that one out. Now when you divide by zero, you should see a message saying division by zero is not possible.", vbInformation, "A big thank you"
End Sub

