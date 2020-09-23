VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Graphical Calculator"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   3615
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command54 
      BackColor       =   &H00008080&
      Caption         =   "Clear Graph"
      Height          =   255
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox SubYinput 
      Height          =   285
      Left            =   2400
      TabIndex        =   71
      Text            =   "50"
      Top             =   880
      Width           =   375
   End
   Begin VB.TextBox SubXinput 
      Height          =   285
      Left            =   2400
      TabIndex        =   70
      Text            =   "50"
      Top             =   440
      Width           =   375
   End
   Begin VB.CommandButton Command53 
      BackColor       =   &H00008080&
      Caption         =   "Plot"
      Height          =   255
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command52 
      BackColor       =   &H00400040&
      Caption         =   "Power Off"
      Height          =   400
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   5565
      Width           =   615
   End
   Begin VB.PictureBox Graph 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      ForeColor       =   &H000000FF&
      Height          =   1800
      Left            =   120
      ScaleHeight     =   1740
      ScaleWidth      =   1740
      TabIndex        =   67
      Top             =   120
      Width           =   1800
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00404000&
      Caption         =   "Radians"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   1200
      TabIndex        =   66
      Top             =   2400
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00004000&
      Caption         =   "Degrees"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   120
      TabIndex        =   65
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command51 
      BackColor       =   &H00C0C0C0&
      Caption         =   "abc"
      Height          =   255
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   5640
      Width           =   495
   End
   Begin VB.CommandButton Command50 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tan"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton Command49 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cos"
      Height          =   255
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton Command48 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sin"
      Height          =   255
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton Command44 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hyp"
      Height          =   255
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton Command43 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sgn"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   5640
      Width           =   495
   End
   Begin VB.CommandButton Command42 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rnd"
      Height          =   255
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command47 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fix"
      Height          =   255
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command46 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Int"
      Height          =   255
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton Command45 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Abs"
      Height          =   255
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   5640
      Width           =   495
   End
   Begin VB.CommandButton Command41 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Dec"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton Command40 
      BackColor       =   &H00C0C0C0&
      Caption         =   "x4"
      Height          =   255
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   4200
      Width           =   495
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   600
      Top             =   2280
   End
   Begin VB.CommandButton Command39 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mod"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command38 
      BackColor       =   &H00C0C0C0&
      Caption         =   "1/x"
      Height          =   255
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command37 
      BackColor       =   &H00C0C0C0&
      Caption         =   "0."
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command36 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CE"
      Height          =   255
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   3120
      Width           =   495
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   840
      Top             =   2280
   End
   Begin VB.CommandButton Command35 
      BackColor       =   &H00E0E0E0&
      Caption         =   "+/-"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Command34 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Diam"
      Height          =   255
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton Command33 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Radi"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton Command32 
      BackColor       =   &H00C0C0C0&
      Caption         =   "n!"
      Height          =   255
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton Command31 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Not"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton Command30 
      BackColor       =   &H00C0C0C0&
      Caption         =   "log"
      Height          =   255
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton Command29 
      BackColor       =   &H00C0C0C0&
      Caption         =   "%"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton Command28 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Exp"
      Height          =   255
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton Command27 
      BackColor       =   &H00C0C0C0&
      Caption         =   "sA"
      Height          =   255
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton Command26 
      BackColor       =   &H00C0C0C0&
      Caption         =   "sP"
      Height          =   255
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2760
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1320
      Top             =   2280
   End
   Begin VB.CommandButton Command25 
      BackColor       =   &H00C0C0C0&
      Caption         =   "XoR"
      Height          =   255
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CTK"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H00C0C0C0&
      Caption         =   "cC"
      Height          =   255
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H00C0C0C0&
      Caption         =   "cA"
      Height          =   255
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pi"
      Height          =   255
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00C0C0C0&
      Caption         =   "x3"
      Height          =   255
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00C0C0C0&
      Caption         =   "x2"
      Height          =   255
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Sqrt"
      Height          =   255
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00C0C0C0&
      Caption         =   "C"
      Height          =   255
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00E0E0E0&
      Caption         =   "."
      Height          =   255
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      Height          =   255
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00E0E0E0&
      Caption         =   "/"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00004000&
      Caption         =   "="
      Height          =   255
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00E0E0E0&
      Caption         =   "X"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "-"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "+"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      Height          =   255
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "8"
      Height          =   255
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "7"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "6"
      Height          =   255
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "5"
      Height          =   255
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "4"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "3"
      Height          =   255
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "2"
      Height          =   255
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Y = Y +               * Sin(b * 3 * 20)"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1920
      TabIndex        =   73
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "X = X +               * Cos(a * 3 * 2)"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1920
      TabIndex        =   72
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label13 
      Caption         =   "0"
      Height          =   495
      Left            =   840
      TabIndex        =   52
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "0"
      Height          =   495
      Left            =   1560
      TabIndex        =   47
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "0"
      Height          =   495
      Left            =   840
      TabIndex        =   46
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "0"
      Height          =   495
      Left            =   840
      TabIndex        =   45
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   495
      Left            =   600
      TabIndex        =   44
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   495
      Left            =   840
      TabIndex        =   38
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   960
      TabIndex        =   36
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2400
      TabIndex        =   31
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "The computation has been added to memory"
      Height          =   375
      Left            =   0
      TabIndex        =   29
      Top             =   3360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "The Answer Is:"
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   5280
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   4920
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Clear 
         Caption         =   "&Clear"
         Shortcut        =   {F6}
      End
      Begin VB.Menu border2 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
      Begin VB.Menu Copy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu Paste 
         Caption         =   "&Paste"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu HelpTopics 
         Caption         =   "&Help Topics"
      End
      Begin VB.Menu border1 
         Caption         =   "-"
      End
      Begin VB.Menu AboutCalculator 
         Caption         =   "&About Calculator"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Calculator -- Written By Robin McKay
Const Pi As Variant = "3.1415926535897932384626433832795"
Dim n   As Integer
Dim Sum As Variant
Dim dS  As Variant
Dim P1  As Variant
Dim X   As Variant
Dim Lim As Variant
Dim Q As Variant
Dim a As Variant
Dim i  As Integer
Dim Px As Variant
Dim T  As Variant
Dim TicColor As Long
Dim SubX As Long
Dim SubY As Long



  

Private Sub AboutCalculator_Click()
Form2.Show
End Sub



Private Sub Clear_Click()
Call Command17_Click
End Sub

Private Sub Command1_Click()
Text1.SelText = CDec(1)
End Sub

Private Sub Command10_Click()
On Error Resume Next
If Label1 = 0 Then
    Label1 = 1
    Label2 = 0 + Label2 + Text1
    Text1 = ""
    Label1 = 0
ElseIf Label1 = 1 Then
    Label1 = 0
    Label2 = Label2 - Text1
    Text1 = ""
ElseIf Label1 = 2 Then
    Label2 = Label2 * Text1
    Label1 = 0
    Text1 = ""
ElseIf Label1 = 3 Then
    Label2 = Label2 / Text1
    Label1 = 0
    Text1 = ""
End If
End Sub

Private Sub Command11_Click()
On Error Resume Next
If Label1 = 0 Then
    Label2 = 0 + Label2 + Text1
    Label1 = 1
    Text1 = ""
    Label1 = 1
ElseIf Label1 = 2 Then
    Label2 = Label2 * Text1
    Label1 = 1
    Text1 = ""
ElseIf Label1 = 1 Then
    Label2 = Label2 - Text1
    Text1 = ""
ElseIf Label1 = 3 Then
    Label2 = Label2 / Text1
    Label1 = 1
    Text1 = ""
End If
End Sub

Private Sub Command12_Click()
On Error Resume Next
If Label1 = 0 Then
    Label2 = 0 + Label2 + Text1
    Label1 = 2
    Text1 = ""
ElseIf Label1 = 1 Then
    Label2 = Label2 - Text1
    Label1 = 2
    Text1 = ""
ElseIf Label1 = 2 Then
    Label2 = Label2 * Text1
    Text1 = ""
ElseIf Label1 = 3 Then
    Label2 = Label2 / Text1
    Label1 = 2
    Text1 = ""
End If
End Sub

Private Sub Command13_Click()
On Error Resume Next
If (Label1 = 3) And (Text1 = 0) Then
    Text1 = "ERROR: Cannot divide by zero"
    'Label5.Visible = True
    Timer3.Enabled = True
    'Label12 = 1
    GoTo 10
    Else
End If
Dim i As Integer
If Label1 = 0 Then
    Label2 = 0 + Label2 + Text1
    Label3 = Label2
    Label2 = 0
    Label1 = 0
    Text1 = Label3.Caption
ElseIf Label1 = 1 Then
    Label2 = Label2 - Text1
    Label3 = Label2
    Label1 = 0
    Label2 = 0
    Text1 = Label3.Caption
ElseIf Label1 = 2 Then
    Label2 = Label2 * Text1
    Label3 = Label2
    Label2 = 0
    Label1 = 0
    Text1 = Label3.Caption
ElseIf Label1 = 3 Then
    Label2 = CDec(Label2) / CDec(Text1)
    Label3 = Label2
    Label1 = 0
    Label2 = 0
    Text1 = Label3.Caption
End If
10
Label8 = 1
End Sub

Private Sub Command14_Click()
On Error Resume Next
If Label1 = 0 Then
    Label2 = 0 + Label2 + Text1
    Label1 = 3
    Text1 = ""
ElseIf Label1 = 1 Then
    Label2 = Label2 - Text1
    Label1 = 3
    Text1 = ""
ElseIf Label1 = 2 Then
    Label2 = Label2 * Text1
    Label1 = 3
    Text1 = ""
ElseIf Label1 = 3 Then
    Label2 = Label2 / Text1
    Text1 = ""
End If
End Sub

Private Sub Command15_Click()
' Handles division by Zero
Text1.SelText = 0
'If (Label1 = 3) And (Text1 = 0) Then
    'Label5.Caption = "Cannot divide by Zero"
    'Label5.Visible = True
    'Timer3.Enabled = True
    'Label12 = 1
'End If
End Sub

Private Sub Command16_Click()
Text1.SelText = "."
End Sub

Private Sub Command17_Click()
On Error Resume Next
If Label12 = 1 Then
    Label1 = 0
    Label3 = 0
    Label2 = 0
    Label10 = 0
    Text1 = ""
ElseIf Label12 = 0 Then
    Label1 = 0
    Label2 = 0
    Label3 = 0
    Label10 = 0
    Text1 = ""
End If
End Sub

Private Sub Command18_Click()
On Error Resume Next
Text1 = Sqr(Text1)
End Sub

Private Sub Command19_Click()
Text1 = Text1 * Text1
End Sub

Private Sub Command2_Click()
Text1.SelText = CDec(2)
End Sub

Private Sub Command20_Click()
Text1 = Text1 * Text1 * Text1
End Sub

Private Sub Command21_Click()
Text1.Text = Pi
End Sub



Private Sub Command22_Click()
On Error Resume Next
Text1 = Pi * Text1 * Pi * Text1
End Sub

Private Sub Command23_Click()
On Error Resume Next
Text1 = 2 * Pi * Text1
End Sub

Private Sub Command24_Click()
On Error Resume Next
Text1 = Text1 * 1.6
End Sub

Private Sub Command25_Click()
On Error GoTo cancelerr:
Dim i As Integer
Dim i2 As Integer
Dim i3 As Integer
    i = InputBox("Please enter the highest computation value:", "Highest")
    i2 = InputBox("Please enter the lowest computation value:", "Lowest")
    i3 = i - i2
    Text1 = "The XOR is equal to:"
    'Label5.Visible = True
    Label6.Visible = True
    Label6 = i3
    Timer2.Enabled = True
cancelerr:
Exit Sub
End Sub



Private Sub Command26_Click()
Text1 = 0 + Text1 + Text1 + Text1 + Text1
End Sub

Private Sub Command27_Click()
Text1 = Text1 * Text1
End Sub

Private Sub Command28_Click()
On Error Resume Next
Text1.SelText = ".e" + "+" + "0"
End Sub

Private Sub Command29_Click()
On Error Resume Next
' This performs a percentage calculation
Text1 = Text1 / 100
Text1 = Text1 * 100
Text1.Text = Text1.Text + "%"
End Sub

Private Sub Command3_Click()
Text1.SelText = CDec(3)
End Sub
Private Sub Command30_Click()
On Error Resume Next
log10 = Log(Text1) / Log(10)
Text1 = log10
End Sub

Private Sub Command31_Click()
On Error Resume Next
If Label10 = 0 Then
    Label10 = Text1.Text
    Label11 = 1
    Text1.Text = "-" + Text1.Text - 1
ElseIf Label11 = 1 Then
    Text1.Text = Label10
    Label11 = 0
    Label10 = 0
End If
End Sub

Private Sub Command32_Click()
' This program supports factorial 10
If Text1 = 1 Then Text1 = 1
If Text1 = 2 Then Text1 = 2
If Text1 = 3 Then Text1 = 6
If Text1 = 4 Then Text1 = 24
If Text1 = 5 Then Text1 = 120
If Text1 = 6 Then Text1 = 720
If Text1 = 7 Then Text1 = 5040
If Text1 = 8 Then Text1 = 40320
If Text1 = 9 Then Text1 = 362880
If Text1 = 10 Then Text1 = 3628800
End Sub

Private Sub Command33_Click()
On Error Resume Next
Text1 = Text1 / 2
End Sub

Private Sub Command34_Click()
On Error Resume Next
Text1 = Text1 * 2
End Sub





Private Sub Command35_Click()
' Use this to toggle plus and minus
On Error Resume Next
If Label9.Caption = 0 Then
    Text1 = Text1 - Text1 - Text1
    Label9 = 1
ElseIf Label9 = 1 Then
    Text1 = Text1 - Text1 - Text1
    Label9 = 0
End If
End Sub

Private Sub Command36_Click()
On Error Resume Next
Text1 = ""
End Sub

Private Sub Command37_Click()
On Error Resume Next
'we can't allow "0." to be entered
'more than once
If Text1 Like "*.*" Then Exit Sub
Text1.SelText = "0."
End Sub



Private Sub Command38_Click()
On Error Resume Next
Text1 = Text1 / Text1 / Text1
End Sub



Private Sub Command39_Click()
On Error Resume Next
Dim i As Integer
Dim i2 As Integer
Dim i3 As Integer
Label13 = 0
    i = InputBox("Please enter the 1st number:", "Number 1")
    i2 = InputBox("Please enter the 2nd number:", "Number 2")
    i3 = i Mod i2
    Label13 = i3
    Text1 = "The MOD is " + Label13
    'Label5.Visible = True
    Label13 = ""
    Timer4.Enabled = True
End Sub

Private Sub Command4_Click()
Text1.SelText = CDec(4)
End Sub



Private Sub Command40_Click()
On Error Resume Next
Text1 = Text1 * Text1 * Text1 * Text1
End Sub

Private Sub Command41_Click()
' A procedure to convert a number to decimal
On Error Resume Next
Text1 = Text1 / 100
End Sub



Private Sub Command42_Click()
On Error Resume Next
Randomize
Text1 = Rnd(Text1)
End Sub

Private Sub Command43_Click()
On Error Resume Next
Text1 = Sgn(Text1)
End Sub

Private Sub Command44_Click()
On Error GoTo cancelerr:
Dim i As Integer
Dim i2 As Integer
Dim i3 As Integer
Dim i4 As Integer
Dim i5 As Integer
    i = InputBox("Please enter number:", "1st Number")
    i2 = InputBox("Please enter number:", "2nd Number")
    i3 = i * i
    i4 = i2 * i2
    i5 = i3 + i4
    Text1 = Sqr(i5)
cancelerr:
Exit Sub
End Sub

Private Sub Command45_Click()
On Error Resume Next
Text1 = Abs(Text1)
End Sub

Private Sub Command46_Click()
On Error Resume Next
Text1 = Int(Text1)
End Sub

Private Sub Command47_Click()
On Error Resume Next
Text1 = Fix(Text1)
End Sub












Private Sub Command48_Click()
' This works out the sin of a number
On Error Resume Next


' Check for non-numeric input argument
  If IsNumeric(Text1) = False Then
     Text1 = "ERROR: Invalid input argument"
     Beep
     Exit Sub
  End If

'If Option 1 = true then convert degrees to radians
  If Option1.Value = True Then
  Q = CDec(Text1) * Pi / 180
Else
'display in radians
Q = CDec(Text1)
End If

' Compute sine of argument
  a = SinRad(Q)
  
  Text1 = a
End Sub

Private Sub Command49_Click()
' This works out the cosine of a number
On Error Resume Next


' Check for non-numeric input argument
  If IsNumeric(Text1) = False Then
     Text1 = "ERROR: Invalid input argument"
     Beep
     Exit Sub
  End If

'If Option 1 = true then convert degrees to radians
If Option1.Value = True Then
  Q = CDec(Text1) * Pi / 180
Else
'display in radians
Q = CDec(Text1)
End If
' Compute sine of argument
  a = CosRad(Q)
  
  Text1 = a
  
End Sub

Private Sub Command5_Click()
Text1.SelText = CDec(5)
End Sub



Private Sub Command50_Click()
' This works out the tangent of a number
On Error Resume Next


' Check for non-numeric input argument
  If IsNumeric(Text1) = False Then
     Text1 = "ERROR: Invalid input argument"
     Beep
     Exit Sub
  End If

' Check for infinite result
  If Abs(Text1) = 90 Or Abs(Text1) = 270 Then
     Text1 = "Infinite result"
     Exit Sub
  End If

'If Option 1 = true then convert degrees to radians
  If Option1.Value = True Then
  Q = CDec(Text1) * Pi / 180
Else
'display in radians
Q = CDec(Text1)
End If

' Compute tangent of argument
  a = TanRad(Q)
  
  Text1 = a
End Sub

Private Sub Command51_Click()
Form3.Show
End Sub

Private Sub Command52_Click()
Exit_Click
End Sub

Private Sub Command53_Click()
'set graph window to pixels
Graph.ScaleMode = 3
Dim a, b, RandomColor
'starting point on graph window
X = 0
Y = 50
'pick a random color
RandomColor = QBColor(Rnd * 14) + 1
'draw 1000 points of a sinewave
For i = 1 To 1000
a = a - 1
b = b + 1
X = X + SubX * Cos(a * 3 * 2)
Y = Y + SubY * Sin(b * 3 * 20)
Graph.PSet (X, Y), RandomColor
Next i
End Sub

Private Sub Command54_Click()
Graph.Cls
redrawgraph
End Sub

Private Sub Command6_Click()
Text1.SelText = CDec(6)
End Sub

Private Sub Command7_Click()
Text1.SelText = CDec(7)
End Sub

Private Sub Command8_Click()
Text1.SelText = CDec(8)
End Sub

Private Sub Command9_Click()
Text1.SelText = CDec(9)
End Sub

Private Sub Degrees_Click()
On Error Resume Next
If Degrees.Checked = False Then
    Degrees.Checked = True
    Radians.Checked = False
End If
End Sub

Private Sub Copy_Click()
On Error Resume Next
Clipboard.SetText Text1.Text
End Sub

Private Sub Exit_Click()
On Error Resume Next
Unload Form1
Unload Form2
Set Form1 = Nothing
Set Form2 = Nothing
End Sub

Private Sub Form_Activate()
'set the degree option button to true
Option2 = True
End Sub

Private Sub Form_Load()
On Error Resume Next
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Label1.Visible = False
Label6.Visible = False
'Radians.Checked = True
SetForeColor Command13, vbGreen
SetForeColor Command1, vbBlue
SetForeColor Command2, vbBlue
SetForeColor Command3, vbBlue
SetForeColor Command4, vbBlue
SetForeColor Command5, vbBlue
SetForeColor Command6, vbBlue
SetForeColor Command7, vbBlue
SetForeColor Command8, vbBlue
SetForeColor Command9, vbBlue
SetForeColor Command15, vbBlue
SetForeColor Command16, vbBlue
SetForeColor Command35, vbBlue
SetForeColor Command48, &HC000C0
SetForeColor Command49, &HC000C0
SetForeColor Command50, &HC000C0
SetForeColor Command10, vbRed
SetForeColor Command11, vbRed
SetForeColor Command12, vbRed
SetForeColor Command14, vbRed
SetForeColor Command52, vbRed
SetForeColor Command53, vbYellow
   
'draw the grid in the graph window
Graph.Line (Graph.Width / 2, 0)-(Graph.Width / 2, Graph.Height), QBColor(7)
Graph.Line (0, Graph.Height / 2)-(Graph.Width, Graph.Height / 2), QBColor(7)
TicColor = QBColor(7)
For i = 0 To Graph.Width Step 45
'horizontal tics
If i = 450 Or i = 1350 Then
TicColor = vbRed
Else
TicColor = QBColor(7)
End If
Graph.Line (i, (Graph.Height / 2) + 35)-(i, (Graph.Height / 2) - 35), TicColor
Next i
For i = 0 To Graph.Height Step 45
'vertical tics
If i = 450 Or i = 1350 Then
TicColor = vbRed
Else
TicColor = QBColor(7)
End If
Graph.Line (Graph.Width / 2 + 35, i)-(Graph.Width / 2 - 35, i), TicColor
Next i
'print "horizontal +10"
Graph.CurrentX = (Graph.Width / 100) * 20
Graph.CurrentY = (Graph.Height / 100) * 35
Graph.Print "10"
'print "vertical +10"
Graph.CurrentX = (Graph.Width / 100) * 35
Graph.CurrentY = (Graph.Height / 100) * 20
Graph.Print "10"
'print "horizontal -10"
Graph.CurrentX = (Graph.Width / 100) * 70
Graph.CurrentY = (Graph.Height / 100) * 53
Graph.Print "-10"
'print "vertical -10"
Graph.CurrentX = (Graph.Width / 100) * 53
Graph.CurrentY = (Graph.Height / 100) * 70
Graph.Print "-10"
'load subX and subY input values
SubX = SubXinput.Text
SubY = SubYinput.Text

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
UnsetForeColor Command13
UnsetForeColor Command1
UnsetForeColor Command2
UnsetForeColor Command3
UnsetForeColor Command4
UnsetForeColor Command5
UnsetForeColor Command6
UnsetForeColor Command7
UnsetForeColor Command8
UnsetForeColor Command9
UnsetForeColor Command10
UnsetForeColor Command11
UnsetForeColor Command12
UnsetForeColor Command14
UnsetForeColor Command15
UnsetForeColor Command16
UnsetForeColor Command35
UnsetForeColor Command48
UnsetForeColor Command49
UnsetForeColor Command50
UnsetForeColor Command52
UnsetForeColor Command53
    'UnsetForeColor Command2
    'UnsetForeColor Command3
    'UnsetForeColor Command4
End Sub

Private Sub HelpTopics_Click()
MsgBox "Some functions use Input Boxes now. I have cleaned up the code slightly and added Sin, Cos and Tan. If you do Sin, Cos or Tan on some small numbers, you may see a minus 2 at the end. This means that you take your answer and move the decimal place 2 places to the left." + vbCrLf + "When you work out fractions, after you have entered all the values, press the + button to add the fractions together, - to take them away, X to times them together or / to divide them. All fraction answers are not given in their simplest form, but still give the correct answer.", vbInformation
End Sub



Private Sub Option1_Click()
Text1.SetFocus

End Sub

Private Sub Option2_Click()
Text1.SetFocus

End Sub

Private Sub Paste_Click()
On Error Resume Next
Text1 = Clipboard.GetText
End Sub

Private Sub Radians_Click()
On Error Resume Next
If Degrees.Checked = False Then
    Degrees.Checked = False
    Radians.Checked = True
End If
End Sub

Private Sub SubXinput_Change()
On Error Resume Next
SubX = SubXinput.Text
End Sub

Private Sub SubYinput_Change()
On Error Resume Next
SubY = SubYinput.Text
End Sub



Private Sub Timer2_Timer()
Timer2.Enabled = False
'Label5.Visible = False
Text1 = ""
Label6.Visible = False
End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False
'Label5.Visible = False
Call Clear_Click
End Sub

Private Sub Timer4_Timer()
Timer4.Enabled = False
'Label5.Visible = False
Text1 = ""
End Sub

Public Function CosRad(Text1)

      
' Initialize variables
   P1 = CDec(1)
    n = 0
  Sum = 0
   dS = CDec(1)
  Lim = CDec(1E-30)
    X = CDec(Trim(Text1))

' Execute cosine computation loop
  Do Until Abs(dS) <= Lim
  dS = P1 * STerm(X, 2 * n)
  Sum = Sum + dS
  n = n + 1
  P1 = -P1
  Loop

  CosRad = Sum
End Function

Private Function SinRad(Text1)
 
      
' Initialize variables
   P1 = CDec(1)
    n = 0
  Sum = 0
   dS = CDec(1)
  Lim = CDec(1E-29)
    X = CDec(Trim(Text1))

' Execute sine computation loop
  Do Until Abs(dS) <= Lim
  dS = P1 * STerm(X, 2 * n + 1)
  Sum = Sum + dS
  n = n + 1
  P1 = -P1
  Loop

  SinRad = Sum
End Function

Private Function TanRad(Text1)
 
Q = CDec(Trim(Text1))

' Compute tangent as Sine/Cosine ratio
  TanRad = SinRad(Q) / CosRad(Q)
End Function

Private Function STerm(x_Val, n_Val)


      X = CDec(x_Val)
      T = CDec(1)
  For i = 1 To Val(n_Val)
      T = T * X / i
  Next i

  STerm = T
End Function

Private Sub redrawgraph()
'redraw the grid in the graph window
Graph.ScaleMode = 1
Graph.Line (Graph.Width / 2, 0)-(Graph.Width / 2, Graph.Height), QBColor(7)
Graph.Line (0, Graph.Height / 2)-(Graph.Width, Graph.Height / 2), QBColor(7)
TicColor = QBColor(7)
For i = 0 To Graph.Width Step 45
'horizontal tics
If i = 450 Or i = 1350 Then
TicColor = vbRed
Else
TicColor = QBColor(7)
End If
Graph.Line (i, (Graph.Height / 2) + 35)-(i, (Graph.Height / 2) - 35), TicColor
Next i
For i = 0 To Graph.Height Step 45
'vertical tics
If i = 450 Or i = 1350 Then
TicColor = vbRed
Else
TicColor = QBColor(7)
End If
Graph.Line (Graph.Width / 2 + 35, i)-(Graph.Width / 2 - 35, i), TicColor
Next i
'print "horizontal +10"
Graph.CurrentX = (Graph.Width / 100) * 20
Graph.CurrentY = (Graph.Height / 100) * 35
Graph.Print "10"
'print "vertical +10"
Graph.CurrentX = (Graph.Width / 100) * 35
Graph.CurrentY = (Graph.Height / 100) * 20
Graph.Print "10"
'print "horizontal -10"
Graph.CurrentX = (Graph.Width / 100) * 70
Graph.CurrentY = (Graph.Height / 100) * 53
Graph.Print "-10"
'print "vertical -10"
Graph.CurrentX = (Graph.Width / 100) * 53
Graph.CurrentY = (Graph.Height / 100) * 70
Graph.Print "-10"
End Sub
