VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E7DCD8&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register your app into MSN actions menu"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All files (*.*) | *.*"
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E7DCD8&
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Text            =   "Keyname"
      Top             =   3960
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E7DCD8&
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Text            =   "Keyname (recommended to use your menu name)"
      Top             =   1920
      Width           =   4095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C8B5AE&
      Caption         =   "Unregister"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C8B5AE&
      Caption         =   "Register"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E7DCD8&
      Height          =   285
      Left            =   600
      TabIndex        =   4
      Text            =   "URL"
      Top             =   2880
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E7DCD8&
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Text            =   "Name to appear in the menu"
      Top             =   2400
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C8B5AE&
      Caption         =   "..."
      Height          =   255
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C8B5AE&
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Text            =   "Path"
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E7DCD8&
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   5055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C8B5AE&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ferret: <M_Larritt@hotmail.com>"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4680
      Width           =   5055
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   5160
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   5160
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   5160
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5160
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0000
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents msnsession As MsgrSessionManager
Attribute msnsession.VB_VarHelpID = -1
Private Sub Command1_Click()
On Error Resume Next
CommonDialog1.ShowOpen
Text1.Text = CommonDialog1.FileName
End Sub
Private Sub Command2_Click()
On Error GoTo needall
msnsession.RegisterApplication Text4.Text, Text2.Text, Text3.Text, Text1.Text, 0
Exit Sub
needall:
MsgBox "All feilds must be complete", vbCritical, "Error"
End Sub
Private Sub Command3_Click()
On Error GoTo notexist
msnsession.UnRegisterApplication Text5.Text
Exit Sub
notexist:
MsgBox "Key doesn't exist", vbCritical, "Error"
End Sub
Private Sub Form_Load()
Set msnsession = New MsgrSessionManager
MsgBox "This Messenger session manager example was created by Ferret, use it freely, it's pretty simple...I'm not sure why you cant set msgrsession, maybe someone would tell me @ M_Larritt@hotmail.com...Have fun!", vbOKOnly, "Welcome!"
End Sub
