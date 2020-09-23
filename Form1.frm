VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Message Boxes"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBut 
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtBut 
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtBut 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.ComboBox cboIcon 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   4080
      List            =   "Form1.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.ComboBox cboStyle 
      Height          =   315
      ItemData        =   "Form1.frx":0053
      Left            =   1440
      List            =   "Form1.frx":0069
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txtPrompt 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show MessagBox"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   1980
      Width           =   1815
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   4680
      Picture         =   "Form1.frx":00BE
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   3
      Left            =   4680
      Picture         =   "Form1.frx":0500
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   1
      Left            =   4680
      Picture         =   "Form1.frx":0942
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   2
      Left            =   4680
      Picture         =   "Form1.frx":0D84
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblBut 
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   2100
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblBut 
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   1740
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblBut 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Icon :"
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Message Style :"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Message Prompt :"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Message Title :"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MsgStyle As VbMsgBoxStyle
Dim Msgico As VbMsgBoxStyle

Private Sub cboIcon_Click()
For x = 0 To imgIcon.Count - 1
    imgIcon(x).Visible = False
Next x
Select Case cboIcon.ListIndex
    Case 0
        Msgico = 0
    Case 1
        Msgico = vbCritical
        imgIcon(0).Visible = True
    Case 2
        Msgico = vbExclamation
        imgIcon(1).Visible = True
    Case 3
        Msgico = vbInformation
        imgIcon(2).Visible = True
    Case 4
        Msgico = vbQuestion
        imgIcon(3).Visible = True
End Select

End Sub

Private Sub cboStyle_Click()
lblBut(1).Visible = False
lblBut(2).Visible = False
txtBut(1).Visible = False
txtBut(2).Visible = False
Select Case cboStyle.ListIndex
Case 0
    MsgStyle = vbOKOnly
    lblBut(0).Caption = "OK Button"
Case 1
    MsgStyle = vbAbortRetryIgnore
    lblBut(1).Visible = True
    lblBut(2).Visible = True
    txtBut(1).Visible = True
    txtBut(2).Visible = True
    lblBut(0).Caption = "Abort Button"
    lblBut(1).Caption = "Retry Button"
    lblBut(2).Caption = "Ignore Button"
 Case 2
    MsgStyle = vbOKCancel
    lblBut(1).Visible = True
    txtBut(1).Visible = True
    lblBut(0).Caption = "OK Button"
    lblBut(1).Caption = "Cancel Button"
 Case 3
    MsgStyle = vbRetryCancel
    lblBut(1).Visible = True
    txtBut(1).Visible = True
    lblBut(0).Caption = "Retry Button"
    lblBut(1).Caption = "Cancel Button"
 Case 4
    MsgStyle = vbYesNo
    lblBut(1).Visible = True
    txtBut(1).Visible = True
    lblBut(0).Caption = "Yes Button"
    lblBut(1).Caption = "No Button"
 Case 5
    MsgStyle = vbYesNoCancel
    lblBut(1).Visible = True
    lblBut(2).Visible = True
    txtBut(1).Visible = True
    txtBut(2).Visible = True
    lblBut(0).Caption = "Yes Button"
    lblBut(1).Caption = "No Button"
    lblBut(2).Caption = "Cancel Button"
End Select

End Sub

Private Sub Command1_Click()
Dim mReturn As String
mReturn = BBmsgbox(Me.hwnd, MsgStyle, txtTitle.Text, txtPrompt.Text, Msgico, txtBut(0), txtBut(1), txtBut(2))
MsgBox "You pressed " + mReturn
'Example Usage
'Include ModMsgBox module in your project
'and call like this :
'Dim mReturn As String
'mReturn = BBmsgbox(Me.hwnd, vbOKCancel, "Bobo Enterprises", "Test Prompt", vbCritical, "OK test", "Cancel test")
'Select Case mReturn
'Case "OK test"
'    'put your action here
'Case "Cancel test"
'    'put your action here
'End Select
End Sub

Private Sub Form_Load()
cboIcon.ListIndex = 0
cboStyle.ListIndex = 0
End Sub

