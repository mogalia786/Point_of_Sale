VERSION 5.00
Begin VB.Form frmlogon 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2760
   ClientLeft      =   4740
   ClientTop       =   4350
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   Picture         =   "frmlogon.frx":0000
   ScaleHeight     =   2760
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4080
      Top             =   300
   End
   Begin VB.TextBox txtpassword 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1680
      Width           =   2730
   End
   Begin VB.TextBox txtusername 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   1080
      Width           =   2730
   End
   Begin VB.Image Image3 
      Height          =   855
      Left            =   1200
      Picture         =   "frmlogon.frx":2031
      Top             =   240
      Width           =   2130
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   4320
      TabIndex        =   7
      Top             =   45
      Width           =   165
   End
   Begin VB.Image Image2 
      Height          =   195
      Left            =   4275
      Picture         =   "frmlogon.frx":2695
      Top             =   45
      Width           =   270
   End
   Begin VB.Label lbllogon 
      BackStyle       =   0  'Transparent
      Caption         =   "Logon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   570
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3150
      TabIndex        =   5
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   495
      TabIndex        =   4
      Top             =   2145
      Width           =   855
   End
   Begin VB.Image cmdcancel 
      Height          =   345
      Left            =   2880
      Picture         =   "frmlogon.frx":2A82
      Top             =   2130
      Width           =   1425
   End
   Begin VB.Image cmdok 
      Height          =   345
      Left            =   225
      Picture         =   "frmlogon.frx":3261
      Top             =   2130
      Width           =   1425
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   1365
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   345
      TabIndex        =   2
      Top             =   1080
      Width           =   1365
   End
   Begin VB.Image Image1 
      Height          =   2760
      Left            =   0
      Picture         =   "frmlogon.frx":3A40
      Top             =   0
      Width           =   4590
   End
End
Attribute VB_Name = "frmlogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancel_Click()
End

End Sub

Private Sub cmdcancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdcancel.Picture = LoadPicture(App.Path & "\buttonup.jpg")
cmdcancel.ToolTipText = "Press to cancel"


End Sub

Private Sub cmdOK_Click()
Dim rslogon As New Recordset
Dim username As String
Dim pass As String
Dim lastname As String
Dim datedif As String
On Error GoTo NL

If Len(txtusername.Text) > 0 And Len(txtpassword.Text) > 0 Then
username = txtusername.Text
pass = txtpassword.Text

With rslogon
.Open "select * from logon where username='" & username & "' and passwords='" & pass & "'", c, adOpenDynamic, adLockOptimistic
End With
If rslogon.EOF = True Then
MsgBox "Access Denied! Please try again.", vbCritical, "Logon Error!"
rslogon.Close

txtusername.Text = ""
txtpassword.Text = ""
txtusername.SetFocus
Exit Sub
Else
rslogon.Close

Unload Me
CurrentUser = username
With rslogon
.Open "select * from LogSheet", c, adOpenDynamic, adLockOptimistic
End With
rslogon.AddNew
rslogon!username = username
rslogon!dateloggedon = Date
rslogon!timeloggedon = Time
rslogon.Update
rslogon.Close
RecordAction CurrentUser, Date, Time, "Successful..", "User Logged On..."
ShiftStartTime = Time
frmMain.Show vbModal
End If
Else
MsgBox "Insufficient Details!Please try again!", vbCritical, "Error!"
End If
Exit Sub
NL:
MsgBox "Connection to server has been lost! Please try logging on at a later stage.", vbCritical, "No Connection!"

End Sub

Private Sub cmdok_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdok.Picture = LoadPicture(App.Path & "\buttonup.jpg")
cmdok.ToolTipText = "Press to confirm details"

End Sub

Private Sub Form_Load()
Call MakeTranslucent(Me, vbBlue)

End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
Image2.Picture = LoadPicture(App.Path & "/closeup.jpg")
End If

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdok.Picture = LoadPicture(App.Path & "/buttondn.jpg")
cmdcancel.Picture = LoadPicture(App.Path & "/buttondn.jpg")

End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
Image2.Picture = LoadPicture(App.Path & "/closedn.jpg")
Unload Me
End If

End Sub

Public Sub Label5_Click()
Dim rslogon As New Recordset
Dim username As String
Dim pass As String
Dim lastname As String
Dim datedif As String
On Error GoTo NL2

If Len(txtusername.Text) > 0 And Len(txtpassword.Text) > 0 Then
username = txtusername.Text
pass = txtpassword.Text

With rslogon
.Open "select * from logon where username='" & username & "' and passwords='" & pass & "'", c, adOpenDynamic, adLockOptimistic
End With
If rslogon.EOF = True Then
MsgBox "Access Denied! Please try again.", vbCritical, "Logon Error!"
rslogon.Close

txtusername.Text = ""
txtpassword.Text = ""
txtusername.SetFocus
Exit Sub
Else
rslogon.Close

Unload Me
CurrentUser = username
With rslogon
.Open "select * from LogSheet", c, adOpenDynamic, adLockOptimistic
End With
rslogon.AddNew
rslogon!username = username
rslogon!dateloggedon = Date
rslogon!timeloggedon = Time
rslogon.Update
rslogon.Close
RecordAction CurrentUser, Date, Time, "Successful..", "User Logged On..."
ShiftStartTime = Time
frmMain.Show vbModal
End If
Else
MsgBox "Insufficient Details!Please try again!", vbCritical, "Error!"
End If
Exit Sub
NL2:
MsgBox "Connection to server has been lost! Please try logging on at a later stage.", vbCritical, "No Connection!"
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdok.Picture = LoadPicture(App.Path & "\buttonup.jpg")
cmdok.ToolTipText = "Press to confirm details"

End Sub

Private Sub Label6_Click()
End

End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdcancel.Picture = LoadPicture(App.Path & "\buttonup.jpg")
cmdcancel.ToolTipText = "Press to cancel"



End Sub

Private Sub Label7_Click()
End

End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
Image2.Picture = LoadPicture(App.Path & "/closeup.jpg")
End If

End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
Image2.Picture = LoadPicture(App.Path & "/closedn.jpg")
Unload Me
End If

End Sub

Private Sub optadmin_Click()

End Sub

Private Sub optuser_Click()
End Sub

Private Sub Timer1_Timer()
If lbllogon.ForeColor = vbRed Then
lbllogon.ForeColor = vbWhite
Else
lbllogon.ForeColor = vbRed
End If

End Sub

Private Sub txtpassword_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyRight Then
Me.Label5_Click
End If

End Sub

Private Sub txtusername_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyRight Then
txtpassword.SetFocus
End If

End Sub
