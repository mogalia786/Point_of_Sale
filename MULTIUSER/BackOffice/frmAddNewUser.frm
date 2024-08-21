VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAddNewUser 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter New Account Details and press "
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   195
   ClientWidth     =   9705
   ControlBox      =   0   'False
   Icon            =   "frmAddNewUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   495
      Left            =   8160
      TabIndex        =   13
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Create User"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   0
      FCOL            =   16777215
      FCOLO           =   33023
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmAddNewUser.frx":0442
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin TabDlg.SSTab tt 
      Height          =   2895
      Left            =   4080
      TabIndex        =   12
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   255
      TabCaption(0)   =   "Admin"
      TabPicture(0)   =   "frmAddNewUser.frx":045E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Admin"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Application Rights"
      TabPicture(1)   =   "frmAddNewUser.frx":047A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(1)=   "Stock"
      Tab(1).Control(2)=   "returns"
      Tab(1).Control(3)=   "Suppliers"
      Tab(1).Control(4)=   "Reports"
      Tab(1).Control(5)=   "Debtors"
      Tab(1).Control(6)=   "Creditors"
      Tab(1).Control(7)=   "Teller"
      Tab(1).ControlCount=   8
      Begin VB.CheckBox Teller 
         Caption         =   "Teller Rights Only"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   -74880
         TabIndex        =   24
         Top             =   1800
         Width           =   3255
      End
      Begin VB.CheckBox Creditors 
         Caption         =   "Creditors"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   -73560
         TabIndex        =   22
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox Debtors 
         Caption         =   "Debtors"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   -73560
         TabIndex        =   21
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox Reports 
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   -73560
         TabIndex        =   20
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox Suppliers 
         Caption         =   "Supplier"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   -74880
         TabIndex        =   19
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox returns 
         Caption         =   "Returns"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   -74880
         TabIndex        =   18
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox Stock 
         Caption         =   "Stock"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   -74880
         TabIndex        =   17
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Admin 
         Caption         =   "Assign Administrator Rights"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "User has the following rights!"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74880
         TabIndex        =   23
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label6 
         Caption         =   "WARNING! Assign this right to the administrator only!!"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   360
         TabIndex        =   16
         Top             =   1200
         Width           =   3375
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000016&
      Caption         =   "Access Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   3855
      Begin VB.TextBox txtconfirm 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtpass 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtusername 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000016&
      Caption         =   "New Account  Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin VB.TextBox txtlastname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtfirstname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "User's Lastname"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User's Firstname"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   495
      Left            =   8160
      TabIndex        =   14
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "E&xit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   0
      FCOL            =   16777215
      FCOLO           =   33023
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmAddNewUser.frx":0496
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
End
Attribute VB_Name = "frmAddNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()




End Sub

Private Sub Image1_Click()
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)


End Sub


Private Sub Image2_Click()

End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub


Private Sub Label8_Click()

End Sub


Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
End Sub


Private Sub Admin_Click()
If Admin.Value Then
    Stock.Value = 1
    Suppliers.Value = 1
    returns.Value = 1
    Debtors.Value = 1
    Creditors.Value = 1
    Reports.Value = 1
    Teller.Value = 1
    tt.TabEnabled(1) = False
Else
    Stock.Value = 0
    Suppliers.Value = 0
    returns.Value = 0
    Debtors.Value = 0
    Creditors.Value = 0
    Reports.Value = 0
    Teller.Value = 0
    tt.TabEnabled(1) = True
End If

End Sub

Private Sub Form_Load()
If Admin.Value Then
    Stock.Value = 1
    Suppliers.Value = 1
    returns.Value = 1
    Debtors.Value = 1
    Creditors.Value = 1
    Reports.Value = 1
    Teller.Value = 1
    tt.TabEnabled(1) = False
Else
    Stock.Value = 0
    Suppliers.Value = 0
    returns.Value = 0
    Debtors.Value = 0
    Creditors.Value = 0
    Reports.Value = 0
    Teller.Value = 0
    tt.TabEnabled(1) = True
End If

End Sub

Private Sub LaVolpeButton1_Click()
Dim rsacc As New Recordset
If Len(txtfirstname.Text) = 0 Then
MsgBox "Please insert Firstname!", vbCritical, "Error!"
txtfirstname.SetFocus
Exit Sub
End If
If Len(txtlastname.Text) = 0 Then
MsgBox "Please insert Lastname!", vbCritical, "Error!"
txtlastname.SetFocus
Exit Sub
End If
If Len(txtusername.Text) = 0 Then
MsgBox "Please insert Username!", vbCritical, "Error!"
txtusername.SetFocus
Exit Sub
End If
If Len(txtpass.Text) = 0 Then
MsgBox "Please insert Password!", vbCritical, "Error!"
txtpass.SetFocus
Exit Sub
End If
If Len(txtconfirm.Text) = 0 Then
MsgBox "Please Confirm Password!", vbCritical, "Error!"
txtconfirm.SetFocus
Exit Sub
End If
With rsacc
.Open "select * from logon where username='" & txtusername.Text & "'", c, adOpenDynamic, adLockOptimistic
End With
If rsacc.EOF = False Then
    MsgBox "Username " & txtusername & " already exist! Please select another username eg. " & txtusername & "123", vbExclamation, "username Exist!"
    txtusername.Text = ""
    txtusername.SetFocus
    rsacc.Close
    Exit Sub
End If
rsacc.Close

With rsacc
.Open "logon", c, adOpenDynamic, adLockOptimistic
End With
If txtconfirm.Text = txtpass.Text Then
With rsacc
.AddNew
!username = txtusername.Text
!passwords = txtpass.Text
!firstname = txtfirstname.Text
!lastname = txtlastname.Text
If Admin.Value Then
    !Admin = "Yes"
Else
    !Admin = "No"
End If
'/////////////////////////
If Stock.Value Then
    !canstock = "Yes"
Else
    !canstock = "No"
End If
'///////////////////////////
If returns.Value Then
    !canreturn = "Yes"
Else
    !canreturn = "No"
End If
'//////////////////////////
If Suppliers.Value Then
    !cansupplier = "Yes"
Else
    !cansupplier = "No"
End If
'//////////////////////////
If Debtors.Value Then
    !candebtor = "Yes"
Else
    !candebtor = "No"
End If
'//////////////////////////
If Creditors.Value Then
    !cancreditor = "Yes"
Else
    !cancreditor = "No"
End If
'//////////////////////////
If Reports.Value Then
    !canreport = "Yes"
Else
    !canreport = "No"
End If
'//////////////////////////
If Teller.Value Then
    !canteller = "Yes"
Else
    !canteller = "No"
End If

.Update
End With
rsacc.Close

res = MsgBox("New Account Details successfully updated. Do you wish to add another Account?", vbQuestion + vbYesNo, "Status")
If res = vbYes Then
Admin.Value = 1
Stock.Value = 1
returns.Value = 1
Suppliers.Value = 1
Debtors.Value = 1
Creditors.Value = 1
Reports.Value = 1
Teller.Value = 1

    txtfirstname.Text = ""
    txtlastname.Text = ""
    txtusername.Text = ""
    txtpass.Text = ""
    txtconfirm.Text = ""
    txtfirstname.SetFocus
Else
    Unload Me
End If

Else
MsgBox "Invalid Password! Try Again!", vbCritical, "Password Error!"
txtpass.Text = ""
txtconfirm.Text = ""
txtpass.SetFocus
rsacc.Close

End If

End Sub

Private Sub LaVolpeButton2_Click()
Unload Me

End Sub
