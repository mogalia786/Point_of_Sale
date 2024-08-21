VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmaccedit 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Account you wish to edit , then edit the account ."
   ClientHeight    =   6150
   ClientLeft      =   5460
   ClientTop       =   1980
   ClientWidth     =   11880
   ControlBox      =   0   'False
   Icon            =   "frmaccedit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
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
      Height          =   2415
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   3855
      Begin VB.TextBox txtfirstname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtlastname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User's Firstname"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "User's Lastname"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   1335
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
      Height          =   2370
      Left            =   4080
      TabIndex        =   1
      Top             =   2880
      Width           =   3870
      Begin VB.TextBox txtusername 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtpass 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         TabIndex        =   3
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtconfirm 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         TabIndex        =   2
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   2625
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   4630
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   -2147483626
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin TabDlg.SSTab tt 
      Height          =   2295
      Left            =   8040
      TabIndex        =   13
      Top             =   3000
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4048
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   255
      TabCaption(0)   =   "Admin"
      TabPicture(0)   =   "frmaccedit.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Admin"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Application Rights"
      TabPicture(1)   =   "frmaccedit.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Teller"
      Tab(1).Control(1)=   "Creditors"
      Tab(1).Control(2)=   "Debtors"
      Tab(1).Control(3)=   "Reports"
      Tab(1).Control(4)=   "Suppliers"
      Tab(1).Control(5)=   "Returns"
      Tab(1).Control(6)=   "Stock"
      Tab(1).ControlCount=   7
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
         TabIndex        =   21
         Top             =   480
         Width           =   3255
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
         TabIndex        =   20
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Returns 
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
         TabIndex        =   19
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox Suppliers 
         Caption         =   "Suppliers"
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
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox Reports 
         Caption         =   "Reports"
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
         TabIndex        =   17
         Top             =   1440
         Width           =   1695
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
         TabIndex        =   16
         Top             =   720
         Width           =   2295
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
         TabIndex        =   15
         Top             =   1080
         Width           =   1935
      End
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
         TabIndex        =   14
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label Label8 
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
         TabIndex        =   23
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label Label6 
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
         TabIndex        =   22
         Top             =   480
         Width           =   3615
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   495
      Left            =   9960
      TabIndex        =   24
      Top             =   5520
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Update Account"
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
      MICON           =   "frmaccedit.frx":047A
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   495
      Left            =   8400
      TabIndex        =   25
      Top             =   5520
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
      MICON           =   "frmaccedit.frx":0496
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
Attribute VB_Name = "frmaccedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim rsacc As New Recordset
Dim ch As ColumnHeader
Dim li As ListItem
lv1.View = lvwReport

Set ch = lv1.ColumnHeaders.Add(, "uname", "Username", lv1.Width / 5)
Set ch = lv1.ColumnHeaders.Add(, "pass", "Password", lv1.Width / 5)
Set ch = lv1.ColumnHeaders.Add(, "Firstname", "Firstname", lv1.Width / 5)
Set ch = lv1.ColumnHeaders.Add(, "lastname", "Lastname", lv1.Width / 5)
Admin.Value = 0
Stock.Value = 0
returns.Value = 0
Suppliers.Value = 0
Debtors.Value = 0
Creditors.Value = 0
Reports.Value = 0
Teller.Value = 0



With rsacc
.Open "select * from logon", c, adOpenDynamic, adLockOptimistic
End With
Do While rsacc.EOF = False
Set li = lv1.ListItems.Add(, , rsacc!username)
With li
.SubItems(1) = rsacc!passwords

.SubItems(2) = rsacc!firstname
.SubItems(3) = rsacc!lastname
End With
rsacc.MoveNext
Loop
rsacc.Close

End Sub



Private Sub LaVolpeButton1_Click()
Dim rsacc As New Recordset
Dim edit As String
Dim mlastname As String

Dim ch As ColumnHeader
Dim li As ListItem
edit = lv1.SelectedItem
mlastname = lv1.SelectedItem.SubItems(3)

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
.Open "select * from logon where username='" & edit & "' and lastname='" & mlastname & "'", c, adOpenDynamic, adLockOptimistic
End With
If txtconfirm.Text = txtpass.Text Then
With rsacc

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
'RecordAction CurrentUser, "Edited Account Details...", Date, Time, "Successful..", txtusername.Text

MsgBox "Account Details successfully updated.", vbInformation, "Status"

txtfirstname.Text = ""
txtlastname.Text = ""
txtusername.Text = ""
txtpass.Text = ""
txtconfirm.Text = ""

Admin.Value = 0
Stock.Value = 0
returns.Value = 0
Suppliers.Value = 0
Debtors.Value = 0
Creditors.Value = 0
Reports.Value = 0
Teller.Value = 0


lv1.ListItems.Clear

With rsacc
.Open "logon", c, adOpenDynamic, adLockOptimistic
End With
Do While rsacc.EOF = False
Set li = lv1.ListItems.Add(, , rsacc!username)
With li
.SubItems(1) = rsacc!passwords

.SubItems(2) = rsacc!firstname
.SubItems(3) = rsacc!lastname
End With
rsacc.MoveNext
Loop
rsacc.Close

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

Private Sub lv1_Click()
Dim rsedit As New Recordset
Dim edit As String
Dim mlastname As String
If lv1.ListItems.Count > 0 Then
edit = lv1.SelectedItem
mlastname = lv1.SelectedItem.SubItems(3)

With rsedit
.Open "select * from logon where username='" & edit & "' and lastname='" & mlastname & "'", c, adOpenDynamic, adLockOptimistic
End With
txtfirstname.Text = rsedit!firstname
txtlastname.Text = rsedit!lastname

txtusername.Text = rsedit!username
txtpass.Text = rsedit!passwords
txtconfirm.Text = txtpass.Text
If rsedit!Admin = "Yes" Then
    Admin.Value = 1
Else
    Admin.Value = 0
End If
If rsedit!canstock = "Yes" Then
    Stock.Value = 1
Else
    Stock.Value = 0
End If
If rsedit!canreturn = "Yes" Then
    returns.Value = 1
Else
    returns.Value = 0
End If
If rsedit!cansupplier = "Yes" Then
    Suppliers.Value = 1
Else
    Suppliers.Value = 0
End If
If rsedit!candebtor = "Yes" Then
    Debtors.Value = 1
Else
    Debtors.Value = 0
End If
If rsedit!cancreditor = "Yes" Then
    Creditors.Value = 1
Else
    Creditors.Value = 0
End If
If rsedit!canreport = "Yes" Then
    Reports.Value = 1
Else
    Reports.Value = 0
End If
If rsedit!canteller = "Yes" Then
    Teller.Value = 1
Else
    Teller.Value = 0
End If

If Admin.Value Then
   Stock.Value = 1
    returns.Value = 1
    Suppliers.Value = 1
   Debtors.Value = 1
   Creditors.Value = 1
   Reports.Value = 1
   Teller.Value = 1
   tt.TabEnabled(1) = False
End If

 '   SAles.Value = 0
  '  Card.Value = 0
  '  Exhibitor.Value = 0
  '  Comp.Value = 0
  '  Ticket.Value = 0
  '  Report.Value = 0
  '  Teller.Value = 0
  '  tt.TabEnabled(1) = True
'End If




rsedit.Close


End If
End Sub

