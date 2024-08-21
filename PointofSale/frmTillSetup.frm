VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmTillSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Till Setup"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLOC 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   600
      Width           =   3255
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Update"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmTillSetup.frx":0000
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
   Begin VB.TextBox txtTillId 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cancel / E&xit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmTillSetup.frx":001C
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Location:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Till ID:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmTillSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim r As New Recordset
Dim cs As New Connection
On Error GoTo NC3
With cs
.ConnectionString = App.Path & "/Temp.mdb"
.Provider = "microsoft.jet.oledb.4.0"
.Open
End With
With r
.Open "Till", cs, adOpenDynamic, adLockOptimistic
End With
txtTillId = r!TillId
r.Close
With r
.Open "select * from Till where tillid='" & txtTillId & "'", c, adOpenDynamic, adLockOptimistic
End With
txtLoc = r!location
r.Close
cs.Close
Exit Sub
NC3:
MsgBox "No connection found! Cannot continue with current operation.", vbCritical, "No Connection"

End Sub

Private Sub LaVolpeButton1_Click()
Dim cs As New Connection
Dim r As New Recordset
On Error GoTo NTS

If Len(txtTillId) = 0 Then
    MsgBox "Please enter a valid till id!", vbExclamation, "Status"
    txtTillId.SetFocus
    Exit Sub
End If
If Len(txtLoc) = 0 Then
    MsgBox "Please enter a valid location!", vbExclamation, "Status"
    txtLoc.SetFocus
    Exit Sub
End If

res = MsgBox("Are you sure you wish to change the Till ID?", vbYesNo + vbQuestion, "Change Till ID?")
If res = vbNo Then Exit Sub
With r
.Open "select * from till where tillid='" & UCase(txtTillId) & "' and location='" & UCase(txtLoc) & "'", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = False Then
    MsgBox "Till ID already exist! Please select another!", vbExclamation, "Status"
    txtTillId.SetFocus
    Exit Sub
End If
r.Close

With cs
.ConnectionString = App.Path & "/Temp.mdb"
.Provider = "microsoft.jet.oledb.4.0"
.Open
End With
With r
.Open "select * from till", cs, adOpenDynamic, adLockOptimistic
End With
aa = r!TillId
r!TillId = UCase(txtTillId)
r.Update
r.Close
With r
.Open "select * from sales where tillid='" & aa & "'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    r!TillId = UCase(txtTillId)
    r.Update
    r.MoveNext
Loop
r.Close

With r
.Open "select * from cardtransaction where tillid='" & aa & "'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    r!TillId = UCase(txtTillId)
    r.Update
    r.MoveNext
Loop
r.Close

With r
.Open "select * from till where tillid='" & aa & "'", c, adOpenDynamic, adLockOptimistic
End With
r!TillId = UCase(txtTillId)
r!location = UCase(txtLoc)
r.Update
r.Close
TillId = UCase(txtTillId)



MsgBox "Till ID successfully updated!", vbInformation, "Status"
Unload Me
Exit Sub
NTS:
MsgBox "No connection to server found! cannot complete current operation. Please try again later!", vbCritical, "No Connection"


End Sub

Private Sub LaVolpeButton2_Click()
Unload Me

End Sub
