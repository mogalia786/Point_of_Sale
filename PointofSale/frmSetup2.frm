VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmSetup2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Till Point Setup"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   Icon            =   "frmSetup2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLoc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   720
      Width           =   2295
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Ok"
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
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   33023
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmSetup2.frx":0442
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
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
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
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   33023
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmSetup2.frx":045E
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
      Alignment       =   1  'Right Justify
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
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Till Point ID:"
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
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmSetup2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LaVolpeButton1_Click()
Dim cs As New Connection
Dim r As New Recordset
On Error GoTo NC
If Len(txtTillId) = 0 Then Exit Sub
If Len(txtLOC) = 0 Then Exit Sub
With r
.Open "select * from till where tillid='" & UCase(txtTillId) & "'", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = False Then
    MsgBox "Till ID already exist! Please select another!", vbExclamation, "Status"
    txtTillId.SetFocus
    c.Close
    Exit Sub
End If
r.Close
With r
.Open "select * from till", c, adOpenDynamic, adLockOptimistic
End With
r.AddNew
r!TillId = UCase(txtTillId.Text)
r!location = UCase(txtLOC)
r.Update
r.Close

With cs
.ConnectionString = App.Path & "/Temp.mdb"
.Provider = "microsoft.jet.oledb.4.0"
.Open
End With
With r
.Open "setup", cs, adOpenDynamic, adLockOptimistic
End With
r!initialized = "Yes"
r.Update
r.Close
With r
.Open "Till", cs, adOpenDynamic, adLockOptimistic
End With
r!TillId = UCase(txtTillId.Text)
r.Update
r.Close

Unload Me
frmlogon.Show
Exit Sub
NC:
MsgBox "Connection to server has been lost! Please retry later!", vbCritical, "No Connection"

End Sub

Private Sub LaVolpeButton2_Click()
End
End Sub
