VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmCompSetup 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Setup"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   2040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Create/Edit Company"
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
      MICON           =   "frmCompSetup.frx":0000
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
   Begin VB.Frame Frame1 
      Caption         =   "Company details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtadd 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txttel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Company Tel:"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Company address:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Company name:"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "E&xit"
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
      MICON           =   "frmCompSetup.frx":001C
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
Attribute VB_Name = "frmCompSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim r As New Recordset
With r
.Open "select * from company", c, adOpenDynamic, adLockOptimistic
End With
txtname = r!cname
txtadd.Text = r!cadd
txttel.Text = r!ctel
r.Close

End Sub

Private Sub LaVolpeButton1_Click()
Dim r As New Recordset
If Len(txtname.Text) = 0 Then
    MsgBox "Please insert nam!", vbExclamation, "Status"
    txtname.SetFocus
    Exit Sub
End If
If Len(txtadd) = 0 Then
    txtadd.Text = "N/A"
End If
If Len(txttel) = 0 Then
    txttel.Text = "N/A"
End If
With r
.Open "select * from company", c, adOpenDynamic, adLockOptimistic
End With
r!cname = UCase(txtname)
r!cadd = UCase(txtadd)
r!ctel = UCase(txttel)
r.Update
r.Close
CompName = UCase(txtname.Text)
MsgBox "Details updated successfully!", vbExclamation, "Updated!"
Unload Me

End Sub

Private Sub LaVolpeButton2_Click()
Unload Me

End Sub
