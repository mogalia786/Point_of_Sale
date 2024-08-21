VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmSupervisor 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supervisor Password Required"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "Enter Supervisor details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   4095
      Begin LVbuttons.LaVolpeButton LaVolpeButton1 
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&OK"
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
         BCOL            =   12632256
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmSupervisor.frx":0000
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
      Begin VB.TextBox txtPWD 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtUN 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   480
         Width           =   2175
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton2 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
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
         BCOL            =   12632256
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmSupervisor.frx":001C
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
         Caption         =   "Supervisor Password:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supervisor Username:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmSupervisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LaVolpeButton1_Click()
Dim r As New Recordset
If Len(txtUN) = 0 Then
    MsgBox "Please enter username!", vbExclamation, "Status"
    txtUN.SetFocus
    Exit Sub
End If
If Len(txtPWD) = 0 Then
    MsgBox "Please enter Password!", vbExclamation, "Status"
    txtPWD.SetFocus
    Exit Sub
End If
With r
.Open "select * from supervisor where susername='" & txtUN & "' and spassword='" & txtPWD & "'", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = False Then
    MsgBox "Supervisor already exist!", vbExclamation, "status"
    txtUN.Text = ""
    txtPWD.Text = ""
    txtUN.SetFocus
    Exit Sub
End If
r.AddNew
r!sUsername = UCase(txtUN)
r!sPassword = UCase(txtPWD)
r.Update
r.Close
MsgBox "Supervisor Account successfully created!", vbInformation, "Status"
Unload Me


End Sub

Private Sub LaVolpeButton2_Click()
Unload Me

End Sub
