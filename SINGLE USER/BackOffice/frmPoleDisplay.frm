VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmPoleDisplay 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pole Display Setup"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   3990
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1440
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
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmPoleDisplay.frx":0000
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
      BackColor       =   &H80000016&
      Caption         =   "Pole Display Welcome Message"
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
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3735
      Begin VB.TextBox Hl 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   840
         MaxLength       =   20
         TabIndex        =   0
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox Hl 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   840
         MaxLength       =   20
         TabIndex        =   1
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Line1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   -360
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Line 2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   -360
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1440
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
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmPoleDisplay.frx":001C
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
Attribute VB_Name = "frmPoleDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim r As New Recordset
With r
.Open "select * from poledisplay", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = False Then
    Hl(0).Text = r!line1
    Hl(1).Text = r!line2
End If
r.Close

End Sub

Private Sub LaVolpeButton1_Click()
Dim r As New Recordset
res = MsgBox("Are you sure you wish to update any changes made?", vbYesNo + vbQuestion, "Update?")
If res = vbNo Then Unload Me
If Len(Hl(0).Text) = 0 Then
    MsgBox "Please insert message line!"
    Hl(0).SetFocus
    Exit Sub
End If
If Len(Hl(1).Text) = 0 Then
    MsgBox "Please insert message line!"
    Hl(1).SetFocus
    Exit Sub
End If

With r
.Open "select * from poledisplay", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = False Then
    r!line1 = Hl(0).Text
    r!line2 = Hl(1).Text
    r.Update
Else
    r.AddNew
    r!line1 = Hl(0).Text
    r!line2 = Hl(1).Text
    r.Update
End If
r.Close
MsgBox "Update successfull!", vbInformation, "Status"
Unload Me
End Sub

Private Sub LaVolpeButton2_Click()
Unload Me

End Sub
