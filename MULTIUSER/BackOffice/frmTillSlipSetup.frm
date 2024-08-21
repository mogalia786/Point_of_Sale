VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmTillSlipSetup 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Till Slip Setup"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   5280
      TabIndex        =   32
      Top             =   6840
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "frmTillSlipSetup.frx":0000
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
   Begin VB.Frame Frame2 
      BackColor       =   &H80000016&
      Caption         =   "Footer"
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
      Height          =   2295
      Left            =   120
      TabIndex        =   21
      Top             =   4440
      Width           =   6495
      Begin VB.TextBox Fl 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   840
         MaxLength       =   44
         TabIndex        =   31
         Top             =   1800
         Width           =   5535
      End
      Begin VB.TextBox Fl 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   840
         MaxLength       =   44
         TabIndex        =   29
         Top             =   1440
         Width           =   5535
      End
      Begin VB.TextBox Fl 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   840
         MaxLength       =   44
         TabIndex        =   27
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox Fl 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   840
         MaxLength       =   44
         TabIndex        =   25
         Top             =   720
         Width           =   5535
      End
      Begin VB.TextBox Fl 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   840
         MaxLength       =   44
         TabIndex        =   23
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Line 5:"
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
         Index           =   14
         Left            =   -360
         TabIndex        =   30
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Line 4:"
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
         Index           =   13
         Left            =   -360
         TabIndex        =   28
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Line 3:"
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
         Index           =   12
         Left            =   -360
         TabIndex        =   26
         Top             =   1080
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
         Index           =   11
         Left            =   -360
         TabIndex        =   24
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Line 1:"
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
         Index           =   10
         Left            =   -360
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "Header"
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
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.TextBox Hl 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   9
         Left            =   840
         MaxLength       =   34
         TabIndex        =   19
         Top             =   3600
         Width           =   5535
      End
      Begin VB.TextBox Hl 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   840
         MaxLength       =   34
         TabIndex        =   17
         Top             =   3240
         Width           =   5535
      End
      Begin VB.TextBox Hl 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   840
         MaxLength       =   34
         TabIndex        =   15
         Top             =   2880
         Width           =   5535
      End
      Begin VB.TextBox Hl 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   840
         MaxLength       =   34
         TabIndex        =   13
         Top             =   2520
         Width           =   5535
      End
      Begin VB.TextBox Hl 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   840
         MaxLength       =   34
         TabIndex        =   11
         Top             =   2160
         Width           =   5535
      End
      Begin VB.TextBox Hl 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   840
         MaxLength       =   34
         TabIndex        =   9
         Top             =   1800
         Width           =   5535
      End
      Begin VB.TextBox Hl 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   840
         MaxLength       =   34
         TabIndex        =   7
         Top             =   1440
         Width           =   5535
      End
      Begin VB.TextBox Hl 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   840
         MaxLength       =   34
         TabIndex        =   5
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox Hl 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   840
         MaxLength       =   34
         TabIndex        =   3
         Top             =   720
         Width           =   5535
      End
      Begin VB.TextBox Hl 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   840
         MaxLength       =   34
         TabIndex        =   1
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Line 10:"
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
         Index           =   9
         Left            =   -360
         TabIndex        =   20
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Line 9:"
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
         Index           =   8
         Left            =   -360
         TabIndex        =   18
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Line 8:"
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
         Index           =   7
         Left            =   -360
         TabIndex        =   16
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Line 7:"
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
         Index           =   6
         Left            =   -360
         TabIndex        =   14
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Line 6:"
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
         Index           =   5
         Left            =   -360
         TabIndex        =   12
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Line 5:"
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
         Index           =   4
         Left            =   -360
         TabIndex        =   10
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Line 4:"
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
         Index           =   3
         Left            =   -360
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Line 3:"
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
         Index           =   2
         Left            =   -360
         TabIndex        =   6
         Top             =   1080
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
         TabIndex        =   4
         Top             =   720
         Width           =   1095
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
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   3840
      TabIndex        =   33
      Top             =   6840
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "frmTillSlipSetup.frx":001C
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
Attribute VB_Name = "frmTillSlipSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim r As New Recordset
With r
.Open "select * from header", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = True Then Exit Sub
    Hl(0) = r!line1
    Hl(1) = r!line2
    Hl(2) = r!line3
    Hl(3) = r!line4
    Hl(4) = r!line5
    Hl(5) = r!line6
    Hl(6) = r!line7
    Hl(7) = r!line8
    Hl(8) = r!line9
    Hl(9) = r!line10
r.Close
With r
.Open "select * from footer", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = True Then Exit Sub
    Fl(0) = r!line1
    Fl(1) = r!line2
    Fl(2) = r!line3
    Fl(3) = r!line4
    Fl(4) = r!line5
   
r.Close

For i = 0 To 9
    If Hl(i).Text = "BLANK" Then
        Hl(i).Text = ""
    End If
Next i
For i = 0 To 4
    If Fl(i).Text = "BLANK" Then
        Fl(i).Text = ""
    End If
Next i
End Sub

Private Sub LaVolpeButton1_Click()
Dim r As New Recordset
res = MsgBox("Are you sure you wish to update any changes made?", vbYesNo + vbQuestion, "Update?")
If res = vbNo Then Unload Me
With r
.Open "select * from header", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = True Then
    r.AddNew
    For i = 0 To 9
        If Hl(i).Text = "" Then
            Hl(i).Text = "BLANK"
        End If
        r.Fields("line" & i + 1) = Hl(i)
    Next i
    r.Update
Else
    For i = 0 To 9
        If Hl(i).Text = "" Then
            Hl(i).Text = "BLANK"
        End If
        r.Fields("line" & i + 1) = Hl(i)
    Next i
    r.Update
End If
r.Close

With r
.Open "select * from footer", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = True Then
    r.AddNew
    For i = 0 To 4
        If Fl(i).Text = "" Then
            Fl(i).Text = "BLANK"
        End If
        r.Fields("line" & i + 1) = Fl(i)
    Next i
    r.Update
Else
    For i = 0 To 4
        If Fl(i).Text = "" Then
            Fl(i).Text = "BLANK"
        End If
        r.Fields("line" & i + 1) = Fl(i)
    Next i
    r.Update
End If
r.Close
MsgBox "Update successfull!", vbInformation, "Status"
Unload Me
End Sub

Private Sub LaVolpeButton2_Click()
Unload Me

End Sub
