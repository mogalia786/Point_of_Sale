VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmCostCode 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cost Code"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   330
      Left            =   6540
      TabIndex        =   21
      Top             =   2430
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   582
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
      COLTYPE         =   2
      BCOL            =   0
      FCOL            =   16777215
      FCOLO           =   33023
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmCostCode.frx":0000
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
   Begin VB.TextBox T0 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6900
      MaxLength       =   1
      TabIndex        =   20
      Text            =   "E"
      Top             =   1350
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "Cost Code"
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
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   7575
      Begin VB.TextBox T9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   6060
         MaxLength       =   1
         TabIndex        =   19
         Text            =   "S"
         Top             =   1140
         Width           =   495
      End
      Begin VB.TextBox T8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   5340
         MaxLength       =   1
         TabIndex        =   18
         Text            =   "R"
         Top             =   1140
         Width           =   495
      End
      Begin VB.TextBox T7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   4620
         MaxLength       =   1
         TabIndex        =   17
         Text            =   "O"
         Top             =   1140
         Width           =   495
      End
      Begin VB.TextBox T6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   3900
         MaxLength       =   1
         TabIndex        =   16
         Text            =   "H"
         Top             =   1140
         Width           =   495
      End
      Begin VB.TextBox T5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   3180
         MaxLength       =   1
         TabIndex        =   15
         Text            =   "K"
         Top             =   1140
         Width           =   495
      End
      Begin VB.TextBox T4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   2460
         MaxLength       =   1
         TabIndex        =   14
         Text            =   "C"
         Top             =   1140
         Width           =   495
      End
      Begin VB.TextBox T3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1740
         MaxLength       =   1
         TabIndex        =   13
         Text            =   "A"
         Top             =   1140
         Width           =   495
      End
      Begin VB.TextBox T2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1020
         MaxLength       =   1
         TabIndex        =   12
         Text            =   "L"
         Top             =   1140
         Width           =   495
      End
      Begin VB.TextBox T1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   300
         MaxLength       =   1
         TabIndex        =   11
         Text            =   "B"
         Top             =   1140
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   9
         Left            =   6720
         TabIndex        =   10
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   8
         Left            =   6000
         TabIndex        =   9
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   7
         Left            =   5280
         TabIndex        =   8
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   6
         Left            =   4560
         TabIndex        =   7
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   5
         Left            =   3840
         TabIndex        =   6
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   4
         Left            =   3120
         TabIndex        =   5
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   3
         Left            =   2400
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   2
         Left            =   1680
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   1
         Left            =   960
         TabIndex        =   2
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   615
         Index           =   19
         Left            =   6720
         Top             =   1080
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   615
         Index           =   18
         Left            =   6720
         Top             =   480
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   615
         Index           =   17
         Left            =   6000
         Top             =   1080
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   615
         Index           =   16
         Left            =   6000
         Top             =   480
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   615
         Index           =   15
         Left            =   5280
         Top             =   1080
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   615
         Index           =   14
         Left            =   5280
         Top             =   480
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   615
         Index           =   13
         Left            =   4560
         Top             =   1080
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   615
         Index           =   12
         Left            =   4560
         Top             =   480
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   615
         Index           =   11
         Left            =   3840
         Top             =   1080
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   615
         Index           =   10
         Left            =   3840
         Top             =   480
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   615
         Index           =   9
         Left            =   3120
         Top             =   1080
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   615
         Index           =   8
         Left            =   3120
         Top             =   480
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   615
         Index           =   7
         Left            =   2400
         Top             =   1080
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   615
         Index           =   6
         Left            =   2400
         Top             =   480
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   615
         Index           =   5
         Left            =   1680
         Top             =   1080
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   615
         Index           =   4
         Left            =   1680
         Top             =   480
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   615
         Index           =   3
         Left            =   960
         Top             =   1080
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   615
         Index           =   2
         Left            =   960
         Top             =   480
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   615
         Index           =   1
         Left            =   240
         Top             =   1080
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   615
         Index           =   0
         Left            =   240
         Top             =   480
         Width           =   615
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   330
      Left            =   5325
      TabIndex        =   22
      Top             =   2430
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   582
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
      COLTYPE         =   2
      BCOL            =   0
      FCOL            =   16777215
      FCOLO           =   33023
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmCostCode.frx":001C
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
Attribute VB_Name = "frmCostCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim r As New Recordset
With r
.Open "select * from costcode", c, adOpenDynamic, adLockOptimistic
End With
T1.Text = r!one
T2.Text = r!two
T3.Text = r!three
T4.Text = r!four
T5.Text = r!five
T6.Text = r!six
T7.Text = r!seven
T8.Text = r!eight
T9.Text = r!nine
T0.Text = r!zero
r.Close




End Sub

Private Sub LaVolpeButton1_Click()
Dim r As New Recordset
Dim tCode(0 To 9) As String
Dim FoundCode As Boolean
FoundCode = False
If Len(T1.Text) = 0 Then
    MsgBox "Field cannot be zero length!", vbExclamation, "Status"
    T1.SetFocus
    Exit Sub
End If
If Len(T2.Text) = 0 Then
    MsgBox "Field cannot be zero length!", vbExclamation, "Status"
    T2.SetFocus
    Exit Sub
End If
If Len(T3.Text) = 0 Then
    MsgBox "Field cannot be zero length!", vbExclamation, "Status"
    T3.SetFocus
    Exit Sub
End If
If Len(T4.Text) = 0 Then
    MsgBox "Field cannot be zero length!", vbExclamation, "Status"
    T4.SetFocus
    Exit Sub
End If
If Len(T5.Text) = 0 Then
    MsgBox "Field cannot be zero length!", vbExclamation, "Status"
    T5.SetFocus
    Exit Sub
End If
If Len(T6.Text) = 0 Then
    MsgBox "Field cannot be zero length!", vbExclamation, "Status"
    T6.SetFocus
    Exit Sub
End If
If Len(T7.Text) = 0 Then
    MsgBox "Field cannot be zero length!", vbExclamation, "Status"
    T7.SetFocus
    Exit Sub
End If
If Len(T8.Text) = 0 Then
    MsgBox "Field cannot be zero length!", vbExclamation, "Status"
    T8.SetFocus
    Exit Sub
End If
If Len(T9.Text) = 0 Then
    MsgBox "Field cannot be zero length!", vbExclamation, "Status"
    T9.SetFocus
    Exit Sub
End If
If Len(T0.Text) = 0 Then
    MsgBox "Field cannot be zero length!", vbExclamation, "Status"
    T0.SetFocus
    Exit Sub
End If

tCode(0) = T0
tCode(1) = T1
tCode(2) = T2
tCode(3) = T3
tCode(4) = T4
tCode(5) = T5
tCode(6) = T6
tCode(7) = T7
tCode(8) = T8
tCode(9) = T9

For i = 0 To 9
    For zx = i + 1 To 9
        If UCase(tCode(i)) = UCase(tCode(zx)) Then
            FoundCode = True
            MsgBox "No two codes can be the same!", vbExclamation, "Status"
            
            If i = 0 Then
                T0 = ""
                T0.SetFocus
            End If
            If i = 1 Then
                T1 = ""
                T1.SetFocus
            End If
            If i = 2 Then
                T2 = ""
                T2.SetFocus
            End If
            If i = 3 Then
                T3 = ""
                T3.SetFocus
            End If
            If i = 4 Then
                T4 = ""
                T4.SetFocus
            End If
            If i = 5 Then
                T5 = ""
                T5.SetFocus
            End If
            If i = 6 Then
                T6 = ""
                T6.SetFocus
            End If
            If i = 7 Then
                T7 = ""
                T7.SetFocus
            End If
            If i = 8 Then
                T8 = ""
                T8.SetFocus
            End If
            If i = 9 Then
                T9 = ""
                T9.SetFocus
            End If

            
    Exit Sub

        Else
            FoundCode = False
        End If
    Next zx
Next i

    
    


res = MsgBox("Are you sure you wish to update changes!", vbYesNo + vbQuestion, "Update?")

If res = vbNo Then
    Exit Sub
End If
With r
.Open "select * from costcode", c, adOpenDynamic, adLockOptimistic
End With
r!one = UCase(T1)
r!two = UCase(T2)
r!three = UCase(T3)
r!four = UCase(T4)
r!five = UCase(T5)
r!six = UCase(T6)
r!seven = UCase(T7)
r!eight = UCase(T8)
r!nine = UCase(T9)
r!zero = UCase(T0)
r.Update
MsgBox "Update successfull!", vbInformation, "Status"
Unload Me



End Sub

Private Sub LaVolpeButton2_Click()
Unload Me

End Sub

Private Sub T0_GotFocus()
T0.SelStart = 0
T0.SelLength = Len(T0)

End Sub

Private Sub T1_GotFocus()
T1.SelStart = 0
T1.SelLength = Len(T1)

End Sub

Private Sub T2_GotFocus()
T2.SelStart = 0
T2.SelLength = Len(T2)

End Sub

Private Sub T3_GotFocus()
T3.SelStart = 0
T3.SelLength = Len(T3)

End Sub

Private Sub T4_GotFocus()
T4.SelStart = 0
T4.SelLength = Len(T4)

End Sub

Private Sub T5_GotFocus()
T5.SelStart = 0
T5.SelLength = Len(T5)

End Sub

Private Sub T6_GotFocus()
T6.SelStart = 0
T6.SelLength = Len(T6)

End Sub

Private Sub T7_GotFocus()
T7.SelStart = 0
T7.SelLength = Len(T7)

End Sub

Private Sub T8_GotFocus()
T8.SelStart = 0
T8.SelLength = Len(T8)

End Sub

Private Sub T9_GotFocus()
T9.SelStart = 0
T9.SelLength = Len(T9)

End Sub
