VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmEditSupplier 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Supplier Details / Opening Balance"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H80000016&
      Caption         =   "Supplier Details"
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
      Height          =   2415
      Left            =   120
      TabIndex        =   14
      Top             =   4680
      Width           =   8415
      Begin VB.ComboBox cTerms 
         Height          =   315
         ItemData        =   "frmEditSupplier.frx":0000
         Left            =   5640
         List            =   "frmEditSupplier.frx":004F
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox tOB 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         TabIndex        =   9
         Top             =   1600
         Width           =   2655
      End
      Begin VB.TextBox tFAX 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         TabIndex        =   7
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox tTEL 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5640
         TabIndex        =   6
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox tADD4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox tADD3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox tADD2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox tADD1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox SUP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Balance:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4200
         TabIndex        =   24
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Terms:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4320
         TabIndex        =   23
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4320
         TabIndex        =   22
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4320
         TabIndex        =   21
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address4:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address3:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address2:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address1:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   6840
      TabIndex        =   11
      Top             =   7560
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
      COLTYPE         =   2
      BCOL            =   0
      FCOL            =   16777215
      FCOLO           =   33023
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmEditSupplier.frx":00CA
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
      Caption         =   "Click Supplier "
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
      Height          =   3255
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   9375
      Begin MSComctlLib.ListView lvw1 
         Height          =   2775
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Supplier"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Address1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Address2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Address3"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Address4"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Tel"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Fax"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Terms"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Opening Balance"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
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
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   4215
      Begin VB.TextBox tSUP 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   3975
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   8160
      TabIndex        =   10
      Top             =   7560
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
      COLTYPE         =   2
      BCOL            =   0
      FCOL            =   16777215
      FCOLO           =   33023
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmEditSupplier.frx":00E6
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
Attribute VB_Name = "frmEditSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LaVolpeButton1_Click()
Unload Me

End Sub

Private Sub LaVolpeButton2_Click()
Dim r As New Recordset
res = MsgBox("Are you sure you wish to change/update supplier details?", vbYesNo + vbQuestion, "Update?")
If res = vbNo Then Exit Sub

If Len(SUP) = 0 Then
    MsgBox "Please select supplier!", vbExclamation, "Status"
    lvw1.SetFocus
    Exit Sub
End If
If Len(tADD1) = 0 Then
    tADD1 = "N/A"
End If
If Len(tADD2) = 0 Then
    tADD2 = "N/A"
End If
If Len(tADD3) = 0 Then
    tADD3 = "N/A"
End If
If Len(tADD4) = 0 Then
    tADD4 = "N/A"
End If
If Len(tTEL) = 0 Then
    tTEL = "N/A"
End If
If Len(tFAX) = 0 Then
    tFAX = "N/A"
End If

If Len(cTerms.Text) = 0 Then
    MsgBox "Please select terms!", vbExclamation, "Status"
    cTerms.SetFocus
    Exit Sub
End If
If Len(tOB) = 0 Then
    tOB = "0"
End If

With r
.Open "select * from supplier where supplier='" & SUP & "'", c, adOpenDynamic, adLockOptimistic
End With
r!add1 = UCase(tADD1)
r!add2 = UCase(tADD2)
r!add3 = UCase(tADD3)
r!add4 = UCase(tADD4)
r!tel = UCase(tTEL)
r!fax = UCase(tFAX)
r!terms = cTerms.Text
r!OpeningBalance = Format(tOB, "#####.00")
r.Update
r.Close
MsgBox "Update successfull!", vbInformation, "Status"
Unload Me


End Sub

Private Sub lvw1_Click()
If lvw1.ListItems.Count = 0 Then Exit Sub
If Len(lvw1.SelectedItem) = 0 Then Exit Sub

SUP = lvw1.SelectedItem
tADD1 = lvw1.SelectedItem.SubItems(1)
tADD2 = lvw1.SelectedItem.SubItems(2)
tADD3 = lvw1.SelectedItem.SubItems(3)
tADD4 = lvw1.SelectedItem.SubItems(4)
tTEL = lvw1.SelectedItem.SubItems(5)
tFAX = lvw1.SelectedItem.SubItems(6)
cTerms.Text = lvw1.SelectedItem.SubItems(7)
tOB.Text = lvw1.SelectedItem.SubItems(8)





End Sub

Private Sub tOB_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    A = InStr(1, tOB, ".")
    If A > 0 Then
        KeyAscii = 0
    End If
End If

If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Or KeyAscii = 46 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If

End Sub

Private Sub tSUP_KeyPress(KeyAscii As Integer)
Dim search As String
Dim rs As New Recordset
lvw1.ListItems.Clear
SUP = ""
tADD1 = ""
tADD2 = ""
tADD3 = ""
tADD4 = ""
tTEL = ""
tFAX = ""
cTerms.ListIndex = -1
tOB = ""


If KeyAscii <> 8 Then
    search = "%" & tSUP.Text + Chr(KeyAscii) + "%"
Else
    If Len(tSUP.Text) > 1 Then
    search = "%" & Mid(tSUP.Text, 1, Len(tSUP) - 1) + "%"
    Else
    Exit Sub
    End If
End If


With rs
.Open "select * from supplier where supplier like '" & search & "' order by supplier", c, adOpenDynamic, adLockOptimistic
If .EOF = False Then
    Do While .EOF = False
        Set li = lvw1.ListItems.Add(, , !supplier)
        'lvw1.ListItems(1).ForeColor = vbBlack
        'lvw1.ListItems(1).Bold = False
        'lvw1.ColumnHeaders(1).Width = 1440

        With li
        .SubItems(1) = rs!add1
        .SubItems(2) = rs!add2
        .SubItems(3) = rs!add3
        .SubItems(4) = rs!add4
        .SubItems(5) = rs!tel
        .SubItems(6) = rs!fax
        .SubItems(7) = rs!terms
        .SubItems(8) = rs!OpeningBalance
        
        
        End With
        .MoveNext
    Loop
End If
.Close
End With

End Sub
