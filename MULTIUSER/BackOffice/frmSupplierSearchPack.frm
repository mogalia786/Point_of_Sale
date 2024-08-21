VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSupplierSearchPack 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplier Search............."
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   8400
      TabIndex        =   4
      Top             =   4560
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
      BCOL            =   14872561
      FCOL            =   0
      FCOLO           =   33023
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmSupplierSearchPack.frx":0000
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
      TabIndex        =   2
      Top             =   1200
      Width           =   9375
      Begin MSComctlLib.ListView lvw1 
         Height          =   2775
         Left            =   120
         TabIndex        =   3
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
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      Begin VB.TextBox tSUP 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmSupplierSearchPack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LaVolpeButton1_Click()
Unload Me

End Sub

Private Sub lvw1_Click()
If lvw1.ListItems.Count = 0 Then Exit Sub
If Len(lvw1.SelectedItem) = 0 Then Exit Sub

frmCreatePackNew.tSUP = lvw1.SelectedItem
frmCreatePackNew.tADD1 = lvw1.SelectedItem.SubItems(1)
frmCreatePackNew.tADD2 = lvw1.SelectedItem.SubItems(2)
frmCreatePackNew.tADD3 = lvw1.SelectedItem.SubItems(3)
frmCreatePackNew.tADD4 = lvw1.SelectedItem.SubItems(4)
frmCreatePackNew.tTEL = lvw1.SelectedItem.SubItems(5)
frmCreatePackNew.tFAX = lvw1.SelectedItem.SubItems(6)
frmCreatePackNew.cTerms.Text = lvw1.SelectedItem.SubItems(7)
frmCreatePackNew.lBAL.Text = lvw1.SelectedItem.SubItems(8)

Unload Me



End Sub

Private Sub tSUP_KeyPress(KeyAscii As Integer)
Dim search As String
Dim rs As New Recordset
lvw1.ListItems.Clear


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
