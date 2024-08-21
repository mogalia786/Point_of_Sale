VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTaxCodes 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Tax Codes"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3285
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   3285
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   4080
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "frmTaxCodes.frx":0000
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
   Begin VB.TextBox tTAX 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "Tax Codes"
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
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin MSComctlLib.ListView lvw1 
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tax(%)"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   4080
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "frmTaxCodes.frx":001C
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   975
      Left            =   120
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tax (%):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   -120
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lTC 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   -120
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTaxCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim r As New Recordset
Dim li As ListItem
With r
.Open "select * from taxcode order by taxcode", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    Set li = lvw1.ListItems.Add(, , r!taxcode)
    With li
    .SubItems(1) = r!tax
    End With
    r.MoveNext
Loop
r.Close

End Sub

Private Sub LaVolpeButton1_Click()
Dim r As New Recordset
If Len(ltc) = 0 Then
    MsgBox "Please select tax code you wish to change!", vbExclamation, "Status"
    lvw1.SetFocus
    Exit Sub
End If
If Len(tTAX) = 0 Then
    MsgBox "Please enter Tax(%) value!", vbExclamation, "Status"
    tTAX.SetFocus
    Exit Sub
End If
res = MsgBox("Are you sure you wish to update tax values?", vbYesNo + vbQuestion, "Update?")
If res = vbNo Then Exit Sub
With r
.Open "select * from taxcode where taxcode='" & ltc & "'", c, adOpenDynamic, adLockOptimistic
End With
r!tax = tTAX
r.Update
r.Close
lvw1.ListItems.Clear
tTAX.Text = ""
ltc.Caption = ""
With r
.Open "select * from taxcode order by taxcode", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    Set li = lvw1.ListItems.Add(, , r!taxcode)
    With li
    .SubItems(1) = r!tax
    End With
    r.MoveNext
Loop
r.Close



End Sub

Private Sub LaVolpeButton2_Click()
Unload Me

End Sub

Private Sub lvw1_Click()
ltc.Caption = lvw1.SelectedItem
tTAX.Text = lvw1.SelectedItem.SubItems(1)

End Sub

Private Sub lvw1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
ltc.Caption = lvw1.SelectedItem
tTAX.Text = lvw1.SelectedItem.SubItems(1)

End Sub

Private Sub lvw1_ItemClick(ByVal Item As MSComctlLib.ListItem)
ltc.Caption = lvw1.SelectedItem
tTAX.Text = lvw1.SelectedItem.SubItems(1)

End Sub

Private Sub tTAX_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    A = InStr(1, tTAX, ".")
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
