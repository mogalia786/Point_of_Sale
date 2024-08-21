VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStockDiscCodes 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assign Discount Codes to Stock"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   ForeColor       =   &H80000016&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H80000016&
      Caption         =   "Stock"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   4455
      Begin MSComctlLib.ListView ListView1 
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2143
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Barcode"
            Object.Width           =   1790
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Tax Code"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   7800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Add to Selected Stock"
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
      MICON           =   "frmStockDiscCodes.frx":0000
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
   Begin VB.Frame Frame4 
      BackColor       =   &H80000016&
      Caption         =   "Select Discount Code to Add"
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
      Height          =   1935
      Left            =   120
      TabIndex        =   11
      Top             =   5760
      Width           =   4455
      Begin MSComctlLib.ListView lvw2 
         Height          =   1455
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2566
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Discount(%)"
            Object.Width           =   1790
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Effective From"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Effective To"
            Object.Width           =   2293
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000016&
      Caption         =   "Available Discount Codes for selected Stock"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   4455
      Begin MSComctlLib.ListView lvw1 
         Height          =   1215
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2143
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Discount(%)"
            Object.Width           =   1790
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Effective From"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Effective To"
            Object.Width           =   2293
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000016&
      Caption         =   "Value"
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
      TabIndex        =   8
      Top             =   1200
      Width           =   4455
      Begin VB.TextBox txtVAL 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "Criteria"
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
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2895
      Begin VB.OptionButton optBC 
         BackColor       =   &H80000016&
         Caption         =   "Bar Code"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optSC 
         BackColor       =   &H80000016&
         Caption         =   "Stock Code"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmStockDiscCodes.frx":001C
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
Attribute VB_Name = "frmStockDiscCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim r As New Recordset
Dim li As ListItem
With r
.Open "select * from disccodes order by disccode", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    Set li = lvw2.ListItems.Add(, , r!disccode)
    With li
    .SubItems(1) = r!discount
    .SubItems(2) = Format(r!Fromdate, "DD/MM/YYYY")
    .SubItems(3) = Format(r!todate, "DD/MM/YYYY")
    End With
    r.MoveNext
Loop
r.Close

End Sub

Private Sub oprSER_Click()


End Sub

Private Sub LaVolpeButton1_Click()
Dim r As New Recordset
Dim SCode As String
Dim BCode As String
Dim SerCode As String
Dim mDiscCode As String
Dim mDiscount As String
Dim mFrom As String
Dim mTo As String
If ListView1.ListItems.Count = 0 Then Exit Sub

If Len(ListView1.SelectedItem) = 0 Then
    MsgBox "Please enter Value Criteria!", vbExclamation, "Status"
    txtVAL.SetFocus
    Exit Sub
End If
If lvw2.ListItems.Count = 0 Then Exit Sub
If Len(lvw2.SelectedItem) = 0 Then Exit Sub
mDiscCode = lvw2.SelectedItem
mDiscount = lvw2.SelectedItem.SubItems(1)
mFrom = Format(lvw2.SelectedItem.SubItems(2), "DD/MM/YYYY")
mTo = Format(lvw2.SelectedItem.SubItems(3), "DD/MM/YYYY")
With r
.Open "select * from stockdiscount where stockcodemain='" & ListView1.SelectedItem & "' and stockcode='" & ListView1.SelectedItem.SubItems(1) & "'and disccode='" & mDiscCode & "'", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = True Then
    r.AddNew
    r!stockcodeMAIN = ListView1.SelectedItem
    r!stockcode = ListView1.SelectedItem.SubItems(1)
    'r!serialnumber = ListView1.SelectedItem.SubItems(2)
    r!disccode = mDiscCode
    r!Fromdate = Format(mFrom, "DD/MM/YYYY")
    r!todate = Format(mTo, "DD/MM/YYYY")
    r!disc = mDiscount
    r.Update
    r.Close
    MsgBox "Discount code successfully added to stock!", vbInformation, "Status"
    txtVAL.Text = ""
    ListView1.ListItems.Clear
    lvw1.ListItems.Clear
    txtVAL.SetFocus
    Exit Sub
Else
    MsgBox "Discount code already exist!", vbExclamation, "Status"
    txtVAL.Text = ""
    ListView1.ListItems.Clear
    lvw1.ListItems.Clear
    txtVAL.SetFocus
    Exit Sub
End If


    
End Sub

Private Sub LaVolpeButton2_Click()
Unload Me

End Sub

Private Sub ListView1_Click()
Dim r As New Recordset
If ListView1.ListItems.Count = 0 Then Exit Sub
If Len(ListView1.SelectedItem) = 0 Then Exit Sub
lvw1.ListItems.Clear
        With r
        .Open "select * from stockdiscount where stockcodemain='" & ListView1.SelectedItem & "' order by disccode", c, adOpenDynamic, adLockOptimistic
        End With
        Do While r.EOF = False
            Set li = lvw1.ListItems.Add(, , r!disccode)
            With li
            .SubItems(1) = r!disc
            .SubItems(2) = Format(r!Fromdate, "DD/MM/YYYY")
            .SubItems(3) = Format(r!todate, "DD/MM/YYYY")
            End With
            r.MoveNext
        Loop
        r.Close

End Sub

Private Sub optBC_Click()
txtVAL.Text = ""
txtVAL.SetFocus
End Sub

Private Sub optSC_Click()
txtVAL.Text = ""
txtVAL.SetFocus
End Sub

Private Sub optSER_Click()
txtVAL.Text = ""
End Sub

Private Sub txtVAL_Change()
lvw1.ListItems.Clear

End Sub

Private Sub txtVAL_KeyPress(KeyAscii As Integer)
Dim search As String
Dim rs As New Recordset
Dim li As ListItem
ListView1.ListItems.Clear


If KeyAscii <> 8 Then
    search = "%" & txtVAL.Text + Chr(KeyAscii) + "%"
Else
    If Len(txtVAL.Text) > 1 Then
    search = "%" & Mid(txtVAL.Text, 1, Len(txtVAL) - 1) + "%"
    Else
    Exit Sub
    End If
End If
If optSC.Value = True Then
    With rs
    .Open "select * from stock where stockcodemain like '" & search & "' order by stockdesc", c, adOpenDynamic, adLockOptimistic
    End With
End If
If optBC.Value = True Then
    With rs
    .Open "select * from stock where stockcode like '" & search & "' order by stockdesc", c, adOpenDynamic, adLockOptimistic
    End With
End If

If rs.EOF = False Then
    Do While rs.EOF = False
        Set li = ListView1.ListItems.Add(, , rs!stockcodeMAIN)
        'lvw1.ListItems(1).ForeColor = vbBlack
        'lvw1.ListItems(1).Bold = False
        'lvw1.ColumnHeaders(1).Width = 1440

        With li
        .SubItems(1) = rs!stockcode
        '.SubItems(2) = rs!serialnumber
        .SubItems(2) = rs!stockdesc
        .SubItems(3) = rs!taxcode
        End With
        rs.MoveNext
    Loop
End If
rs.Close


End Sub

