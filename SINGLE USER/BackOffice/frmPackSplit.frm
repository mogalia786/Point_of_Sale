VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPackSplit 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pack Split"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   6120
      TabIndex        =   13
      Top             =   5040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Split Pack"
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
      MICON           =   "frmPackSplit.frx":0000
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
      Caption         =   "Criteria"
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
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7215
      Begin VB.OptionButton optSN 
         BackColor       =   &H80000016&
         Caption         =   "Serialnumber"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5280
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optBarcode 
         BackColor       =   &H80000016&
         Caption         =   "Barcode"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptStockCode 
         BackColor       =   &H80000016&
         Caption         =   "Stock Code"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000016&
      Caption         =   "Select Pack to Split"
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
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   7215
      Begin MSComctlLib.ListView lvw1 
         Height          =   1575
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2778
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Stock Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Barcode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Serialnumber"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000016&
      Caption         =   "Search..."
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
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   4335
      Begin VB.TextBox txtsn 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox txtBarcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Serialnumber:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Code:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   5040
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
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmPackSplit.frx":001C
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
Attribute VB_Name = "frmPackSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LaVolpeButton1_Click()
Dim r As New Recordset
Dim r1 As New Recordset
Dim r2 As New Recordset
res = MsgBox("Are you sure you wish to Split Pack?", vbYesNo + vbQuestion, "Split Pack?")
If res = vbNo Then Exit Sub
If lvw1.ListItems.Count = 0 Then Exit Sub
If Len(lvw1.SelectedItem) = 0 Then Exit Sub
With r
.Open "select * from packs where pstockcode='" & lvw1.SelectedItem & "' and pbarcode='" & lvw1.SelectedItem.SubItems(1) & "' and pserialnumber ='" & lvw1.SelectedItem.SubItems(2) & "'", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = False Then
    With r1
    .Open "select * from stock where stockcodemain='" & lvw1.SelectedItem & "' and stockcode='" & lvw1.SelectedItem.SubItems(1) & "'", c, adOpenDynamic, adLockOptimistic
    End With
    r1!QTY = CDbl(r1!QTY) - 1
    r1.Update
    r1.Close
End If
With r1
.Open "delete from serialnumber where stockcode='" & lvw1.SelectedItem & "' and serialnumber='" & lvw1.SelectedItem.SubItems(2) & "'", c, adOpenDynamic, adLockOptimistic
End With

Do While r.EOF = False
    With r1
    .Open "select * from serialnumber where stockcode='" & r!cstockcode & "' and serialnumber='" & cserialnumber & "'", c, adOpenDynamic, adLockOptimistic
    End With
    If r1.EOF = True Then
        r1.AddNew
        r1!stockcode = r!cstockcode
        r1!serialnumber = r!cserialnumber
        r1.Update
    End If
    r1.Close
    With r1
    .Open "select * from stock where stockcodemain='" & r!cstockcode & "' and stockcode='" & r!cbarcode & "'", c, adOpenDynamic, adLockOptimistic
    End With
    r1!QTY = CDbl(r1!QTY) + 1
    r1.Update
    r1.Close
    r.MoveNext
Loop
r.Close
With r
.Open "DELETE from packs where pstockcode='" & lvw1.SelectedItem & "' and pbarcode='" & lvw1.SelectedItem.SubItems(1) & "' and pserialnumber ='" & lvw1.SelectedItem.SubItems(2) & "'", c, adOpenDynamic, adLockOptimistic
End With
res = MsgBox("Pack Split Successfull! Do you wish to unpack another?", vbYesNo + vbQuestion, "Unpack another?")
If res = vbYes Then
    txtcode.Text = ""
    txtBarcode.Text = ""
    txtsn.Text = ""
    txtBarcode.Enabled = False
    txtsn.Enabled = False
    txtcode.Enabled = True
    OptStockCode.Value = True
    lvw1.ListItems.Clear
    txtcode.SetFocus
Else
    Unload Me
End If




End Sub

Private Sub LaVolpeButton2_Click()
Unload Me

End Sub

Private Sub Option1_Click()

End Sub

Private Sub Option2_Click()

End Sub

Private Sub optBarcode_Click()
txtcode.Enabled = False
txtBarcode.Enabled = True
txtsn.Enabled = False
txtsn.Text = ""
txtcode = ""
lvw1.ListItems.Clear
txtBarcode.SetFocus

End Sub

Private Sub optDesc_Click()
txtDesc.Enabled = True
txtcode.Enabled = False
txtBarcode.Enabled = False

txtsn.Enabled = False
txtcode = ""
txtBarcode = ""
txtsn.Text = ""
lvw1.ListItems.Clear

txtDesc.SetFocus

End Sub

Private Sub optSN_Click()
txtcode.Enabled = False
txtBarcode.Enabled = False
txtsn.Enabled = True
txtBarcode.Text = ""
txtcode = ""
lvw1.ListItems.Clear
txtsn.SetFocus

End Sub

Private Sub OptStockCode_Click()
txtcode.Enabled = True
txtBarcode.Enabled = False
txtsn.Enabled = False
txtBarcode = ""
txtsn.Text = ""
lvw1.ListItems.Clear

txtcode.SetFocus

End Sub

Private Sub txtBarcode_KeyPress(KeyAscii As Integer)
Dim search As String
Dim rs As New Recordset
Dim lExist As Boolean

lvw1.ListItems.Clear


If KeyAscii <> 8 Then
    search = "%" & txtBarcode.Text + Chr(KeyAscii) + "%"
Else
    If Len(txtBarcode.Text) > 1 Then
    search = "%" & Mid(txtBarcode.Text, 1, Len(txtBarcode) - 1) + "%"
    Else
    Exit Sub
    End If
End If


With rs
.Open "select * from packs where pbarcode like '" & search & "' order by pstockcode", c, adOpenDynamic, adLockOptimistic
If .EOF = False Then
    Do While .EOF = False
        For i = 1 To lvw1.ListItems.Count
            lExist = False
            If lvw1.ListItems.Count > 0 Then
                If lvw1.ListItems(i) = !PSTOCKCODE And lvw1.ListItems(i).SubItems(1) = !pbarcode And lvw1.ListItems(i).SubItems(2) = !pserialnumber Then
                    lExist = True
                End If
            End If
            
        Next i
        If lExist = False Then
            Set li = lvw1.ListItems.Add(, , !PSTOCKCODE)
            'lvw1.ListItems(1).ForeColor = vbBlack
            'lvw1.ListItems(1).Bold = False
            'lvw1.ColumnHeaders(1).Width = 1440
    
            With li
            .SubItems(1) = rs!pbarcode
            .SubItems(2) = rs!pserialnumber
            '.SubItems(3) = rs!stockdesc
            '.SubItems(4) = rs!taxcode
            '.SubItems(5) = rs!unitprice
    
            End With
        End If
        .MoveNext
    Loop
End If
.Close
End With

End Sub


Private Sub txtcode_KeyPress(KeyAscii As Integer)
Dim search As String
Dim rs As New Recordset
Dim lExist As Boolean

lvw1.ListItems.Clear



If KeyAscii <> 8 Then
    search = "%" & txtcode.Text + Chr(KeyAscii) + "%"
Else
    If Len(txtcode.Text) > 1 Then
    search = "%" & Mid(txtcode.Text, 1, Len(txtcode) - 1) + "%"
    Else
    Exit Sub
    End If
End If


With rs
.Open "select * from packs where pstockcode like '" & search & "' order by pstockcode", c, adOpenDynamic, adLockOptimistic
If .EOF = False Then
    Do While .EOF = False
        For i = 1 To lvw1.ListItems.Count
            lExist = False
            If lvw1.ListItems.Count > 0 Then
                If lvw1.ListItems(i) = !PSTOCKCODE And lvw1.ListItems(i).SubItems(1) = !pbarcode And lvw1.ListItems(i).SubItems(2) = !pserialnumber Then
                    lExist = True
                End If
            End If
            
        Next i
        If lExist = False Then
            Set li = lvw1.ListItems.Add(, , !PSTOCKCODE)
            'lvw1.ListItems(1).ForeColor = vbBlack
            'lvw1.ListItems(1).Bold = False
            'lvw1.ColumnHeaders(1).Width = 1440
    
            With li
            .SubItems(1) = rs!pbarcode
            .SubItems(2) = rs!pserialnumber
            '.SubItems(3) = rs!stockdesc
            '.SubItems(4) = rs!taxcode
            '.SubItems(5) = rs!unitprice
    
            End With
        End If
        .MoveNext
    Loop
End If
.Close
End With

End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
Dim search As String
Dim rs As New Recordset
lvw1.ListItems.Clear


If KeyAscii <> 8 Then
    search = "%" & txtDesc.Text + Chr(KeyAscii) + "%"
Else
    If Len(txtDesc.Text) > 1 Then
    search = "%" & Mid(txtDesc.Text, 1, Len(txtDesc) - 1) + "%"
    Else
    Exit Sub
    End If
End If


With rs
.Open "select * from pack where stockdesc like '" & search & "' order by stockcodemain", c, adOpenDynamic, adLockOptimistic
If .EOF = False Then
    Do While .EOF = False
        Set li = lvw1.ListItems.Add(, , !stockcodeMAIN)
        'lvw1.ListItems(1).ForeColor = vbBlack
        'lvw1.ListItems(1).Bold = False
        'lvw1.ColumnHeaders(1).Width = 1440

        With li
        '.SubItems(1) = rs!stockcode
        '.SubItems(2) = rs!serialnumber
        '.SubItems(3) = rs!stockdesc
        '.SubItems(4) = rs!TaxCode
        '.SubItems(5) = rs!unitprice
        End With
        .MoveNext
    Loop
End If
.Close
End With

End Sub

Private Sub txtsn_KeyPress(KeyAscii As Integer)
Dim search As String
Dim rs As New Recordset
Dim lExist As Boolean
lvw1.ListItems.Clear



If KeyAscii <> 8 Then
    search = "%" & txtsn.Text + Chr(KeyAscii) + "%"
Else
    If Len(txtsn.Text) > 1 Then
    search = "%" & Mid(txtsn.Text, 1, Len(txtsn) - 1) + "%"
    Else
    Exit Sub
    End If
End If


With rs
.Open "select * from packs where pserialnumber like '" & search & "' order by pstockcode", c, adOpenDynamic, adLockOptimistic
If .EOF = False Then
    Do While .EOF = False
        For i = 1 To lvw1.ListItems.Count
            lExist = False
            If lvw1.ListItems.Count > 0 Then
                If lvw1.ListItems(i) = !PSTOCKCODE And lvw1.ListItems(i).SubItems(1) = !pbarcode And lvw1.ListItems(i).SubItems(2) = !pserialnumber Then
                    lExist = True
                End If
            End If
            
        Next i
        If lExist = False Then
            Set li = lvw1.ListItems.Add(, , !PSTOCKCODE)
            'lvw1.ListItems(1).ForeColor = vbBlack
            'lvw1.ListItems(1).Bold = False
            'lvw1.ColumnHeaders(1).Width = 1440
    
            With li
            .SubItems(1) = rs!pbarcode
            .SubItems(2) = rs!pserialnumber
            '.SubItems(3) = rs!stockdesc
            '.SubItems(4) = rs!taxcode
            '.SubItems(5) = rs!unitprice
    
            End With
        End If
        .MoveNext
    Loop
End If
.Close
End With

End Sub




