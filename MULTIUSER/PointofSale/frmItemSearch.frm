VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmItemSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Search"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   2640
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
      MICON           =   "frmItemSearch.frx":0000
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
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   7215
      Begin VB.OptionButton optBarcode 
         Caption         =   "Barcode"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton optDesc 
         Caption         =   "Item Description"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton OptStockCode 
         Caption         =   "Stock Code"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
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
      ForeColor       =   &H00400000&
      Height          =   3015
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   6855
      Begin MSComctlLib.ListView lvw1 
         Height          =   2415
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4260
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
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
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
      ForeColor       =   &H00400000&
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   4335
      Begin VB.TextBox txtBarcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox txtcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   360
         Width           =   2775
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
         TabIndex        =   13
         Top             =   1560
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
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
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
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   4920
      Picture         =   "frmItemSearch.frx":001C
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "frmItemSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Paint()
If frmMain.lCRIT = "Stock Code" Then
    optDesc.Value = True
End If
If frmMain.lCRIT = "Barcode" Then
    optBarcode.Value = True
End If
If frmMain.lCRIT = "Serial#" Then
    OptStockCode.Value = True
End If

'If frmMain.lCRIT = "Stock Code" Then
 '   txtdesc.SetFocus
'End If
'If frmMain.lCRIT = "Barcode" Then
'    txtBarcode.SetFocus
'End If
'If frmMain.lCRIT = "Serial#" Then
'    txtCode.SetFocus
'End If
End Sub

Private Sub LaVolpeButton1_Click()
Unload Me

End Sub

Private Sub lvw1_KeyDown(KeyCode As Integer, Shift As Integer)
If lvw1.ListItems.Count = 0 Then Exit Sub
If Len(lvw1.SelectedItem) = 0 Then Exit Sub

If KeyCode = vbKeyReturn Then
    If optDesc.Value Then
        frmMain.lCRIT = "Stock Code"
        frmMain.txtcode.Text = lvw1.SelectedItem
    End If
    If OptStockCode.Value Then
        frmMain.lCRIT = "Stock Code"
        frmMain.txtcode.Text = lvw1.SelectedItem
    End If
    If optBarcode.Value Then
        frmMain.lCRIT = "Barcode"
        frmMain.txtcode.Text = lvw1.SelectedItem.SubItems(1)
    End If
    Unload Me
End If

    
End Sub

Private Sub Option1_Click()

End Sub

Private Sub Option2_Click()

End Sub

Private Sub optBarcode_Click()
frmMain.lCRIT = "Barcode"
txtcode.Enabled = False
txtDesc.Enabled = False
txtBarcode.Enabled = True
'txtsn.Enabled = False
txtDesc.Text = ""
txtcode.Text = ""
lvw1.ListItems.Clear

txtBarcode.SetFocus

End Sub

Private Sub optDesc_Click()
frmMain.lCRIT = "Stock Code"
txtDesc.Enabled = True
txtcode.Enabled = False
txtBarcode.Enabled = False
'txtsn.Enabled = False
txtBarcode.Text = ""
txtcode.Text = ""
lvw1.ListItems.Clear


txtDesc.SetFocus

End Sub

Private Sub optSN_Click()
frmMain.lCRIT = "Serial#"
txtcode.Enabled = False
txtDesc.Enabled = False
txtBarcode.Enabled = False
txtsn.Enabled = True

txtsn.SetFocus

End Sub

Private Sub OptStockCode_Click()
frmMain.lCRIT = "Stock Code"
txtcode.Enabled = True
txtDesc.Enabled = False
txtBarcode.Enabled = False
'txtsn.Enabled = False
txtDesc.Text = ""
txtBarcode.Text = ""
lvw1.ListItems.Clear


txtcode.SetFocus

End Sub

Private Sub txtBarcode_KeyPress(KeyAscii As Integer)
Dim search As String
Dim rs As New Recordset
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
.Open "select * from stock where stockcode like '" & search & "' order by stockcode", c, adOpenDynamic, adLockOptimistic
If .EOF = False Then
    Do While .EOF = False
        Set li = lvw1.ListItems.Add(, , !stockcodemain)
        'lvw1.ListItems(1).ForeColor = vbBlack
        'lvw1.ListItems(1).Bold = False
        'lvw1.ColumnHeaders(1).Width = 1440

        With li
        .SubItems(1) = rs!stockcode
        '.SubItems(2) = rs!serialnumber
        .SubItems(2) = rs!STOCKDESC
        End With
        .MoveNext
    Loop
End If
.Close
End With

End Sub


Private Sub txtcode_KeyPress(KeyAscii As Integer)
Dim search As String
Dim rs As New Recordset
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
.Open "select * from stock where stockcodemain like '" & search & "' order by stockcode", c, adOpenDynamic, adLockOptimistic
If .EOF = False Then
    Do While .EOF = False
        Set li = lvw1.ListItems.Add(, , !stockcodemain)
        'lvw1.ListItems(1).ForeColor = vbBlack
        'lvw1.ListItems(1).Bold = False
        'lvw1.ColumnHeaders(1).Width = 1440

        With li
        .SubItems(1) = rs!stockcode
        '.SubItems(2) = rs!serialnumber
        .SubItems(2) = rs!STOCKDESC
        End With
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
.Open "select * from stock where stockdesc like '" & search & "' order by stockdesc", c, adOpenDynamic, adLockOptimistic
If .EOF = False Then
    Do While .EOF = False
        Set li = lvw1.ListItems.Add(, , !stockcodemain)
        'lvw1.ListItems(1).ForeColor = vbBlack
        'lvw1.ListItems(1).Bold = False
        'lvw1.ColumnHeaders(1).Width = 1440

        With li
        .SubItems(1) = rs!stockcode
        '.SubItems(2) = rs!serialnumber
        .SubItems(2) = rs!STOCKDESC
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
.Open "select * from stock where serialnumber like '" & search & "' order by stockcode", c, adOpenDynamic, adLockOptimistic
If .EOF = False Then
    Do While .EOF = False
        Set li = lvw1.ListItems.Add(, , !stockcodemain)
        'lvw1.ListItems(1).ForeColor = vbBlack
        'lvw1.ListItems(1).Bold = False
        'lvw1.ColumnHeaders(1).Width = 1440

        With li
        .SubItems(1) = rs!stockcode
        .SubItems(2) = rs!serialnumber
        .SubItems(3) = rs!STOCKDESC
        End With
        .MoveNext
    Loop
End If
.Close
End With

End Sub


