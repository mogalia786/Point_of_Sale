VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAddSerial 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Serial#"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BackColor       =   &H80000016&
      Caption         =   "Add Serial Numbers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   13
      Top             =   3720
      Width           =   7215
      Begin VB.TextBox txtsn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Top             =   720
         Width           =   2055
      End
      Begin VB.ListBox lSn 
         Appearance      =   0  'Flat
         Height          =   1200
         Left            =   4560
         TabIndex        =   15
         Top             =   240
         Width           =   2535
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton3 
         Height          =   255
         Left            =   3600
         TabIndex        =   14
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "&Add"
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
         MICON           =   "frmAddSerial.frx":0000
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton4 
         Height          =   255
         Left            =   3600
         TabIndex        =   17
         Top             =   960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "&Remove"
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
         MICON           =   "frmAddSerial.frx":001C
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
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Number"
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
         Left            =   -120
         TabIndex        =   18
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000016&
      Caption         =   "Criteria (Click on Item to view details)"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   7215
      Begin MSComctlLib.ListView lvw1 
         Height          =   1215
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2143
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Stock Code"
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
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   7215
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtBarcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   5520
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
         Width           =   1575
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
         Left            =   0
         TabIndex        =   10
         Top             =   360
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
         Left            =   2280
         TabIndex        =   9
         Top             =   360
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
         Left            =   4320
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
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
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.OptionButton OptStockCode 
         BackColor       =   &H80000016&
         Caption         =   "Stock Code"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optDesc 
         BackColor       =   &H80000016&
         Caption         =   "Item Description"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optBarcode 
         BackColor       =   &H80000016&
         Caption         =   "Barcode"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   5880
      TabIndex        =   19
      Top             =   5520
      Width           =   1455
      _ExtentX        =   2566
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
      BCOL            =   14872561
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmAddSerial.frx":0038
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   4320
      TabIndex        =   20
      Top             =   5520
      Width           =   1455
      _ExtentX        =   2566
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
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmAddSerial.frx":0054
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
Attribute VB_Name = "frmAddSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SelSC As String
Private Sub LaVolpeButton1_Click()
Dim r As New Recordset
Dim r10 As New Recordset
Dim r11 As New Recordset
Dim Adj As Double
If lvw1.ListItems.Count = 0 Then
    MsgBox "Please select stock code of item you wish to add serial#!", vbExclamation, "Status"
    txtsn.Text = ""
    If optDesc.Value = True Then
        txtBarcode.Enabled = False
        txtcode.Enabled = False
        txtDesc.Enabled = True
        txtDesc.SetFocus
    ElseIf OptStockCode.Value = True Then
        txtBarcode.Enabled = False
        txtcode.Enabled = True
        txtDesc.Enabled = False
        
        txtcode.SetFocus
    ElseIf optBarcode.Value = True Then
        txtBarcode.Enabled = True
        txtcode.Enabled = False
        txtDesc.Enabled = False
        
        txtBarcode.SetFocus
    End If
    Exit Sub
End If

If Len(lvw1.SelectedItem) = 0 Then
    MsgBox "Please select stock!", vbExclamation, "Status"
    lvw1.SetFocus
    Exit Sub
End If
res = MsgBox("Are you sure you wish to adjust stock level?", vbYesNo + vbQuestion, "Adjust Stock?")
If res = vbNo Then Exit Sub
With r11
.Open "select * from serialnumber", c, adOpenDynamic, adLockOptimistic
End With
        For i = 0 To lSn.ListCount - 1
            r11.AddNew
            r11!stockcode = lvw1.SelectedItem
            r11!serialnumber = lSn.List(i)
            r11.Update
        Next i
r11.Close
res = MsgBox("Stock succesfully adjusted! Do you wish to adjust another?", vbYesNo + vbQuestion, "Adjust another?")
If res = vbYes Then
    lvw1.ListItems.Clear
    txtcode = ""
    txtBarcode = ""
    txtDesc.Text = ""
    lSn.Clear
    txtsn.Text = ""

    If optDesc.Value = True Then
        txtBarcode.Enabled = False
        txtcode.Enabled = False
        txtDesc.Enabled = True
        txtDesc.SetFocus
    ElseIf OptStockCode.Value = True Then
        txtBarcode.Enabled = False
        txtcode.Enabled = True
        txtDesc.Enabled = False
        
        txtcode.SetFocus
    ElseIf optBarcode.Value = True Then
        txtBarcode.Enabled = True
        txtcode.Enabled = False
        txtDesc.Enabled = False
        
        txtBarcode.SetFocus
    End If

Else
    Unload Me
End If



End Sub

Private Sub LaVolpeButton2_Click()
Unload Me

End Sub


Public Sub LaVolpeButton3_Click()
Dim r As New Recordset
If Len(txtsn) = 0 Then Exit Sub
If lvw1.ListItems.Count = 0 Then
    MsgBox "Please select stock code of item you wish to add serial#!", vbExclamation, "Status"
    txtsn.Text = ""
    If optDesc.Value = True Then
        txtBarcode.Enabled = False
        txtcode.Enabled = False
        txtDesc.Enabled = True
        txtDesc.SetFocus
    ElseIf OptStockCode.Value = True Then
        txtBarcode.Enabled = False
        txtcode.Enabled = True
        txtDesc.Enabled = False
        
        txtcode.SetFocus
    ElseIf optBarcode.Value = True Then
        txtBarcode.Enabled = True
        txtcode.Enabled = False
        txtDesc.Enabled = False
        
        txtBarcode.SetFocus
    End If
    Exit Sub
End If
If Len(lvw1.SelectedItem) = 0 Then
    MsgBox "Please select stock code of item you wish to add serial#!", vbExclamation, "Status"
    txtsn.Text = ""
    If optDesc.Value = True Then
        txtBarcode.Enabled = False
        txtcode.Enabled = False
        txtDesc.Enabled = True
        txtDesc.SetFocus
    ElseIf OptStockCode.Value = True Then
        txtBarcode.Enabled = False
        txtcode.Enabled = True
        txtDesc.Enabled = False
        
        txtcode.SetFocus
    ElseIf optBarcode.Value = True Then
        txtBarcode.Enabled = True
        txtcode.Enabled = False
        txtDesc.Enabled = False
        
        txtBarcode.SetFocus
    End If
    Exit Sub
End If
If lSn.ListCount > 0 Then
For i = 0 To lSn.ListCount - 1
    If lSn.List(i) = txtsn.Text Then
        MsgBox "Cannot add serial number as serial number already exist!", vbExclamation, "Status"
        txtsn.Text = ""
        txtsn.SetFocus
        Exit Sub
    End If
Next i
End If

With r
.Open "select * from serialnumber where serialnumber='" & txtsn & "'", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = True Then
    lSn.AddItem txtsn.Text
    txtsn.Text = ""
    txtsn.SetFocus
Else
    MsgBox "Cannot add serial number! This serial number already exists.", vbExclamation, "Status"
    txtsn.Text = ""
    txtsn.SetFocus
End If
r.Close

End Sub

Private Sub LaVolpeButton4_Click()
If lSn.ListCount = 0 Then Exit Sub
    If Len(lSn.List(lSn.ListIndex)) = 0 Then Exit Sub
    lSn.RemoveItem (lSn.ListIndex)
    txtsn.SetFocus
End Sub

Private Sub lvw1_Click()
Dim r As New Recordset
Dim r1 As New Recordset
Dim r2 As New Recordset
Dim QTYSOLD As Double
Dim qtyOnhand As Double
Dim NumofPurchase As Double
Dim TotalCP As Double
Dim TotalPurchased As Double
Dim QtyPacked As Double
Dim QtyUnPacked As Double
Dim mDIFF As Double
Dim TotAdj As Double

If lvw1.ListItems.Count = 0 Then Exit Sub
If Len(lvw1.SelectedItem) = 0 Then Exit Sub

SelSC = lvw1.SelectedItem





End Sub

Private Sub Option1_Click()

End Sub

Private Sub Option2_Click()

End Sub

Private Sub optBarcode_Click()
txtcode.Enabled = False
txtDesc.Enabled = False
txtBarcode.Enabled = True
'txtsn.Enabled = False

txtcode = ""
txtDesc = ""
lvw1.ListItems.Clear
lSC = ""
ldesc = ""
lTAX = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lPACKED = ""
lUNPACKED = ""
ladj = ""
txtBarcode.SetFocus

End Sub

Private Sub optDesc_Click()
txtDesc.Enabled = True
txtcode.Enabled = False
txtBarcode.Enabled = False
'txtsn.Enabled = False
txtcode = ""
txtBarcode = ""
lvw1.ListItems.Clear
lSC = ""
ldesc = ""
lTAX = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lPACKED = ""
lUNPACKED = ""
ladj = ""
txtDesc.SetFocus

End Sub

Private Sub optSN_Click()
txtcode.Enabled = False
txtDesc.Enabled = False
txtBarcode.Enabled = False
txtsn.Enabled = True

txtsn.SetFocus

End Sub

Private Sub OptStockCode_Click()
txtcode.Enabled = True
txtDesc.Enabled = False
txtBarcode.Enabled = False
'txtsn.Enabled = False
txtBarcode = ""
txtDesc = ""
lvw1.ListItems.Clear
lSC = ""
ldesc = ""
lTAX = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lPACKED = ""
lUNPACKED = ""
ladj = ""
txtcode.SetFocus

End Sub

Private Sub txtAdj_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If

End Sub


Private Sub txtBarcode_KeyPress(KeyAscii As Integer)
Dim search As String
Dim rs As New Recordset
lvw1.ListItems.Clear
lSC = ""
ldesc = ""
lTAX = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lPACKED = ""
lUNPACKED = ""
ladj = ""

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
.Open "select distinct(stockcodemain) from stock where stockcode like '" & search & "' order by stockcodeMAIN", c, adOpenDynamic, adLockOptimistic
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
        '.SubItems(4) = rs!taxcode
        '.SubItems(5) = rs!unitprice

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
lSC = ""
ldesc = ""
lTAX = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lPACKED = ""
lUNPACKED = ""
ladj = ""


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
.Open "select distinct(stockcodemain) from stock where stockcodemain like '" & search & "' order by stockcodeMAIN", c, adOpenDynamic, adLockOptimistic
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
        '.SubItems(4) = rs!taxcode
        '.SubItems(5) = rs!unitprice

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
lSC = ""
ldesc = ""
lTAX = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lPACKED = ""
lUNPACKED = ""
ladj = ""

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
.Open "select distinct(stockcodemain) from stock where stockdesc like '" & search & "' order by stockcodemain", c, adOpenDynamic, adLockOptimistic
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

Private Sub txtsn_KeyUp(KeyCode As Integer, Shift As Integer)


If lvw1.ListItems.Count = 0 Then
    MsgBox "Please select stock code of item you wish to add serial#!", vbExclamation, "Status"
    txtsn.Text = ""
    If optDesc.Value = True Then
        txtBarcode.Enabled = False
        txtcode.Enabled = False
        txtDesc.Enabled = True
        txtDesc.SetFocus
    ElseIf OptStockCode.Value = True Then
        txtBarcode.Enabled = False
        txtcode.Enabled = True
        txtDesc.Enabled = False
        
        txtcode.SetFocus
    ElseIf optBarcode.Value = True Then
        txtBarcode.Enabled = True
        txtcode.Enabled = False
        txtDesc.Enabled = False
        
        txtBarcode.SetFocus
    End If
    Exit Sub
End If
If Len(lvw1.SelectedItem) = 0 Then
    MsgBox "Please select stock code of item you wish to add serial#!", vbExclamation, "Status"
    txtsn.Text = ""
    If optDesc.Value = True Then
        txtBarcode.Enabled = False
        txtcode.Enabled = False
        txtDesc.Enabled = True
        txtDesc.SetFocus
    ElseIf OptStockCode.Value = True Then
        txtBarcode.Enabled = False
        txtcode.Enabled = True
        txtDesc.Enabled = False
        
        txtcode.SetFocus
    ElseIf optBarcode.Value = True Then
        txtBarcode.Enabled = True
        txtcode.Enabled = False
        txtDesc.Enabled = False
        
        txtBarcode.SetFocus
    End If
    Exit Sub
End If

    
If KeyCode = vbKeyReturn Then
    LaVolpeButton3_Click
End If

End Sub



