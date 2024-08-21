VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCreatePackExisting 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Pack from existing stock"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   9000
      TabIndex        =   9
      Top             =   4440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Create Pack"
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
      MICON           =   "frmCreatePackExisting.frx":0000
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
      Caption         =   "Scan / Enter serial number of item to add to pack"
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
      Height          =   855
      Left            =   120
      TabIndex        =   20
      Top             =   240
      Width           =   5535
      Begin VB.TextBox txtsn 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   5295
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000016&
      Caption         =   "Select items to add to pack"
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
      Height          =   3135
      Left            =   5760
      TabIndex        =   19
      Top             =   1200
      Width           =   4335
      Begin MSComctlLib.ListView lvw1 
         Height          =   2655
         Left            =   120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   4683
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
            Text            =   "Serial Number"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000016&
      Caption         =   "Pack Details"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   3135
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   5535
      Begin VB.TextBox tSell 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3240
         TabIndex        =   8
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox tCOST 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         TabIndex        =   6
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox tSC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox tBC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox tSN 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox tDESC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1440
         Width           =   2415
      End
      Begin VB.ComboBox cTaxCode 
         Height          =   315
         ItemData        =   "frmCreatePackExisting.frx":001C
         Left            =   4200
         List            =   "frmCreatePackExisting.frx":001E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000016&
         Caption         =   "Calculate Selling Price using MarkUp"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   2280
         Width           =   4455
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Code:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Number:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   -240
         TabIndex        =   16
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Code:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3240
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "2. Recommended Selling price (Incl.):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   -600
         TabIndex        =   13
         Top             =   2640
         Width           =   3735
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1. Cost price per Item (Excl.):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   -600
         TabIndex        =   12
         Top             =   1800
         Width           =   3015
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   7800
      TabIndex        =   10
      Top             =   4440
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
      MICON           =   "frmCreatePackExisting.frx":0020
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
Attribute VB_Name = "frmCreatePackExisting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Dim r As New Recordset
Dim r1 As New Recordset
Dim nnCINC As Double
If Len(cTaxCode.Text) = 0 Then
    MsgBox "Please select tax code!"
    cTaxCode.SetFocus
    Exit Sub
End If

If Len(tCOST.Text) = 0 Then
    If Check1.Value Then
        tSell.Text = ""
    Else
        tSell.Text = SP
    End If
    
Exit Sub
End If
If Check1.Value Then
    With r
    .Open "select * from markup", c, adOpenDynamic, adLockOptimistic
    End With
    With r1
    .Open "select * from taxcode where taxcode='" & cTaxCode.Text & "'", c, adOpenDynamic, adLockOptimistic
    End With
    x1tax = CDbl(r1!tax)
    r1.Close
    'nnCINC = CDbl(tCOST) + (CDbl(tCOST) * x1tax / 100)
    nnCINC = CDbl(tCOST)
    
    tSell.Text = nnCINC * (CDbl(r!markup) + 100) / 100
    tSell.Text = CDbl(tSell.Text) + (CDbl(tSell) * x1tax / 100)
    tSell = Format(tSell, "#####.00")
    
    
Else
   tSell = Format(SP, "#####.00")
End If

End Sub

Private Sub cTaxCode_Click()
Dim r As New Recordset
Dim r1 As New Recordset
Dim nnCINC As Double

If Len(tCOST.Text) = 0 Then
    If Check1.Value Then
        tSell.Text = ""
    Else
        tSell.Text = SP
    End If
    
Exit Sub
End If
If Check1.Value Then
    With r
    .Open "select * from markup", c, adOpenDynamic, adLockOptimistic
    End With
    With r1
    .Open "select * from taxcode where taxcode='" & cTaxCode.Text & "'", c, adOpenDynamic, adLockOptimistic
    End With
    If r1.EOF = False Then
    x1tax = CDbl(r1!tax)
    r1.Close
    nnCINC = CDbl(tCOST)
    
    tSell.Text = nnCINC * (CDbl(r!markup) + 100) / 100
    tSell.Text = CDbl(tSell.Text) + (CDbl(tSell) * x1tax / 100)
    tSell = Format(tSell, "#####.00")
    Else
    tsel = ""
    End If
Else
   tSell = Format(SP, "#####.00")
End If

End Sub

Private Sub Form_Load()
Dim r As New Recordset
With r
.Open "select * from taxcode order by taxcode", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    cTaxCode.AddItem r!taxcode
    r.MoveNext
Loop
r.Close
cTaxCode.ListIndex = 1

End Sub

Private Sub LaVolpeButton1_Click()
Dim r As New Recordset
Dim r1 As New Recordset

If Len(tSC) = 0 Then
    MsgBox "Please enter Stock Code!", vbInformation, "Status"
    tSC.SetFocus
    Exit Sub
End If
If Len(tBC) = 0 Then
    tBC = "N/A"
End If
If Len(tSN) = 0 Then
    tSN = "N/A"
End If
If Len(tDESC) = 0 Then
    MsgBox "Please enter Description!", vbInformation, "Status"
    tDESC.SetFocus
    Exit Sub
End If
If Len(tCOST) = 0 Then
    MsgBox "Please enter CostPrice!", vbInformation, "Status"
    tCOST.SetFocus
    Exit Sub
End If
If Len(cTaxCode.Text) = 0 Then
    MsgBox "Please select Tax Code!", vbInformation, "Status"
    cTaxCode.SetFocus
    Exit Sub
End If
If Len(tSell.Text) = 0 Then
    MsgBox "Please enter selling price!", vbExclamation, "Status"
    tSell.SetFocus
    Exit Sub
End If


For i = 1 To lvw1.ListItems.Count
    With r
    .Open "select * from packs", c, adOpenDynamic, adLockOptimistic
    End With
    r.AddNew
    r!PSTOCKCODE = UCase(tSC.Text)
    r!pbarcode = tBC.Text
    r!pserialnumber = tSN.Text
    r!cstockcode = lvw1.ListItems(i)
    r!cbarcode = lvw1.ListItems(i).SubItems(1)
    r!cserialnumber = lvw1.ListItems(i).SubItems(2)
    r.Update
    r.Close
Next i

For i = 1 To lvw1.ListItems.Count
    With r
    .Open "delete from serialnumber where serialnumber='" & lvw1.ListItems(i).SubItems(2) & "'", c, adOpenDynamic, adLockOptimistic
    End With
Next i
If tSN <> "N/A" Then
    With r
    .Open "select * from serialnumber", c, adOpenDynamic, adLockOptimistic
    End With
    r.AddNew
    r!stockcode = UCase(tSC)
    r!serialnumber = tSN
    r.Update
    r.Close
End If
For i = 1 To lvw1.ListItems.Count
    With r
    .Open "select * from stock where stockcodemain='" & lvw1.ListItems(i) & "' and stockcode='" & lvw1.ListItems(i).SubItems(1) & "'", c, adOpenDynamic, adLockOptimistic
    End With
    r!QTY = r!QTY - 1
    r.Update
    r.Close
Next i

With r
.Open "select * from stock WHERE STOCKCODEMAIN='" & tSC & "' AND STOCKCODE='" & tBC & "'", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = True Then
r.AddNew
r!stockcodeMAIN = UCase(tSC)
r!stockcode = tBC
'r!serialnumber = tSN
r!stockdesc = UCase(tDESC)
r!taxcode = cTaxCode.Text
r!unitprice = Format(tSell, "####.00")
r!QTY = "1"
r.Update
Else

r!stockcodeMAIN = UCase(tSC)
r!stockcode = tBC
'r!serialnumber = tSN
r!stockdesc = UCase(tDESC)
r!taxcode = cTaxCode.Text
r!unitprice = Format(tSell, "####.00")
r!QTY = CDbl(r!QTY) + 1
r.Update
End If
r.Close

    


    With r
    .Open "select * from stockpurchasehistory", c, adOpenDynamic, adLockOptimistic
    End With
    r.AddNew
    r!stockcodeMAIN = UCase(tSC)
    r!stockcode = tBC
    r!serialnumber = tSN
    r!supplier = "Internal"
    r!costofItemEXC = Format(tCOST, "#####.00")
    With r1
    .Open "select * from taxcode where taxcode='" & cTaxCode & "'", c, adOpenDynamic, adLockOptimistic
    End With
    
    r!vatinput = Format((CDbl(tCOST) * CDbl(r1!tax) / 100), "#####.00")
    r!costofiteminc = Format(CDbl(tCOST) + CDbl(r!vatinput), "#####.00")
    r!sellingpriceINc = Format(tSell, "#####.00")
    r!datepurchased = Format(Date, "DD/MM/YYYY")
    r!qtypurchased = "1"
    r.Update
    r.Close
With r
.Open "SELECT * FROM PACKSLIST WHERE PSTOCKCODE='" & tSC.Text & "'", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = True Then
    r.AddNew
    r!PSTOCKCODE = UCase(tSC)
    r.Update
End If
r.Close

res = MsgBox("Pack successfully created! Do you wish to create another?", vbYesNo + vbQuestion, "Create another?")
If res = vbYes Then
    lvw1.ListItems.Clear
    txtsn.Text = ""
    tSC = ""
    tBC = ""
    tSN = ""
    tDESC = ""
    tCOST = ""
    Check1.Value = False
    tSell = ""
    cTaxCode.ListIndex = -1
    txtsn.SetFocus
Else
    Unload Me
End If


End Sub

Private Sub LaVolpeButton2_Click()
Unload Me

End Sub

Private Sub tCOST_Change()
Dim r As New Recordset
Dim r1 As New Recordset
Dim nnCINC As Double
If Len(cTaxCode.Text) = 0 Then
    MsgBox "Please select tax code!"
    cTaxCode.SetFocus
    Exit Sub
End If

If Len(tCOST.Text) = 0 Then
    If Check1.Value Then
        tSell.Text = ""
    Else
        tSell.Text = SP
    End If
    
Exit Sub
End If
If Check1.Value Then
    With r
    .Open "select * from markup", c, adOpenDynamic, adLockOptimistic
    End With
    With r1
    .Open "select * from taxcode where taxcode='" & cTaxCode.Text & "'", c, adOpenDynamic, adLockOptimistic
    End With
    x1tax = CDbl(r1!tax)
    r1.Close
    nnCINC = CDbl(tCOST)
    
    tSell.Text = nnCINC * (CDbl(r!markup) + 100) / 100
    tSell.Text = CDbl(tSell.Text) + (CDbl(tSell) * x1tax / 100)
    tSell = Format(tSell, "#####.00")
    
    
Else
   tSell = Format(SP, "#####.00")
End If


End Sub

Private Sub tCOST_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    A = InStr(1, tCOST, ".")
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

Private Sub tSell_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    A = InStr(1, tSell, ".")
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

Private Sub txtsn_KeyUp(KeyCode As Integer, Shift As Integer)
Dim r As New Recordset
Dim r1 As New Recordset
Dim li As ListItem
If KeyCode = vbKeyReturn Then
    If Len(txtsn) = 0 Then Exit Sub
    With r
    .Open "select * from packs where pserialnumber='" & txtsn & "'", c, adOpenDynamic, adLockOptimistic
    End With
    If r.EOF = False Then
        MsgBox "This item cannot be added to pack as it is already a pack!", vbExclamation, "Status"
        txtsn.Text = ""
        txtsn.SetFocus
        Exit Sub
    End If
    r.Close
    
    With r
    .Open "select * from SERIALNUMBER where serialnumber='" & txtsn & "'", c, adOpenDynamic, adLockOptimistic
    End With
    If r.EOF = False Then
        For i = 1 To lvw1.ListItems.Count
            If lvw1.ListItems(i).SubItems(2) = txtsn Then
                MsgBox "Item already selected!", vbExclamation, "Status"
                txtsn.Text = ""
                txtsn.SetFocus
                Exit Sub
            End If
        Next i
            
        Set li = lvw1.ListItems.Add(, , r!stockcode)
        With r1
        .Open "select * from stock where stockcodemain='" & r!stockcode & "'", c, adOpenDynamic, adLockOptimistic
        End With
        
        li.SubItems(1) = r1!stockcode
        li.SubItems(2) = r!serialnumber
        txtsn.Text = ""
        txtsn.SetFocus
    Else
        MsgBox "Item not found!"
        txtsn.Text = ""
        txtsn.SetFocus
        Exit Sub
    End If
    r.Close
End If

End Sub
