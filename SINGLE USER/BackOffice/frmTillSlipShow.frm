VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTillSlipShow 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View / Print Till Slip Voucher"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Print Voucher"
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
      MICON           =   "frmTillSlipShow.frx":0000
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
      Caption         =   "Select Voucher Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3735
      Begin LVbuttons.LaVolpeButton LaVolpeButton1 
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "&View Vouchers"
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
         MICON           =   "frmTillSlipShow.frx":001C
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
      Begin MSComCtl2.DTPicker cDates 
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   54788097
         CurrentDate     =   38557
      End
   End
   Begin MSComctlLib.ListView lvw1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7858
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Voucher Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "TillID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Stock Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Unit Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Qty"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton3 
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   5400
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
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmTillSlipShow.frx":0038
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
   Begin VB.Label LPrint 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Printing Voucher Please wait........"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   4215
   End
End
Attribute VB_Name = "frmTillSlipShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub ListView1_Click()

End Sub

Private Sub LaVolpeButton1_Click()
Dim r1 As New Recordset
Dim r As New Recordset
Dim li As ListItem
lvw1.ListItems.Clear
        With r1
        .Open "select * from sales where saledate='" & Year(cDates.Value) & "-" & Month(cDates.Value) & "-" & Day(cDates.Value) & " 00:00:00.000' order by saletime", c, adOpenDynamic, adLockOptimistic
        End With
Do While r1.EOF = False
    Set li = lvw1.ListItems.Add(, , r1!saleno)
    With li
    .SubItems(1) = r1!saletime
    .SubItems(2) = r1!TillId
    .SubItems(3) = r1!itemcodemain
    With r
    .Open "select * from stock where stockcodemain='" & r1!itemcodemain & "' and stockcode='" & r1!itemcode & "'", c, adOpenDynamic, adLockOptimistic
    End With
    .SubItems(4) = r!stockdesc
    .SubItems(5) = r1!unitprice
    .SubItems(6) = r1!qty
    .SubItems(7) = r1!total
    End With
    r.Close
    r1.MoveNext
Loop
r1.Close


End Sub

Private Sub LaVolpeButton2_Click()
Dim r As New Recordset
Dim r1 As New Recordset
Dim r2 As New Recordset
Dim r3 As New Recordset
Dim rMain As New Recordset
If lvw1.ListItems.Count = 0 Then Exit Sub
If Len(lvw1.SelectedItem) = 0 Then Exit Sub

With rMain
.Open "select * from sales where saleno='" & lvw1.SelectedItem & "'", c, adOpenDynamic, adLockOptimistic
End With

With r1
.Open "select * from header", c, adOpenDynamic, adLockOptimistic
End With
    Printer.Font.Bold = True
    Printer.Font.Name = "Arial"
    Printer.Font.Size = 10
    'Printer.RightToLeft = True
    For i = 1 To 10
        Column = "line" & i
        If r1.Fields(Column) <> "BLANK" Then
            Printer.Print r1.Fields(Column)
        Else
            Printer.Print ""
        End If
    Next i
    r1.Close
    LPrint.Visible = True
    Printer.Print rMain!saletime
    Printer.Print "Till" & vbTab & vbTab & ":" & rMain!TillId
    Printer.Print "Invoice No." & vbTab & ":" & rMain!saleno
    Printer.Font.Size = 8
    Printer.Print "------------------------------------------------------------------"
    Printer.Print "QTY  StockCode  Serial#       Barcode"
    Printer.Print "UNIT INC  DESCRIP               TOTAL INC"
    Printer.Print "------------------------------------------------------------------"
    Printer.Print ""
    
    TotNumItems = 0
    Do While rMain.EOF = False
        Printer.Print rMain!qty & "  " & rMain!itemcodemain & "  " & rMain!serialnumber & "  " & rMain!itemcode
        TotNumItems = TotNumItems + rMain!qty
        With r1
        .Open "select * from stock where stockcodemain='" & rMain!itemcodemain & "' and stockcode='" & rMain!itemcode & "'", c, adOpenDynamic, adLockOptimistic
        End With
        
        If Len(r1!stockdesc) <= 22 Then
            Select Case Len(r1!stockdesc)
                Case 1
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & "*********************" & vbTab & "R" & rMain!total
                Case 2
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & "********************" & vbTab & "R" & rMain!total
                Case 3
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & "*******************" & vbTab & "R" & rMain!total
                Case 4
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & "******************" & vbTab & "R" & rMain!total
                Case 5
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & "*****************" & vbTab & "R" & rMain!total
                Case 6
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & "****************" & vbTab & "R" & rMain!total
                Case 7
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & "***************" & vbTab & "R" & rMain!total
                Case 8
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & "**************" & vbTab & "R" & rMain!total
                Case 9
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & "*************" & vbTab & "R" & rMain!total
                Case 10
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & "************" & vbTab & "R" & rMain!total
                Case 11
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & "***********" & vbTab & "R" & rMain!total
                Case 12
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & "**********" & vbTab & "R" & rMain!total
                Case 13
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & "*********" & vbTab & "R" & rMain!total
                Case 14
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & "********" & vbTab & "R" & rMain!total
                Case 15
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & "*******" & vbTab & "R" & rMain!total
                Case 16
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & "******" & vbTab & "R" & rMain!total
                Case 17
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & "*****" & vbTab & "R" & rMain!total
                Case 18
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & "****" & vbTab & "R" & rMain!total
                Case 19
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & "***" & vbTab & "R" & rMain!total
                Case 20
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & "**" & vbTab & "R" & rMain!total
                Case 21
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & "*" & vbTab & "R" & rMain!total
                Case 22
                    Printer.Print "R" & rMain!unitprice & "  " & r1!stockdesc & vbTab & "R" & rMain!total
            End Select
        Else
            Printer.Print "R" & rMain!unitprice & "  " & Mid(r1!stockdesc, 1, 22) & vbTab & "R" & rMain!total
        End If
        r1.Close
        Printer.Print "------------------------------------------------------------------"
        Printer.Print ""
    rMain.MoveNext
    Loop
    rMain.Close
    Printer.Print "TOTAL ITEMS ON INVOICE : ********" & TotNumItems & "********"
    Printer.Print ""
    Printer.Print "------------------------------------------------------------------"
    Printer.Print vbTab & vbTab & "VAT SUMMARY"
    Printer.Print "CODE" & vbTab & "%" & vbTab & "GOODS" & vbTab & "VAT"
    Printer.Print "------------------------------------------------------------------"
    With r1
    .Open "select * from taxcode order by taxcode", c, adOpenDynamic, adLockOptimistic
    End With
    Do While r1.EOF = False
        With r
        .Open "select * from sales where taxcode='" & r1!taxcode & "' and saleno='" & lvw1.SelectedItem & "'", c, adOpenDynamic, adLockOptimistic
        End With
            TVAT = 0
            tgoods = 0
            Do While r.EOF = False
                tgoods = Format(tgoods + CDbl(r!total), "#####.00")
                TVAT = Format(TVAT + CDbl(r!vat), "#####.00")
                r.MoveNext
            Loop
            r.Close
            If TVAT > 0 Then
                Printer.Print r1!taxcode & vbTab & r1!tax & vbTab & Format((tgoods * 100 / (r1!tax + 100)), "#####.00") & vbTab & Format(TVAT, "#####.00")
                GTGoods = CDbl(GTGoods) + Format((CDbl(tgoods * 100 / (r1!tax + 100))), "#####.00")
                GTVAT = CDbl(GTVAT) + CDbl(TVAT)
            End If
            
        r1.MoveNext
    Loop
    r1.Close
    Printer.Print "------------------------------------------------------------------"

    Printer.Print "TOTALS" & vbTab & Format(GTGoods, "#####.00") & vbTab & Format(GTVAT, "#####.00")
    Printer.Print "------------------------------------------------------------------"
    Printer.Print "------------------------------------------------------------------"
    With r1
    .Open "select * from sales where saleno='" & lvw1.SelectedItem & "'", c, adOpenDynamic, adLockOptimistic
    End With
    Do While r1.EOF = False
        rTOTDISC = rTOTDISC + CDbl(r1!TOTDISC)
        r1.MoveNext
    Loop
    r1.Close
    Printer.Print "TOTAL DISCOUNT" & vbTab & vbTab & "R" & Format(rTOTDISC, "#####.00")
    'Printer.Print "------------------------------------------------------------------"
    'Printer.Print vbTab & vbTab & "PAYMENT SUMMARY"
    'Printer.Print "AMOUNT DUE" & vbTab & vbTab & vbTab & "R" & lTOTAL; ""
    'Printer.Print ""
    'Printer.Print "CASH" & vbTab & vbTab & vbTab & vbTab & "R" & Format(txtTEN.Text, "#####.00")
    'Printer.Print "CHANGE" & vbTab & vbTab & vbTab & "R" & Format(lchange, "#####.00")
    Printer.Print "------------------------------------------------------------------"
    Printer.Print "------------------------------------------------------------------"
    
    
    With r1
    .Open "select * from footer", c, adOpenDynamic, adLockOptimistic
    End With
    Printer.Font.Bold = True
    
    Printer.Font.Name = "Arial"
    Printer.Font.Size = 8
    For i = 1 To 5
        Column = "line" & i
        If r1.Fields(Column) <> "BLANK" Then
            Printer.Print r1.Fields(Column)
        Else
            Printer.Print ""
        End If
    Next i
    r1.Close
    Printer.EndDoc
LPrint.Visible = False

End Sub

Private Sub LaVolpeButton3_Click()
Unload Me

End Sub


