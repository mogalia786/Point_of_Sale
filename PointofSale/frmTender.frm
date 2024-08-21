VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmTender 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tender"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSCommLib.MSComm ms1 
      Left            =   2040
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame4 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3720
      TabIndex        =   16
      Top             =   5880
      Width           =   855
      Begin VB.Label E 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Esc"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Update Sale ONLY"
      ForeColor       =   &H00400000&
      Height          =   615
      Left            =   1920
      TabIndex        =   15
      Top             =   5880
      Width           =   1695
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "F10"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Update Sale && Print"
      ForeColor       =   &H00400000&
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   5880
      Width           =   1695
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "F1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame f2 
      Caption         =   "Cheque Number"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   4215
      Begin VB.TextBox txtCHQ 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame f1 
      Caption         =   "Credit Card Number"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   4215
      Begin VB.TextBox txtCC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Payment Type"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      Begin VB.OptionButton optCHQ 
         Caption         =   "Cheque"
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
         Left            =   2640
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optCREDIT 
         Caption         =   "Credit Card"
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
         Left            =   1320
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton optCASH 
         Caption         =   "Cash"
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
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.TextBox txtTEN 
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
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   2400
      TabIndex        =   9
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tendered"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label lCHANGE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   2400
      TabIndex        =   10
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label lTOTAL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1000.00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   2400
      TabIndex        =   8
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000009&
      FillStyle       =   0  'Solid
      Height          =   2415
      Left            =   -120
      Top             =   3360
      Width           =   4815
   End
End
Attribute VB_Name = "frmTender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
txtTEN.SetFocus
End Sub

Private Sub Form_Load()
lTOTAL.Caption = Format(frmMain.lTOTAL.Caption, "####.00")
lchange.Caption = ""
txtTEN.Text = ""
End Sub

Private Sub optCASH_Click()
f1.Enabled = False
f2.Enabled = False
txtTEN.Text = ""
lchange.Caption = ""
txtTEN.Enabled = True
txtTEN.SetFocus

End Sub

Private Sub optCASH_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
'ms1.PortOpen = True

        'ms1.Output = Chr(27) + "@"
    'ms1.PortOpen = False
    Unload Me
End If

End Sub

Private Sub optCHQ_Click()
f1.Enabled = False
f2.Enabled = True
txtCHQ.SetFocus
txtTEN.Text = lTOTAL.Caption
lchange.Caption = "0.00"
txtTEN.Enabled = False

End Sub

Private Sub optCHQ_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
'ms1.PortOpen = True

        'ms1.Output = Chr(27) + "@"
 'ms1.PortOpen = False
    Unload Me
End If

End Sub

Private Sub optCREDIT_Click()
f1.Enabled = True
f2.Enabled = False
txtCC.SetFocus
txtTEN.Text = lTOTAL.Caption
lchange.Caption = "0.00"
txtTEN.Enabled = False

End Sub

Private Sub optCREDIT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
'ms1.PortOpen = True

        'ms1.Output = Chr(27) + "@"
 'ms1.PortOpen = False
    Unload Me
End If

End Sub

Private Sub txtCC_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
'ms1.PortOpen = True

        'ms1.Output = Chr(27) + "@"

    Unload Me
End If
Dim r As New Recordset
Dim R1 As New Recordset
Dim r2 As New Recordset
Dim TotNumItems As Integer
Dim rTOTDISC As Double
If KeyCode = vbKeyEscape Then
    'ms1.PortOpen = True

    'ms1.Output = Chr(27) + "@"
    'ms1.PortOpen = False
    Unload Me
End If
If KeyCode = vbKeyF10 Then

    If Len(txtTEN) = 0 Then
        MsgBox "Please enter tendered amount!"
        txtTEN.SetFocus
        Exit Sub
    End If
    If CDbl(txtTEN) < CDbl(lTOTAL) Then
        MsgBox "Tendered amount cannot be less than Total amount!", vbExclamation, "Error!"
        txtTEN.Text = ""
        lchange.Caption = "0.00"
        txtTEN.SetFocus
        Exit Sub
    End If
    
    If CDbl(lchange) < 0 Then
        MsgBox "Tendered amount is less then amount due!", vbExclamation, "status"
        txtTEN.Text = ""
        lchange = ""
        txtTEN.SetFocus
        Exit Sub
    End If
    If CDbl(lTOTAL) = 0 Then
        MsgBox "Cannot generate sale!", vbExclamation, "Status"
        txtTEN.SetFocus
        Exit Sub
    End If
    lineN = 1
    With R1
    .Open "invoicenumber", HC, adOpenDynamic, adLockOptimistic
    End With
    invnum = R1!invoicenum + 1
    R1!invoicenum = invnum
    R1.Update
    R1.Close
    If CDbl(invnum) < 10 Then
        invnum = "00000" & invnum
    End If
    If CDbl(invnum) > 9 And CDbl(invnum) < 100 Then
        invnum = "0000" & invnum
    End If
    If CDbl(invnum) > 99 And CDbl(invnum) < 1000 Then
        invnum = "000" & invnum
    End If
    If CDbl(invnum) > 999 And CDbl(invnum) < 10000 Then
        invnum = "00" & invnum
    End If
    If CDbl(invnum) > 9999 And CDbl(invnum) < 100000 Then
        invnum = "0" & invnum
    End If
    If CDbl(invnum) > 99999 And CDbl(invnum) < 1000000 Then
        invnum = invnum
    End If
    
    invnum = TillId & "/" & invnum
    With R1
    .Open "sale", HC, adOpenDynamic, adLockOptimistic
    End With
    Do While R1.EOF = False
        With r
        .Open "select * from sales", c, adOpenDynamic, adLockOptimistic
        End With
        r.AddNew
        r!saledate = Format(Date, "DD/MM/YYYY")
        r!saletime = Time
        r!TillId = TillId
        r!saleno = invnum
        r!itemcodemain = R1!itemcodemain
        r!itemcode = R1!itemcode
        r!serialnumber = R1!serialnumber
        r!taxcode = R1!taxcode
        r!disccode = R1!disccode
        r!qty = R1!qty
        r!UNITPRICE = R1!UNITPRICE
        r!total = Format(CDbl(r!qty * r!UNITPRICE), "#####.00")
        r!vat = Format(CDbl(R1!vat), "#####.00")
        r!discadded = Format(R1!discadded, "#####.00")
        r!TOTDISC = Format(R1!TOTDISC, "#####.00")
        r!UNITDISC = Format(R1!UNITDISC, "#####.00")
        With r2
        .Open "select * from stock where stockcodemain='" & R1!itemcodemain & "' and stockcode='" & R1!itemcode & "'", c, adOpenDynamic, adLockOptimistic
        End With
        
        r!department = r2!department
        r2.Update
        r2.Close
        
        r.Update
        r.Close
        With r
        .Open "Select * from stock where stockcode='" & R1!itemcode & "' and stockcodemain='" & R1!itemcodemain & "'", c, adOpenDynamic, adLockOptimistic
        End With
        r!qty = r!qty - R1!qty
        r.Update
        r.Close
        R1.MoveNext
    Loop
    R1.Close
    
    With R1
    .Open "delete * from sale", HC, adOpenDynamic, adLockOptimistic
    End With
    frmMain.T1.Text = ""
    frmMain.txtcode.Text = ""
    frmMain.lTOTAL = "0,00"
    frmMain.txtDesc.Caption = "Scan Item"
    MyChange = lchange.Caption
  
    Unload Me
    frmChange.Show vbModal
End If

'///////////////////////////////////////////PRINT TILL SLIP
If KeyCode = vbKeyF1 Then
    If Len(txtTEN) = 0 Then
        MsgBox "Please enter tendered amount!"
        txtTEN.SetFocus
        Exit Sub
    End If
    If CDbl(txtTEN) < CDbl(lTOTAL) Then
        MsgBox "Tendered amount cannot be less than Total amount!", vbExclamation, "Error!"
        txtTEN.Text = ""
        lchange.Caption = "0.00"
        txtTEN.SetFocus
        Exit Sub
    End If
    
    If CDbl(lchange) < 0 Then
        MsgBox "Tendered amount is less then amount due!", vbExclamation, "status"
        txtTEN.Text = ""
        lchange = ""
        txtTEN.SetFocus
        Exit Sub
    End If
    If CDbl(lTOTAL) = 0 Then
        MsgBox "Cannot generate sale!", vbExclamation, "Status"
        txtTEN.SetFocus
        Exit Sub
    End If
    With R1
    .Open "invoicenumber", HC, adOpenDynamic, adLockOptimistic
    End With
    lineN = 1
    invnum = R1!invoicenum + 1
    R1!invoicenum = invnum
    R1.Update
    R1.Close
    If CDbl(invnum) < 10 Then
        invnum = "00000" & invnum
    End If
    If CDbl(invnum) > 9 And CDbl(invnum) < 100 Then
        invnum = "0000" & invnum
    End If
    If CDbl(invnum) > 99 And CDbl(invnum) < 1000 Then
        invnum = "000" & invnum
    End If
    If CDbl(invnum) > 999 And CDbl(invnum) < 10000 Then
        invnum = "00" & invnum
    End If
    If CDbl(invnum) > 9999 And CDbl(invnum) < 100000 Then
        invnum = "0" & invnum
    End If
    If CDbl(invnum) > 99999 And CDbl(invnum) < 1000000 Then
        invnum = invnum
    End If
    
    invnum = TillId & "/" & invnum
    With R1
    .Open "sale", HC, adOpenDynamic, adLockOptimistic
    End With
    Do While R1.EOF = False
        With r
        .Open "select * from sales", c, adOpenDynamic, adLockOptimistic
        End With
        r.AddNew
        r!saledate = Format(Date, "DD/MM/YYYY")
        r!saletime = Time
        r!TillId = TillId
        r!saleno = invnum
        r!itemcodemain = R1!itemcodemain
        r!itemcode = R1!itemcode
        r!serialnumber = R1!serialnumber
        r!taxcode = R1!taxcode
        r!disccode = R1!disccode
        r!qty = R1!qty
        r!UNITPRICE = R1!UNITPRICE
        r!total = Format(CDbl(r!qty * r!UNITPRICE), "#####.00")
        r!vat = Format(CDbl(R1!vat), "#####.00")
        r!discadded = Format(R1!discadded, "#####.00")
        r!TOTDISC = Format(R1!TOTDISC, "#####.00")
        r!UNITDISC = Format(R1!UNITDISC, "#####.00")
        With r2
        .Open "select * from stock where stockcodemain='" & R1!itemcodemain & "' and stockcode='" & R1!itemcode & "'", c, adOpenDynamic, adLockOptimistic
        End With
        
        r!department = r2!department
        r2.Update
        r2.Close
        r.Update
        r.Close
        With r
        .Open "Select * from stock where stockcode='" & R1!itemcode & "' and stockcodemain='" & R1!itemcodemain & "'", c, adOpenDynamic, adLockOptimistic
        End With
        r!qty = r!qty - R1!qty
        r.Update
        r.Close
        R1.MoveNext
    Loop
    R1.Close
    OpenTill
    '////////////////////////////////Printing
    With R1
    .Open "select * from header", c, adOpenDynamic, adLockOptimistic
    End With
    Printer.Font.Bold = True
    Printer.Font.Name = "Arial"
    Printer.Font.Size = 10
    'Printer.RightToLeft = True
    For i = 1 To 10
        Column = "line" & i
        If R1.Fields(Column) <> "BLANK" Then
            Printer.Print R1.Fields(Column)
        Else
            Printer.Print ""
        End If
    Next i
    R1.Close
    
    Printer.Print Now
    Printer.Print "Till" & vbTab & vbTab & ":" & TillId
    Printer.Print "Invoice No." & vbTab & ":" & invnum
    Printer.Print "Cashier" & vbTab & vbTab & ":" & UCase(CurrentUser)
    Printer.Font.Size = 8
    Printer.Print "------------------------------------------------------------------"
    Printer.Print "QTY  StockCode  Serial#       Barcode"
    Printer.Print "UNIT INC  DESCRIP               TOTAL INC"
    Printer.Print "------------------------------------------------------------------"
    Printer.Print ""
    
    With R1
    .Open "sale", HC, adOpenDynamic, adLockOptimistic
    End With
    TotNumItems = 0
    Do While R1.EOF = False
        Printer.Print R1!qty & "  " & R1!itemcodemain & "  " & R1!serialnumber & "  " & R1!itemcode
        TotNumItems = TotNumItems + R1!qty
        If Len(R1!Description) <= 22 Then
            Select Case Len(R1!Description)
                Case 1
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*********************" & vbTab & "R" & R1!total
                Case 2
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "********************" & vbTab & "R" & R1!total
                Case 3
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*******************" & vbTab & "R" & R1!total
                Case 4
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "******************" & vbTab & "R" & R1!total
                Case 5
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*****************" & vbTab & "R" & R1!total
                Case 6
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "****************" & vbTab & "R" & R1!total
                Case 7
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "***************" & vbTab & "R" & R1!total
                Case 8
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "**************" & vbTab & "R" & R1!total
                Case 9
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*************" & vbTab & "R" & R1!total
                Case 10
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "************" & vbTab & "R" & R1!total
                Case 11
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "***********" & vbTab & "R" & R1!total
                Case 12
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "**********" & vbTab & "R" & R1!total
                Case 13
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*********" & vbTab & "R" & R1!total
                Case 14
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "********" & vbTab & "R" & R1!total
                Case 15
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*******" & vbTab & "R" & R1!total
                Case 16
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "******" & vbTab & "R" & R1!total
                Case 17
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*****" & vbTab & "R" & R1!total
                Case 18
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "****" & vbTab & "R" & R1!total
                Case 19
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "***" & vbTab & "R" & R1!total
                Case 20
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "**" & vbTab & "R" & R1!total
                Case 21
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*" & vbTab & "R" & R1!total
                Case 22
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & vbTab & "R" & R1!total
            End Select
        Else
            Printer.Print "R" & R1!UNITPRICE & "  " & Mid(R1!Description, 1, 22) & vbTab & "R" & R1!total
        End If
        Printer.Print "------------------------------------------------------------------"
        Printer.Print ""
    R1.MoveNext
    Loop
    R1.Close
    Printer.Print "TOTAL ITEMS ON INVOICE : ********" & TotNumItems & "********"
    Printer.Print ""
    Printer.Print "------------------------------------------------------------------"
    Printer.Print vbTab & vbTab & "VAT SUMMARY"
    Printer.Print "CODE" & vbTab & "%" & vbTab & "GOODS" & vbTab & "VAT"
    Printer.Print "------------------------------------------------------------------"
    With R1
    .Open "select * from taxcode order by taxcode", c, adOpenDynamic, adLockOptimistic
    End With
    Do While R1.EOF = False
        With r
        .Open "select * from sale where taxcode='" & R1!taxcode & "'", HC, adOpenDynamic, adLockOptimistic
        End With
            tvat = 0
            tgoods = 0
            Do While r.EOF = False
                tgoods = Format(tgoods + CDbl(r!total), "#####.00")
                tvat = Format(tvat + CDbl(r!vat), "#####.00")
                r.MoveNext
            Loop
            r.Close
            If tvat > 0 Then
                Printer.Print R1!taxcode & vbTab & R1!tax & vbTab & Format((tgoods * 100 / (R1!tax + 100)), "#####.00") & vbTab & Format(tvat, "#####.00")
                GTGoods = CDbl(GTGoods) + Format((CDbl(tgoods * 100 / (R1!tax + 100))), "#####.00")
                GTVat = CDbl(GTVat) + CDbl(tvat)
            End If
            
        R1.MoveNext
    Loop
    R1.Close
    Printer.Print "------------------------------------------------------------------"

    Printer.Print "TOTALS" & vbTab & Format(GTGoods, "#####.00") & vbTab & Format(GTVat, "#####.00")
    Printer.Print "------------------------------------------------------------------"
    Printer.Print "------------------------------------------------------------------"
    With R1
    .Open "select * from sale", HC, adOpenDynamic, adLockOptimistic
    End With
    Do While R1.EOF = False
        rTOTDISC = rTOTDISC + CDbl(R1!TOTDISC)
        R1.MoveNext
    Loop
    R1.Close
    Printer.Print "TOTAL DISCOUNT" & vbTab & vbTab & "R" & Format(rTOTDISC, "#####.00")
    Printer.Print "------------------------------------------------------------------"
    Printer.Print vbTab & vbTab & "PAYMENT SUMMARY"
    Printer.Print "AMOUNT DUE" & vbTab & vbTab & vbTab & "R" & lTOTAL; ""
    Printer.Print ""
    Printer.Print "CASH" & vbTab & vbTab & vbTab & vbTab & "R" & Format(txtTEN.Text, "#####.00")
    Printer.Print "CHANGE" & vbTab & vbTab & vbTab & "R" & Format(lchange, "#####.00")
    Printer.Print "------------------------------------------------------------------"
    Printer.Print "------------------------------------------------------------------"
    
    
    With R1
    .Open "select * from footer", c, adOpenDynamic, adLockOptimistic
    End With
    Printer.Font.Bold = True
    
    Printer.Font.Name = "Arial"
    Printer.Font.Size = 8
    For i = 1 To 5
        Column = "line" & i
        If R1.Fields(Column) <> "BLANK" Then
            Printer.Print R1.Fields(Column)
        Else
            Printer.Print ""
        End If
    Next i
    R1.Close
    Printer.EndDoc
    
    With R1
    .Open "delete * from sale", HC, adOpenDynamic, adLockOptimistic
    End With
    frmMain.T1.Text = ""
    frmMain.txtcode.Text = ""
    frmMain.lTOTAL = "0,00"
    frmMain.txtDesc.Caption = "Scan Item"
    MyChange = lchange.Caption
  
    Unload Me
    frmChange.Show vbModal
End If

End Sub

Private Sub txtCHQ_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
'ms1.PortOpen = True

        'ms1.Output = Chr(27) + "@"
 'ms1.PortOpen = False
    Unload Me
End If
Dim r As New Recordset
Dim R1 As New Recordset
Dim r2 As New Recordset
Dim TotNumItems As Integer
Dim rTOTDISC As Double
If KeyCode = vbKeyEscape Then
    'ms1.PortOpen = True

    'ms1.Output = Chr(27) + "@"
    'ms1.PortOpen = False
    Unload Me
End If
If KeyCode = vbKeyF10 Then

    If Len(txtTEN) = 0 Then
        MsgBox "Please enter tendered amount!"
        txtTEN.SetFocus
        Exit Sub
    End If
    If CDbl(txtTEN) < CDbl(lTOTAL) Then
        MsgBox "Tendered amount cannot be less than Total amount!", vbExclamation, "Error!"
        txtTEN.Text = ""
        lchange.Caption = "0.00"
        txtTEN.SetFocus
        Exit Sub
    End If
    
    If CDbl(lchange) < 0 Then
        MsgBox "Tendered amount is less then amount due!", vbExclamation, "status"
        txtTEN.Text = ""
        lchange = ""
        txtTEN.SetFocus
        Exit Sub
    End If
    If CDbl(lTOTAL) = 0 Then
        MsgBox "Cannot generate sale!", vbExclamation, "Status"
        txtTEN.SetFocus
        Exit Sub
    End If
    lineN = 1
    With R1
    .Open "invoicenumber", HC, adOpenDynamic, adLockOptimistic
    End With
    invnum = R1!invoicenum + 1
    R1!invoicenum = invnum
    R1.Update
    R1.Close
    If CDbl(invnum) < 10 Then
        invnum = "00000" & invnum
    End If
    If CDbl(invnum) > 9 And CDbl(invnum) < 100 Then
        invnum = "0000" & invnum
    End If
    If CDbl(invnum) > 99 And CDbl(invnum) < 1000 Then
        invnum = "000" & invnum
    End If
    If CDbl(invnum) > 999 And CDbl(invnum) < 10000 Then
        invnum = "00" & invnum
    End If
    If CDbl(invnum) > 9999 And CDbl(invnum) < 100000 Then
        invnum = "0" & invnum
    End If
    If CDbl(invnum) > 99999 And CDbl(invnum) < 1000000 Then
        invnum = invnum
    End If
    
    invnum = TillId & "/" & invnum
    With R1
    .Open "sale", HC, adOpenDynamic, adLockOptimistic
    End With
    Do While R1.EOF = False
        With r
        .Open "select * from sales", c, adOpenDynamic, adLockOptimistic
        End With
        r.AddNew
        r!saledate = Format(Date, "DD/MM/YYYY")
        r!saletime = Time
        r!TillId = TillId
        r!saleno = invnum
        r!itemcodemain = R1!itemcodemain
        r!itemcode = R1!itemcode
        r!serialnumber = R1!serialnumber
        r!taxcode = R1!taxcode
        r!disccode = R1!disccode
        r!qty = R1!qty
        r!UNITPRICE = R1!UNITPRICE
        r!total = Format(CDbl(r!qty * r!UNITPRICE), "#####.00")
        r!vat = Format(CDbl(R1!vat), "#####.00")
        r!discadded = Format(R1!discadded, "#####.00")
        r!TOTDISC = Format(R1!TOTDISC, "#####.00")
        r!UNITDISC = Format(R1!UNITDISC, "#####.00")
        With r2
        .Open "select * from stock where stockcodemain='" & R1!itemcodemain & "' and stockcode='" & R1!itemcode & "'", c, adOpenDynamic, adLockOptimistic
        End With
        
        r!department = r2!department
        r2.Update
        r2.Close
        
        r.Update
        r.Close
        With r
        .Open "Select * from stock where stockcode='" & R1!itemcode & "' and stockcodemain='" & R1!itemcodemain & "'", c, adOpenDynamic, adLockOptimistic
        End With
        r!qty = r!qty - R1!qty
        r.Update
        r.Close
        R1.MoveNext
    Loop
    R1.Close
    
    With R1
    .Open "delete * from sale", HC, adOpenDynamic, adLockOptimistic
    End With
    frmMain.T1.Text = ""
    frmMain.txtcode.Text = ""
    frmMain.lTOTAL = "0,00"
    frmMain.txtDesc.Caption = "Scan Item"
    MyChange = lchange.Caption
  
    Unload Me
    frmChange.Show vbModal
End If

'///////////////////////////////////////////PRINT TILL SLIP
If KeyCode = vbKeyF1 Then
    If Len(txtTEN) = 0 Then
        MsgBox "Please enter tendered amount!"
        txtTEN.SetFocus
        Exit Sub
    End If
    If CDbl(txtTEN) < CDbl(lTOTAL) Then
        MsgBox "Tendered amount cannot be less than Total amount!", vbExclamation, "Error!"
        txtTEN.Text = ""
        lchange.Caption = "0.00"
        txtTEN.SetFocus
        Exit Sub
    End If
    
    If CDbl(lchange) < 0 Then
        MsgBox "Tendered amount is less then amount due!", vbExclamation, "status"
        txtTEN.Text = ""
        lchange = ""
        txtTEN.SetFocus
        Exit Sub
    End If
    If CDbl(lTOTAL) = 0 Then
        MsgBox "Cannot generate sale!", vbExclamation, "Status"
        txtTEN.SetFocus
        Exit Sub
    End If
    With R1
    .Open "invoicenumber", HC, adOpenDynamic, adLockOptimistic
    End With
    lineN = 1
    invnum = R1!invoicenum + 1
    R1!invoicenum = invnum
    R1.Update
    R1.Close
    If CDbl(invnum) < 10 Then
        invnum = "00000" & invnum
    End If
    If CDbl(invnum) > 9 And CDbl(invnum) < 100 Then
        invnum = "0000" & invnum
    End If
    If CDbl(invnum) > 99 And CDbl(invnum) < 1000 Then
        invnum = "000" & invnum
    End If
    If CDbl(invnum) > 999 And CDbl(invnum) < 10000 Then
        invnum = "00" & invnum
    End If
    If CDbl(invnum) > 9999 And CDbl(invnum) < 100000 Then
        invnum = "0" & invnum
    End If
    If CDbl(invnum) > 99999 And CDbl(invnum) < 1000000 Then
        invnum = invnum
    End If
    
    invnum = TillId & "/" & invnum
    With R1
    .Open "sale", HC, adOpenDynamic, adLockOptimistic
    End With
    Do While R1.EOF = False
        With r
        .Open "select * from sales", c, adOpenDynamic, adLockOptimistic
        End With
        r.AddNew
        r!saledate = Format(Date, "DD/MM/YYYY")
        r!saletime = Time
        r!TillId = TillId
        r!saleno = invnum
        r!itemcodemain = R1!itemcodemain
        r!itemcode = R1!itemcode
        r!serialnumber = R1!serialnumber
        r!taxcode = R1!taxcode
        r!disccode = R1!disccode
        r!qty = R1!qty
        r!UNITPRICE = R1!UNITPRICE
        r!total = Format(CDbl(r!qty * r!UNITPRICE), "#####.00")
        r!vat = Format(CDbl(R1!vat), "#####.00")
        r!discadded = Format(R1!discadded, "#####.00")
        r!TOTDISC = Format(R1!TOTDISC, "#####.00")
        r!UNITDISC = Format(R1!UNITDISC, "#####.00")
        With r2
        .Open "select * from stock where stockcodemain='" & R1!itemcodemain & "' and stockcode='" & R1!itemcode & "'", c, adOpenDynamic, adLockOptimistic
        End With
        
        r!department = r2!department
        r2.Update
        r2.Close
        r.Update
        r.Close
        With r
        .Open "Select * from stock where stockcode='" & R1!itemcode & "' and stockcodemain='" & R1!itemcodemain & "'", c, adOpenDynamic, adLockOptimistic
        End With
        r!qty = r!qty - R1!qty
        r.Update
        r.Close
        R1.MoveNext
    Loop
    R1.Close
    OpenTill
    '////////////////////////////////Printing
    With R1
    .Open "select * from header", c, adOpenDynamic, adLockOptimistic
    End With
    Printer.Font.Bold = True
    Printer.Font.Name = "Arial"
    Printer.Font.Size = 10
    'Printer.RightToLeft = True
    For i = 1 To 10
        Column = "line" & i
        If R1.Fields(Column) <> "BLANK" Then
            Printer.Print R1.Fields(Column)
        Else
            Printer.Print ""
        End If
    Next i
    R1.Close
    
    Printer.Print Now
    Printer.Print "Till" & vbTab & vbTab & ":" & TillId
    Printer.Print "Invoice No." & vbTab & ":" & invnum
    Printer.Print "Cashier" & vbTab & vbTab & ":" & UCase(CurrentUser)
    Printer.Font.Size = 8
    Printer.Print "------------------------------------------------------------------"
    Printer.Print "QTY  StockCode  Serial#       Barcode"
    Printer.Print "UNIT INC  DESCRIP               TOTAL INC"
    Printer.Print "------------------------------------------------------------------"
    Printer.Print ""
    
    With R1
    .Open "sale", HC, adOpenDynamic, adLockOptimistic
    End With
    TotNumItems = 0
    Do While R1.EOF = False
        Printer.Print R1!qty & "  " & R1!itemcodemain & "  " & R1!serialnumber & "  " & R1!itemcode
        TotNumItems = TotNumItems + R1!qty
        If Len(R1!Description) <= 22 Then
            Select Case Len(R1!Description)
                Case 1
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*********************" & vbTab & "R" & R1!total
                Case 2
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "********************" & vbTab & "R" & R1!total
                Case 3
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*******************" & vbTab & "R" & R1!total
                Case 4
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "******************" & vbTab & "R" & R1!total
                Case 5
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*****************" & vbTab & "R" & R1!total
                Case 6
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "****************" & vbTab & "R" & R1!total
                Case 7
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "***************" & vbTab & "R" & R1!total
                Case 8
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "**************" & vbTab & "R" & R1!total
                Case 9
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*************" & vbTab & "R" & R1!total
                Case 10
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "************" & vbTab & "R" & R1!total
                Case 11
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "***********" & vbTab & "R" & R1!total
                Case 12
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "**********" & vbTab & "R" & R1!total
                Case 13
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*********" & vbTab & "R" & R1!total
                Case 14
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "********" & vbTab & "R" & R1!total
                Case 15
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*******" & vbTab & "R" & R1!total
                Case 16
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "******" & vbTab & "R" & R1!total
                Case 17
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*****" & vbTab & "R" & R1!total
                Case 18
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "****" & vbTab & "R" & R1!total
                Case 19
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "***" & vbTab & "R" & R1!total
                Case 20
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "**" & vbTab & "R" & R1!total
                Case 21
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*" & vbTab & "R" & R1!total
                Case 22
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & vbTab & "R" & R1!total
            End Select
        Else
            Printer.Print "R" & R1!UNITPRICE & "  " & Mid(R1!Description, 1, 22) & vbTab & "R" & R1!total
        End If
        Printer.Print "------------------------------------------------------------------"
        Printer.Print ""
    R1.MoveNext
    Loop
    R1.Close
    Printer.Print "TOTAL ITEMS ON INVOICE : ********" & TotNumItems & "********"
    Printer.Print ""
    Printer.Print "------------------------------------------------------------------"
    Printer.Print vbTab & vbTab & "VAT SUMMARY"
    Printer.Print "CODE" & vbTab & "%" & vbTab & "GOODS" & vbTab & "VAT"
    Printer.Print "------------------------------------------------------------------"
    With R1
    .Open "select * from taxcode order by taxcode", c, adOpenDynamic, adLockOptimistic
    End With
    Do While R1.EOF = False
        With r
        .Open "select * from sale where taxcode='" & R1!taxcode & "'", HC, adOpenDynamic, adLockOptimistic
        End With
            tvat = 0
            tgoods = 0
            Do While r.EOF = False
                tgoods = Format(tgoods + CDbl(r!total), "#####.00")
                tvat = Format(tvat + CDbl(r!vat), "#####.00")
                r.MoveNext
            Loop
            r.Close
            If tvat > 0 Then
                Printer.Print R1!taxcode & vbTab & R1!tax & vbTab & Format((tgoods * 100 / (R1!tax + 100)), "#####.00") & vbTab & Format(tvat, "#####.00")
                GTGoods = CDbl(GTGoods) + Format((CDbl(tgoods * 100 / (R1!tax + 100))), "#####.00")
                GTVat = CDbl(GTVat) + CDbl(tvat)
            End If
            
        R1.MoveNext
    Loop
    R1.Close
    Printer.Print "------------------------------------------------------------------"

    Printer.Print "TOTALS" & vbTab & Format(GTGoods, "#####.00") & vbTab & Format(GTVat, "#####.00")
    Printer.Print "------------------------------------------------------------------"
    Printer.Print "------------------------------------------------------------------"
    With R1
    .Open "select * from sale", HC, adOpenDynamic, adLockOptimistic
    End With
    Do While R1.EOF = False
        rTOTDISC = rTOTDISC + CDbl(R1!TOTDISC)
        R1.MoveNext
    Loop
    R1.Close
    Printer.Print "TOTAL DISCOUNT" & vbTab & vbTab & "R" & Format(rTOTDISC, "#####.00")
    Printer.Print "------------------------------------------------------------------"
    Printer.Print vbTab & vbTab & "PAYMENT SUMMARY"
    Printer.Print "AMOUNT DUE" & vbTab & vbTab & vbTab & "R" & lTOTAL; ""
    Printer.Print ""
    Printer.Print "CASH" & vbTab & vbTab & vbTab & vbTab & "R" & Format(txtTEN.Text, "#####.00")
    Printer.Print "CHANGE" & vbTab & vbTab & vbTab & "R" & Format(lchange, "#####.00")
    Printer.Print "------------------------------------------------------------------"
    Printer.Print "------------------------------------------------------------------"
    
    
    With R1
    .Open "select * from footer", c, adOpenDynamic, adLockOptimistic
    End With
    Printer.Font.Bold = True
    
    Printer.Font.Name = "Arial"
    Printer.Font.Size = 8
    For i = 1 To 5
        Column = "line" & i
        If R1.Fields(Column) <> "BLANK" Then
            Printer.Print R1.Fields(Column)
        Else
            Printer.Print ""
        End If
    Next i
    R1.Close
    Printer.EndDoc
    
    With R1
    .Open "delete * from sale", HC, adOpenDynamic, adLockOptimistic
    End With
    frmMain.T1.Text = ""
    frmMain.txtcode.Text = ""
    frmMain.lTOTAL = "0,00"
    frmMain.txtDesc.Caption = "Scan Item"
    MyChange = lchange.Caption
  
    Unload Me
    frmChange.Show vbModal
End If

End Sub

Private Sub txtTEN_Change()
Dim Tot As Double
Dim Ten As Double
Dim Change As Double

If Len(lTOTAL) = 0 Then
    Tot = 0
Else
    Tot = CDbl(lTOTAL)
End If
If Len(txtTEN) = 0 Then
    Ten = 0
Else
    Ten = CDbl(txtTEN)
End If
Change = Ten - Tot
If Change >= 0 Then
lchange = Format(Change, "####.00##")
Else
lchange = "0.00"
End If

End Sub

Private Sub txtTEN_KeyDown(KeyCode As Integer, Shift As Integer)
Dim r As New Recordset
Dim R1 As New Recordset
Dim r2 As New Recordset
Dim TotNumItems As Integer
Dim rTOTDISC As Double
If KeyCode = vbKeyEscape Then
    'ms1.PortOpen = True

    'ms1.Output = Chr(27) + "@"
    'ms1.PortOpen = False
    Unload Me
End If
If KeyCode = vbKeyF10 Then

    If Len(txtTEN) = 0 Then
        MsgBox "Please enter tendered amount!"
        txtTEN.SetFocus
        Exit Sub
    End If
    If CDbl(txtTEN) < CDbl(lTOTAL) Then
        MsgBox "Tendered amount cannot be less than Total amount!", vbExclamation, "Error!"
        txtTEN.Text = ""
        lchange.Caption = "0.00"
        txtTEN.SetFocus
        Exit Sub
    End If
    
    If CDbl(lchange) < 0 Then
        MsgBox "Tendered amount is less then amount due!", vbExclamation, "status"
        txtTEN.Text = ""
        lchange = ""
        txtTEN.SetFocus
        Exit Sub
    End If
    If CDbl(lTOTAL) = 0 Then
        MsgBox "Cannot generate sale!", vbExclamation, "Status"
        txtTEN.SetFocus
        Exit Sub
    End If
    lineN = 1
    With R1
    .Open "invoicenumber", HC, adOpenDynamic, adLockOptimistic
    End With
    invnum = R1!invoicenum + 1
    R1!invoicenum = invnum
    R1.Update
    R1.Close
    If CDbl(invnum) < 10 Then
        invnum = "00000" & invnum
    End If
    If CDbl(invnum) > 9 And CDbl(invnum) < 100 Then
        invnum = "0000" & invnum
    End If
    If CDbl(invnum) > 99 And CDbl(invnum) < 1000 Then
        invnum = "000" & invnum
    End If
    If CDbl(invnum) > 999 And CDbl(invnum) < 10000 Then
        invnum = "00" & invnum
    End If
    If CDbl(invnum) > 9999 And CDbl(invnum) < 100000 Then
        invnum = "0" & invnum
    End If
    If CDbl(invnum) > 99999 And CDbl(invnum) < 1000000 Then
        invnum = invnum
    End If
    
    invnum = TillId & "/" & invnum
    With R1
    .Open "sale", HC, adOpenDynamic, adLockOptimistic
    End With
    Do While R1.EOF = False
        With r
        .Open "select * from sales", c, adOpenDynamic, adLockOptimistic
        End With
        r.AddNew
        r!saledate = Format(Date, "DD/MM/YYYY")
        r!saletime = Time
        r!TillId = TillId
        r!saleno = invnum
        r!itemcodemain = R1!itemcodemain
        r!itemcode = R1!itemcode
        r!serialnumber = R1!serialnumber
        r!taxcode = R1!taxcode
        r!disccode = R1!disccode
        r!qty = R1!qty
        r!UNITPRICE = R1!UNITPRICE
        r!total = Format(CDbl(r!qty * r!UNITPRICE), "#####.00")
        r!vat = Format(CDbl(R1!vat), "#####.00")
        r!discadded = Format(R1!discadded, "#####.00")
        r!TOTDISC = Format(R1!TOTDISC, "#####.00")
        r!UNITDISC = Format(R1!UNITDISC, "#####.00")
        With r2
        .Open "select * from stock where stockcodemain='" & R1!itemcodemain & "' and stockcode='" & R1!itemcode & "'", c, adOpenDynamic, adLockOptimistic
        End With
        
        r!department = r2!department
        r2.Update
        r2.Close
        
        r.Update
        r.Close
        With r
        .Open "Select * from stock where stockcode='" & R1!itemcode & "' and stockcodemain='" & R1!itemcodemain & "'", c, adOpenDynamic, adLockOptimistic
        End With
        r!qty = r!qty - R1!qty
        r.Update
        r.Close
        R1.MoveNext
    Loop
    R1.Close
    
    With R1
    .Open "delete * from sale", HC, adOpenDynamic, adLockOptimistic
    End With
    frmMain.T1.Text = ""
    frmMain.txtcode.Text = ""
    frmMain.lTOTAL = "0,00"
    frmMain.txtDesc.Caption = "Scan Item"
    MyChange = lchange.Caption
  
    Unload Me
    frmChange.Show vbModal
End If

'///////////////////////////////////////////PRINT TILL SLIP
If KeyCode = vbKeyF1 Then
    If Len(txtTEN) = 0 Then
        MsgBox "Please enter tendered amount!"
        txtTEN.SetFocus
        Exit Sub
    End If
    If CDbl(txtTEN) < CDbl(lTOTAL) Then
        MsgBox "Tendered amount cannot be less than Total amount!", vbExclamation, "Error!"
        txtTEN.Text = ""
        lchange.Caption = "0.00"
        txtTEN.SetFocus
        Exit Sub
    End If
    
    If CDbl(lchange) < 0 Then
        MsgBox "Tendered amount is less then amount due!", vbExclamation, "status"
        txtTEN.Text = ""
        lchange = ""
        txtTEN.SetFocus
        Exit Sub
    End If
    If CDbl(lTOTAL) = 0 Then
        MsgBox "Cannot generate sale!", vbExclamation, "Status"
        txtTEN.SetFocus
        Exit Sub
    End If
    With R1
    .Open "invoicenumber", HC, adOpenDynamic, adLockOptimistic
    End With
    lineN = 1
    invnum = R1!invoicenum + 1
    R1!invoicenum = invnum
    R1.Update
    R1.Close
    If CDbl(invnum) < 10 Then
        invnum = "00000" & invnum
    End If
    If CDbl(invnum) > 9 And CDbl(invnum) < 100 Then
        invnum = "0000" & invnum
    End If
    If CDbl(invnum) > 99 And CDbl(invnum) < 1000 Then
        invnum = "000" & invnum
    End If
    If CDbl(invnum) > 999 And CDbl(invnum) < 10000 Then
        invnum = "00" & invnum
    End If
    If CDbl(invnum) > 9999 And CDbl(invnum) < 100000 Then
        invnum = "0" & invnum
    End If
    If CDbl(invnum) > 99999 And CDbl(invnum) < 1000000 Then
        invnum = invnum
    End If
    
    invnum = TillId & "/" & invnum
    With R1
    .Open "sale", HC, adOpenDynamic, adLockOptimistic
    End With
    Do While R1.EOF = False
        With r
        .Open "select * from sales", c, adOpenDynamic, adLockOptimistic
        End With
        r.AddNew
        r!saledate = Format(Date, "DD/MM/YYYY")
        r!saletime = Time
        r!TillId = TillId
        r!saleno = invnum
        r!itemcodemain = R1!itemcodemain
        r!itemcode = R1!itemcode
        r!serialnumber = R1!serialnumber
        r!taxcode = R1!taxcode
        r!disccode = R1!disccode
        r!qty = R1!qty
        r!UNITPRICE = R1!UNITPRICE
        r!total = Format(CDbl(r!qty * r!UNITPRICE), "#####.00")
        r!vat = Format(CDbl(R1!vat), "#####.00")
        r!discadded = Format(R1!discadded, "#####.00")
        r!TOTDISC = Format(R1!TOTDISC, "#####.00")
        r!UNITDISC = Format(R1!UNITDISC, "#####.00")
        With r2
        .Open "select * from stock where stockcodemain='" & R1!itemcodemain & "' and stockcode='" & R1!itemcode & "'", c, adOpenDynamic, adLockOptimistic
        End With
        
        r!department = r2!department
        r2.Update
        r2.Close
        r.Update
        r.Close
        With r
        .Open "Select * from stock where stockcode='" & R1!itemcode & "' and stockcodemain='" & R1!itemcodemain & "'", c, adOpenDynamic, adLockOptimistic
        End With
        r!qty = r!qty - R1!qty
        r.Update
        r.Close
        R1.MoveNext
    Loop
    R1.Close
    OpenTill
    '////////////////////////////////Printing
    With R1
    .Open "select * from header", c, adOpenDynamic, adLockOptimistic
    End With
    Printer.Font.Bold = True
    Printer.Font.Name = "Arial"
    Printer.Font.Size = 10
    'Printer.RightToLeft = True
    For i = 1 To 10
        Column = "line" & i
        If R1.Fields(Column) <> "BLANK" Then
            Printer.Print R1.Fields(Column)
        Else
            Printer.Print ""
        End If
    Next i
    R1.Close
    
    Printer.Print Now
    Printer.Print "Till" & vbTab & vbTab & ":" & TillId
    Printer.Print "Invoice No." & vbTab & ":" & invnum
    Printer.Print "Cashier" & vbTab & vbTab & ":" & UCase(CurrentUser)
    Printer.Font.Size = 8
    Printer.Print "------------------------------------------------------------------"
    Printer.Print "QTY  StockCode  Serial#       Barcode"
    Printer.Print "UNIT INC  DESCRIP               TOTAL INC"
    Printer.Print "------------------------------------------------------------------"
    Printer.Print ""
    
    With R1
    .Open "sale", HC, adOpenDynamic, adLockOptimistic
    End With
    TotNumItems = 0
    Do While R1.EOF = False
        Printer.Print R1!qty & "  " & R1!itemcodemain & "  " & R1!serialnumber & "  " & R1!itemcode
        TotNumItems = TotNumItems + R1!qty
        If Len(R1!Description) <= 22 Then
            Select Case Len(R1!Description)
                Case 1
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*********************" & vbTab & "R" & R1!total
                Case 2
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "********************" & vbTab & "R" & R1!total
                Case 3
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*******************" & vbTab & "R" & R1!total
                Case 4
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "******************" & vbTab & "R" & R1!total
                Case 5
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*****************" & vbTab & "R" & R1!total
                Case 6
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "****************" & vbTab & "R" & R1!total
                Case 7
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "***************" & vbTab & "R" & R1!total
                Case 8
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "**************" & vbTab & "R" & R1!total
                Case 9
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*************" & vbTab & "R" & R1!total
                Case 10
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "************" & vbTab & "R" & R1!total
                Case 11
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "***********" & vbTab & "R" & R1!total
                Case 12
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "**********" & vbTab & "R" & R1!total
                Case 13
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*********" & vbTab & "R" & R1!total
                Case 14
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "********" & vbTab & "R" & R1!total
                Case 15
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*******" & vbTab & "R" & R1!total
                Case 16
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "******" & vbTab & "R" & R1!total
                Case 17
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*****" & vbTab & "R" & R1!total
                Case 18
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "****" & vbTab & "R" & R1!total
                Case 19
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "***" & vbTab & "R" & R1!total
                Case 20
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "**" & vbTab & "R" & R1!total
                Case 21
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & "*" & vbTab & "R" & R1!total
                Case 22
                    Printer.Print "R" & R1!UNITPRICE & "  " & R1!Description & vbTab & "R" & R1!total
            End Select
        Else
            Printer.Print "R" & R1!UNITPRICE & "  " & Mid(R1!Description, 1, 22) & vbTab & "R" & R1!total
        End If
        Printer.Print "------------------------------------------------------------------"
        Printer.Print ""
    R1.MoveNext
    Loop
    R1.Close
    Printer.Print "TOTAL ITEMS ON INVOICE : ********" & TotNumItems & "********"
    Printer.Print ""
    Printer.Print "------------------------------------------------------------------"
    Printer.Print vbTab & vbTab & "VAT SUMMARY"
    Printer.Print "CODE" & vbTab & "%" & vbTab & "GOODS" & vbTab & "VAT"
    Printer.Print "------------------------------------------------------------------"
    With R1
    .Open "select * from taxcode order by taxcode", c, adOpenDynamic, adLockOptimistic
    End With
    Do While R1.EOF = False
        With r
        .Open "select * from sale where taxcode='" & R1!taxcode & "'", HC, adOpenDynamic, adLockOptimistic
        End With
            tvat = 0
            tgoods = 0
            Do While r.EOF = False
                tgoods = Format(tgoods + CDbl(r!total), "#####.00")
                tvat = Format(tvat + CDbl(r!vat), "#####.00")
                r.MoveNext
            Loop
            r.Close
            If tvat > 0 Then
                Printer.Print R1!taxcode & vbTab & R1!tax & vbTab & Format((tgoods * 100 / (R1!tax + 100)), "#####.00") & vbTab & Format(tvat, "#####.00")
                GTGoods = CDbl(GTGoods) + Format((CDbl(tgoods * 100 / (R1!tax + 100))), "#####.00")
                GTVat = CDbl(GTVat) + CDbl(tvat)
            End If
            
        R1.MoveNext
    Loop
    R1.Close
    Printer.Print "------------------------------------------------------------------"

    Printer.Print "TOTALS" & vbTab & Format(GTGoods, "#####.00") & vbTab & Format(GTVat, "#####.00")
    Printer.Print "------------------------------------------------------------------"
    Printer.Print "------------------------------------------------------------------"
    With R1
    .Open "select * from sale", HC, adOpenDynamic, adLockOptimistic
    End With
    Do While R1.EOF = False
        rTOTDISC = rTOTDISC + CDbl(R1!TOTDISC)
        R1.MoveNext
    Loop
    R1.Close
    Printer.Print "TOTAL DISCOUNT" & vbTab & vbTab & "R" & Format(rTOTDISC, "#####.00")
    Printer.Print "------------------------------------------------------------------"
    Printer.Print vbTab & vbTab & "PAYMENT SUMMARY"
    Printer.Print "AMOUNT DUE" & vbTab & vbTab & vbTab & "R" & lTOTAL; ""
    Printer.Print ""
    Printer.Print "CASH" & vbTab & vbTab & vbTab & vbTab & "R" & Format(txtTEN.Text, "#####.00")
    Printer.Print "CHANGE" & vbTab & vbTab & vbTab & "R" & Format(lchange, "#####.00")
    Printer.Print "------------------------------------------------------------------"
    Printer.Print "------------------------------------------------------------------"
    
    
    With R1
    .Open "select * from footer", c, adOpenDynamic, adLockOptimistic
    End With
    Printer.Font.Bold = True
    
    Printer.Font.Name = "Arial"
    Printer.Font.Size = 8
    For i = 1 To 5
        Column = "line" & i
        If R1.Fields(Column) <> "BLANK" Then
            Printer.Print R1.Fields(Column)
        Else
            Printer.Print ""
        End If
    Next i
    R1.Close
    Printer.EndDoc
    
    With R1
    .Open "delete * from sale", HC, adOpenDynamic, adLockOptimistic
    End With
    frmMain.T1.Text = ""
    frmMain.txtcode.Text = ""
    frmMain.lTOTAL = "0,00"
    frmMain.txtDesc.Caption = "Scan Item"
    MyChange = lchange.Caption
  
    Unload Me
    frmChange.Show vbModal
End If



End Sub

Private Sub txtTEN_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    A = InStr(1, txtTEN, ".")
    If A > 0 Then
        KeyAscii = 0
    End If
End If

If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Or KeyAscii = 44 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If

End Sub
