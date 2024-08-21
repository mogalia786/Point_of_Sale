VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmReturns 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Returns"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   6360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Process Return"
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
      BCOL            =   12632256
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmReturns.frx":0000
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
   Begin VB.Frame Frame5 
      Caption         =   "Customer's Details"
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
      Height          =   1215
      Left            =   120
      TabIndex        =   14
      Top             =   4920
      Width           =   4335
      Begin VB.TextBox txttel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contact number:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Reason for Return"
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
      TabIndex        =   13
      Top             =   3960
      Width           =   4335
      Begin VB.ComboBox cReason 
         Height          =   315
         ItemData        =   "frmReturns.frx":001C
         Left            =   240
         List            =   "frmReturns.frx":001E
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Quantity Returned"
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
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   2175
      Begin VB.TextBox txtqty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Text            =   "1"
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Scan Item Barcode / Serial number"
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
      TabIndex        =   9
      Top             =   1080
      Width           =   4335
      Begin VB.TextBox txtcode 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   4095
      End
      Begin VB.OptionButton optSN 
         Caption         =   "Serial number"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton optBC 
         Caption         =   "Barcode"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "If you are not scanning item but entering item code PLEASE PRESS ENTER TO VERIFY ITEM CODE"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Till Slip Number"
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
      TabIndex        =   8
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtSlip 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   3855
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   6360
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
      BCOL            =   12632256
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmReturns.frx":0020
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
Attribute VB_Name = "frmReturns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UnloadMe As Boolean
Private Sub Form_Load()
Dim r As New Recordset
UnloadMe = False
With r
.Open "select *  from returnreasons order by reason", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    cReason.AddItem r!reason
    r.MoveNext
Loop
r.Close
iTillSlip = ""
iStockcode = ""
iBarcode = ""
iSerial = ""
iQty = ""
iReason = ""
iName = ""
iContact = ""


End Sub

Private Sub LaVolpeButton1_Click()
Dim r As New Recordset
Dim R1 As New Recordset
Dim VATPERITEM As Double
Dim REFUNDAMT As Double
Dim VATREFUND As Double
If Len(iStockcode) = 0 Or Len(iBarcode) = 0 Or Len(iSerial) = 0 Then
    MsgBox "Item code has not been verified! Please re-enter item code!"
    txtcode.Text = ""
    txtcode.SetFocus
End If

If Len(txtSlip.Text) = 0 Then
    MsgBox "Please enter slip number!", vbExclamation, "Status"
    txtSlip.SetFocus
    Exit Sub
End If
If Len(txtcode.Text) = 0 Then
    MsgBox "Please enter item code!", vbExclamation, "Status"
    txtcode.SetFocus
    Exit Sub
End If
If Len(txtqty.Text) = 0 Then
    MsgBox "Please enter quantity!", vbExclamation, "Status"
    txtSqty.SetFocus
    Exit Sub
End If
If Len(cReason.Text) = 0 Then
    MsgBox "Please enter reason!", vbExclamation, "Status"
    cReason.SetFocus
    Exit Sub
End If
If Len(txtname.Text) = 0 Then
    MsgBox "Please enter customer name!", vbExclamation, "Status"
    txtname.SetFocus
    Exit Sub
End If
If Len(txttel.Text) = 0 Then
    MsgBox "Please enter contact number!", vbExclamation, "Status"
    txttel.SetFocus
    Exit Sub
End If

iTillSlip = txtSlip
iQty = txtqty.Text
iReason = cReason.Text

iName = txtname.Text
iContact = txttel.Text
With r
.Open "select * from sales where saleno='" & txtSlip & "' and itemcodemain='" & iStockcode & "' and itemcode='" & iBarcode & "' and serialnumber='" & iSerial & "'", c, adOpenDynamic, adLockOptimistic
End With

With R1
.Open "select * from returns", c, adOpenDynamic, adLockOptimistic
End With
R1.AddNew
R1!returndate = Format(Date, "DD/MM/YYYY")
R1!department = iDepartment
R1!teller = CurrentUser
R1!TillId = TillId
R1!stockcode = iStockcode
R1!barcode = iBarcode
R1!serialnumber = iSerial
R1!qtyreturned = iQty
R1!tillslip = iTillSlip
R1!reason = iReason
R1!customername = iName
R1!contactnumber = iContact
R1!REFUNDAMT = Format(CDbl(iQty) * CDbl(r!UNITPRICE), "#####.00")
REFUNDAMT = Format(CDbl(iQty) * CDbl(r!UNITPRICE), "#####.00")
R1!refundvat = Format(CDbl(r!vat) / CDbl(r!qty) * CDbl(iQty), "#####.00")
VATREFUND = Format(CDbl(r!vat) / CDbl(r!qty) * CDbl(iQty), "#####.00")
R1.Update
R1.Close

VATPERITEM = 0
VATPERITEM = Format(CDbl(r!vat) / CDbl(r!qty), "#####.00")
r!qty = CDbl(r!qty) - CDbl(iQty)
r!total = Format(CDbl(r!total) - REFUNDAMT, "#####.00")
r!vat = Format(CDbl(r!qty) * VATPERITEM, "#####.00")
r.Update
'If CDbl(r!qty) = 0 Then
'    r.Delete
'End If
r.Close
With r
.Open "select * from stock where stockcodemain='" & iStockcode & "' and stockcode='" & iBarcode & "'", c, adOpenDynamic, adLockOptimistic
End With

r!qty = CDbl(r!qty) + CDbl(iQty)
r.Update
r.Close

ReturnResult REFUNDAMT, VATREFUND

Unload Me

End Sub

Private Sub LaVolpeButton2_Click()
Unload Me
End Sub

Private Sub LaVolpeButton2_GotFocus()
'UnloadMe = True
End Sub

Private Sub txtcode_Change()
            iStockcode = ""
            iBarcode = ""
            iSerial = ""
            iDepartment = ""
            
End Sub

Private Sub txtcode_GotFocus()
Dim r As New Recordset
If UnloadMe = False Then
With r
.Open "select * from sales where saleno='" & txtSlip & "'", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = True Then
    MsgBox "Till Slip not found!", vbExclamation, "Not Found!"
    txtSlip.SetFocus
    Exit Sub
End If
End If

End Sub

Private Sub txtcode_KeyUp(KeyCode As Integer, Shift As Integer)
Dim r As New Recordset
If Len(txtcode.Text) = 0 Then Exit Sub

If KeyCode = vbKeyReturn Then
    If Len(txtSlip) = 0 Then
        MsgBox "Please enter till slip number!", vbExclamation, "Status"
        txtSlip.SetFocus
        Exit Sub
    End If

    If optBC Then
        With r
        .Open "select * from sales where saleno='" & txtSlip.Text & "' and serialnumber<>'N/A'", c, adOpenDynamic, adLockOptimistic
        End With
        If r.EOF = False Then
            MsgBox "Please scan serial number only!", vbExclamation, "Status"
            optSN.Value = True
            txtcode.Text = ""
            txtcode.SetFocus
            Exit Sub
        Else
            r.Close
        End If
        
        With r
        .Open "select * from sales where saleno='" & txtSlip & "' and itemcode='" & txtcode & "'", c, adOpenDynamic, adLockOptimistic
        End With
        If r.EOF = True Then
            MsgBox "Item not found on till slip!", vbCritical, "Not FOund!"
            txtcode.Text = ""
            iStockcode = ""
            iBarcode = ""
            iSerial = ""
            iDepartment = ""
            txtcode.SetFocus
            Exit Sub
        Else
            iStockcode = r!itemcodemain
            iBarcode = r!itemcode
            iSerial = r!serialnumber
            iDepartment = r!department
            txtqty.SetFocus
        End If
    Else
      With r
        .Open "select * from sales where saleno='" & txtSlip & "' and serialnumber='" & txtcode & "'", c, adOpenDynamic, adLockOptimistic
        End With
        If r.EOF = True Then
            MsgBox "Item not found on till slip!", vbCritical, "Not FOund!"
            txtcode.Text = ""
            iStockcode = ""
            iBarcode = ""
            iSerial = ""
            iDepartment = ""
            txtcode.SetFocus
            Exit Sub
        Else
            iStockcode = r!itemcodemain
            iBarcode = r!itemcode
            iSerial = r!serialnumber
            iDepartment = r!department
            txtqty.SetFocus
        End If
    End If
 End If
 
End Sub

Public Sub ReturnResult(RA As Double, RV As Double)
Dim r As New Recordset
    Printer.Font.Bold = True
    Printer.Font.Name = "Arial"
    Printer.Font.Size = 10
    'Printer.RightToLeft = True
    Printer.Print vbTab & "REFUND DOCUMENT"
    Printer.Print Now
    Printer.Print "Till" & vbTab & vbTab & ":" & TillId
    Printer.Print "Cashier" & vbTab & vbTab & ":" & UCase(CurrentUser)
    Printer.Font.Size = 8
    Printer.Print "------------------------------------------------------------------"
    Printer.Print "Sale No:" & vbTab & UCase(iTillSlip)
    Printer.Print "Quantity:" & vbTab & iQty
    With r
    .Open "select * from stock where stockcodemain='" & iStockcode & "' and stockcode='" & iBarcode & "'", c, adOpenDynamic, adLockOptimistic
    End With
    If Len(r!STOCKDESC) < 15 Then
    Printer.Print "Description:" & vbTab & r!STOCKDESC
    Else
    Printer.Print "Description:" & vbTab & Mid(r!STOCKDESC, 1, 14)
    End If
    r.Close
    Printer.Print "Stock Code:" & vbTab & iStockcode
    Printer.Print "Barcode:" & vbTab & iBarcode
    Printer.Print "Serial Number:" & vbTab & iSerial
    Printer.Print "Refund Total:" & vbTab & RA
    Printer.Print "VAT Refund:" & vbTab & RV
    Printer.Print "------------------------------------------------------------------"
    Printer.Print ""
    Printer.Print "------------------------------------------------------------------"
    Printer.Print "Customer:" & vbTab & UCase(iName)
    Printer.Print "Tel:" & vbTab & vbTab & iContact
    If Len(iReason) < 41 Then
    Printer.Print "Reason:" & iReason
    Else
    Printer.Print "Reason:" & Mid(iReason, 1, 40)
    End If
    Printer.Print "------------------------------------------------------------------"
    Printer.Print "------------------------------------------------------------------"
    Printer.Print ""
    Printer.Print ""
    Printer.Print "Customer's Signature____________________"
    Printer.Print ""
    Printer.Print "------------------------------------------------------------------"
    Printer.EndDoc
    OpenTill
    

End Sub

Private Sub txtcode_LostFocus()
'If Len(iStockcode) = 0 Or Len(iBarcode) = 0 Or Len(iSerial) = 0 Then
 '   txtcode.SetFocus
'End If

End Sub

Private Sub txtqty_GotFocus()
If Len(iStockcode) = 0 Or Len(iBarcode) = 0 Or Len(iSerial) = 0 Or Len(iDepartment) = 0 Then
    txtcode.SetFocus
End If

End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If

End Sub


Private Sub txtqty_LostFocus()
Dim r As New Recordset
With r
.Open "select * from sales where itemcodemain='" & iStockcode & "' and itemcode='" & iBarcode & "' and serialnumber='" & iSerial & "'", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = True Then
    MsgBox "Code not verified on till slip!", vbExclamation, "Status"
    txtcode.SetFocus
    Exit Sub
End If

If r!qty = "0" Then
    MsgBox "This item has already been returned!", vbExclamation, "Status"
    txtqty.Text = "1"
    txtcode.Text = ""
    txtSlip.Text = ""
    txtSlip.SetFocus
    Exit Sub
End If

If CDbl(r!qty) < CDbl(txtqty) Then
    MsgBox "Returned quantity cannot exceed quantity purchased!", vbExclamation, "Status"
    txtqty.Text = ""
    txtqty.SetFocus
End If
r.Close

End Sub


