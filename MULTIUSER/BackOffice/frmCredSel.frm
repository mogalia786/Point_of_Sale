VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCredSel 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Creditor Account"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
   ControlBox      =   0   'False
   Icon            =   "frmCredSel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&View Details"
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
      MICON           =   "frmCredSel.frx":0442
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
      Caption         =   "Select Creditor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4335
      Begin MSComCtl2.DTPicker Fromdate 
         Height          =   255
         Left            =   2520
         TabIndex        =   1
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   55246849
         CurrentDate     =   38337
      End
      Begin VB.ComboBox Cred 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Generate Account From:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Creditor:"
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
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1440
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
      MICON           =   "frmCredSel.frx":045E
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
Attribute VB_Name = "frmCredSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim r As New Recordset

With r
.Open "select * from supplier order by supplier", c, adOpenDynamic, adLockOptimistic
End With

Do While r.EOF = False
    Cred.AddItem r!supplier
    r.MoveNext
Loop
r.Close


End Sub

Private Sub LaVolpeButton1_Click()
Dim pbody As String
Dim r As New Recordset
Dim r1 As New Recordset
Dim wd As New Word.Application
Dim xHour As String
Dim xMin As String
Dim xSec As String
Dim OpeningBalance As Double
Dim DebitBalance As Double
Dim CreditBalance As Double
If Len(Cred.Text) > 0 Then
    CredFromDate = Format(Fromdate.Value, "DD/MM/YYYY")
    creditor = Cred.Text
Else
Exit Sub
End If

With r
.Open "select * from supplier where supplier='" & creditor & "'", c, adOpenDynamic, adLockOptimistic
End With
OpeningBalance = r!OpeningBalance
r.Close

With r
.Open "select * from creditorsinvoice where creditor='" & creditor & "'and invdate<'" & Year(Fromdate) & "-" & Month(Fromdate) & "-" & Day(Fromdate) & " 00:00:00.000'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
        OpeningBalance = OpeningBalance + CDbl(r!tendered)
    r.MoveNext
Loop
r.Close

With r
.Open "select * from creditorspayment where creditor='" & creditor & "'and paymentdate<'" & Year(Fromdate) & "-" & Month(Fromdate) & "-" & Day(Fromdate) & " 00:00:00.000'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
        OpeningBalance = OpeningBalance - CDbl(r!amount)
    r.MoveNext
Loop
r.Close

With r
.Open "select * from creditorscreditnote where creditor='" & creditor & "'and notedate<'" & Year(Fromdate) & "-" & Month(Fromdate) & "-" & Day(Fromdate) & " 00:00:00.000'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
        OpeningBalance = OpeningBalance - CDbl(r!amount)
    r.MoveNext
Loop
r.Close


xHour = Hour(Time)
xMin = Minute(Time)
xSec = Second(Time)

If Dir("\\" & servername & "\" & sharename & "\Creditors", vbDirectory) = "" Then
MkDir "\\" & servername & "\" & sharename & "\Creditors"
End If

Open "\\" & servername & "\" & sharename & "\Creditors\" & Day(Date) & "#" & MonthName(Month(Date)) & "#" & Year(Date) & xHour & xMin & xSec & ".doc" For Output As #2




With r
.Open "select * from supplier where supplier='" & creditor & "'", c, adOpenDynamic, adLockOptimistic
End With
pbody = pbody + "<html>"
pbody = pbody + "<head>"
pbody = pbody + "<title>Untitled Document</title>"
pbody = pbody + "<meta http-equiv=Content-Type content=text/html; charset=iso-8859-1>"
pbody = pbody + "</head>"

pbody = pbody + "<body bgcolor=#FFFFFF text=#000000>"
pbody = pbody + "<div align=center>"
pbody = pbody + "  <p><b><font size=4>Transaction List</font></b></p>"
pbody = pbody + "  <p><b><font size=4>Creditors Account from " & Fromdate.Value & "</font></b></p>"

pbody = pbody + "  <table width=100% border=1 cellspacing=1 cellpadding=0>"


pbody = pbody + "      <td width=18% bgcolor=#000000> "
pbody = pbody + "        <div align=left><font color=#FFFFFF><b><font size=2>Account / Creditor "
pbody = pbody + "          Name:</font></b></font></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=82%> "
pbody = pbody + "        <div align=left><b><font color=#000000><font "
pbody = pbody + "size=2>" & UCase(r!supplier) & "</font></font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "    </tr>"

pbody = pbody + "    <tr> "
pbody = pbody + "      <td width=18% bgcolor=#000000> "
pbody = pbody + "        <div align=left><font color=#FFFFFF><b><font "
pbody = pbody + "size=2>Address1:</font></b></font></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=82%> "
pbody = pbody + "        <div align=left><b><font color=#000000><font "
pbody = pbody + "size=2>" & r!add1 & "</font></font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "    </tr>"

pbody = pbody + "    <tr> "
pbody = pbody + "      <td width=18% bgcolor=#000000> "
pbody = pbody + "        <div align=left><font color=#FFFFFF><b><font "
pbody = pbody + "size=2>Address2:</font></b></font></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=82%> "
pbody = pbody + "        <div align=left><b><font color=#000000><font "
pbody = pbody + "size=2>" & r!add2 & "</font></font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "    </tr>"

pbody = pbody + "    <tr> "
pbody = pbody + "      <td width=18% bgcolor=#000000> "
pbody = pbody + "        <div align=left><font color=#FFFFFF><b><font "
pbody = pbody + "size=2>Address3:</font></b></font></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=82%> "
pbody = pbody + "        <div align=left><b><font color=#000000><font "
pbody = pbody + "size=2>" & r!add3 & "</font></font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "    </tr>"

pbody = pbody + "    <tr> "
pbody = pbody + "      <td width=18% bgcolor=#000000> "
pbody = pbody + "        <div align=left><font color=#FFFFFF><b><font "
pbody = pbody + "size=2>Address4:</font></b></font></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=82%> "
pbody = pbody + "        <div align=left><b><font color=#000000><font "
pbody = pbody + "size=2>" & r!add4 & "</font></font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "    </tr>"





pbody = pbody + "  </table>"
r.Close
pbody = pbody + "  <hr>"

pbody = pbody + "  <hr>"

pbody = pbody + "  <table width=100% border=1 cellspacing=1 cellpadding=0>"
pbody = pbody + "    <tr bgcolor=#000000> "

pbody = pbody + "      <td width=4%> "
pbody = pbody + "        <div align=center><font color=#FFFFFF><b><font "
pbody = pbody + "size=2>Date</font></b></font></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=20%> "
pbody = pbody + "        <div align=center><b><font size=2 color=#FFFFFF>Transaction</font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=11%> "
pbody = pbody + "        <div align=center><b><font size=2><font color=#000000><font "
pbody = pbody + "color=#FFFFFF>Inv "
pbody = pbody + "          /Chq/CN No</font></font></font></b></div>"
pbody = pbody + "      </td>"


pbody = pbody + "      <td width=5%> "
pbody = pbody + "        <div align=center><b><font size=2><font color=#000000><font "
pbody = pbody + "color=#FFFFFF>Ref</font></font></font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=7%> "
pbody = pbody + "        <div align=center><b><font size=2><font color=#000000><font "
pbody = pbody + "color=#FFFFFF>Payment "
pbody = pbody + "          Due </font></font></font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=8%> "
pbody = pbody + "        <div align=center><b><font size=2><font color=#000000><font "
pbody = pbody + "color=#FFFFFF>Debit</font></font></font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=16%> "
pbody = pbody + "        <div align=center><b><font size=2><font color=#000000><font "
pbody = pbody + "color=#FFFFFF>Credit</font></font></font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "    </tr>"

'//////////////////////////////////////////////////START WITH ACCOUNT OPENING BALANCE////////////////
pbody = pbody + "    <tr> "
pbody = pbody + "      <td width=4%> "
pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif></font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=20%> "
pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>Opening Balance</font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=11%> "
pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif></font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=5%> "
pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif></font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=7%> "
pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif></font></b></div>"
pbody = pbody + "      </td>"
If OpeningBalance > 0 Then
pbody = pbody + "      <td width=8%> "
pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>&nbsp</font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=16%> "
pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>R" & OpeningBalance & "</font></b></div>"
pbody = pbody + "      </td>"
CreditBalance = CreditBalance + CDbl(OpeningBalance)
Else
pbody = pbody + "      <td width=8%> "
pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>R" & OpeningBalance & "</font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=16%> "
pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>&nbsp</font></b></div>"
pbody = pbody + "      </td>"
DebitBalance = DebitBalance + CDbl(OpeningBalance)
End If

pbody = pbody + "    </tr>"

'/////////////////CLOSE OPENING BALANCE///////////////////

With r
.Open "select * from creditorsinvoice where creditor='" & creditor & "'and invdate>='" & Year(Fromdate) & "-" & Month(Fromdate) & "-" & Day(Fromdate) & " 00:00:00.000'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    With r1
    .Open "select * from credtemp", c, adOpenDynamic, adLockOptimistic
    End With
    r1.AddNew
    r1!ddate = r!invdate
    r1!transtype = "Invoice"
    r1!docno = r!invno
    r1!paymentdate = r!paymentdate
    r1!contra = "n/a"
    r1!amount = Format(r!tendered, "#####.00")
    r1.Update
    r1.Close
    r.MoveNext
Loop
r.Close


With r
.Open "select * from creditorspayment where creditor='" & creditor & "'and paymentdate>='" & Year(Fromdate) & "-" & Month(Fromdate) & "-" & Day(Fromdate) & " 00:00:00.000'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    With r1
    .Open "select * from credtemp", c, adOpenDynamic, adLockOptimistic
    End With
    r1.AddNew
    r1!ddate = r!paymentdate
    r1!transtype = "Payment"
    r1!docno = r!chqno
    r1!paymentdate = "01/01/1900"
    r1!contra = r!invno
    r1!amount = Format(r!amount, "#####.00")
    r1.Update
    r1.Close
    r.MoveNext
Loop
r.Close

With r
.Open "select * from creditorscreditnote where creditor='" & creditor & "'and notedate>='" & Year(Fromdate) & "-" & Month(Fromdate) & "-" & Day(Fromdate) & " 00:00:00.000'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    With r1
    .Open "select * from credtemp", c, adOpenDynamic, adLockOptimistic
    End With
    r1.AddNew
    r1!ddate = r!notedate
    r1!transtype = "Credit Note"
    r1!docno = r!noteno
    r1!paymentdate = "01/01/1900"
    r1!contra = r!invno
    r1!amount = Format(r!amount, "#####.00")
    r1.Update
    r1.Close
    r.MoveNext
Loop
r.Close



With r
.Open "select * from credtemp order by ddate", c, adOpenDynamic, adLockOptimistic
End With

Do While r.EOF = False
'///////////////START ACCOUNT LINES/////////////////////
    pbody = pbody + "    <tr> "
    pbody = pbody + "      <td width=4%> "
    pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>" & r!ddate & "</font></b></div>"
    pbody = pbody + "      </td>"
        pbody = pbody + "      <td width=20%> "
        pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>" & r!transtype & "</font></b></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=11%> "
        pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>" & r!docno & "</font></b></div>"
        pbody = pbody + "      </td>"
        
    
    pbody = pbody + "      <td width=5%> "
    pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>" & r!contra & "</font></b></div>"
    pbody = pbody + "      </td>"
        If r!paymentdate <> "01/01/1900" Then
            pbody = pbody + "      <td width=7%> "
            pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>" & r!paymentdate & "</font></b></div>"
            pbody = pbody + "      </td>"
        Else
            pbody = pbody + "      <td width=7%> "
            pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>n/a</font></b></div>"
            pbody = pbody + "      </td>"
        End If
        
    If r!transtype = "Invoice" Then
        pbody = pbody + "      <td width=8%> "
        pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>&nbsp</font></b></div>"
        pbody = pbody + "      </td>"
        pbody = pbody + "      <td width=16%> "
        pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>R" & r!amount & "</font></b></div>"
        pbody = pbody + "      </td>"
        CreditBalance = CreditBalance + CDbl(r!amount)
    ElseIf r!transtype = "Payment" Then
        pbody = pbody + "      <td width=8%> "
        pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>R" & r!amount & "</font></b></div>"
        pbody = pbody + "      </td>"
        pbody = pbody + "      <td width=16%> "
        pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>&nbsp</font></b></div>"
        pbody = pbody + "      </td>"
        DebitBalance = DebitBalance + CDbl(r!amount)
    ElseIf r!transtype = "Credit Note" Then
        pbody = pbody + "      <td width=8%> "
        pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>R" & r!amount & "</font></b></div>"
        pbody = pbody + "      </td>"
        pbody = pbody + "      <td width=16%> "
        pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>&nbsp</font></b></div>"
        pbody = pbody + "      </td>"
        DebitBalance = DebitBalance + CDbl(r!amount)

    End If
    pbody = pbody + "    </tr>"
    r.MoveNext
Loop
r.Close

'////////////////CLOSE ACCOUNT LINES///////////////////////////////////

'//////////////////TOTAL//////////////////////////////////////////////
pbody = pbody + "    <tr> "
pbody = pbody + "      <td width=4%>&nbsp;</td>"
pbody = pbody + "      <td width=20%>&nbsp;</td>"
pbody = pbody + "      <td width=11%>&nbsp;</td>"
pbody = pbody + "      <td width=5%>&nbsp;</td>"
pbody = pbody + "      <td width=7% bgcolor=#000000> "
pbody = pbody + "        <div align=center><b><font size=2 color=#FFFFFF>Total "
pbody = pbody + "Balance</font></b></div>"
pbody = pbody + "      </td>"
If CreditBalance > DebitBalance Then
    pbody = pbody + "      <td width=8%> "
    pbody = pbody + "        <div align=left><b><b><font size=2>&nbsp</font></b></b></div>"
    pbody = pbody + "      </td>"
    pbody = pbody + "      <td width=16%> "
    pbody = pbody + "        <div align=left><b><b><font size=2>R" & Format(CreditBalance - DebitBalance, "####.00") & "</font></b></b></div>"
    pbody = pbody + "      </td>"
Else
    pbody = pbody + "      <td width=8%> "
    pbody = pbody + "        <div align=left><b><b><font size=2>R" & Format(DebitBalance - CreditBalance, "####.00") & "</font></b></b></div>"
    pbody = pbody + "      </td>"
    pbody = pbody + "      <td width=16%> "
    pbody = pbody + "        <div align=left><b><b><font size=2>&nbsp</font></b></b></div>"
    pbody = pbody + "      </td>"
End If

pbody = pbody + "    </tr>"
'///////////////////////////////////////////
pbody = pbody + "  </table>"
pbody = pbody + "  <p align=left>&nbsp;</p>"
pbody = pbody + "</div>"
pbody = pbody + "</body>"
pbody = pbody + "</html>"
Print #2, pbody
Close #2
With r1
.Open "delete from credtemp", c, adOpenDynamic, adLockOptimistic
End With

    wd.Documents.Open "\\" & servername & "\" & sharename & "\Creditors\" & Day(Date) & "#" & MonthName(Month(Date)) & "#" & Year(Date) & xHour & xMin & xSec & ".doc"
    wd.Visible = True
RecordAction CurrentUser, Date, Time, "Successful..", "Generated Account for " & creditor & "..."


End Sub

Private Sub LaVolpeButton2_Click()
Unload Me

End Sub

Private Sub LaVolpeButton3_Click()
Dim pbody As String
Dim r As New Recordset
Dim wd As New Word.Application
Dim xHour As String
Dim xMin As String
Dim xSec As String
Dim OpeningBalance As Double
Dim DebitBalance As Double
Dim CreditBalance As Double
On Error GoTo nomail
If Len(txtemail.Text) = 0 Then
    MsgBox "Please enter email address!", vbExclamation, "E Mail"
    txtemail.SetFocus
    Exit Sub
End If
If Len(Cred.Text) > 0 Then
    CredFromDate = Format(Fromdate.Value, "DD/MM/YYYY")
    creditor = Cred.Text
Else
Exit Sub
End If

With r
.Open "select * from exporterdetails where exporter='" & creditor & "'", c, adOpenDynamic, adLockOptimistic
End With
OpeningBalance = r!OpeningBalance
r.Close
With r
.Open "select * from creditors where exporter='" & creditor & "'and transdate<'" & Year(Fromdate) & "-" & Month(Fromdate) & "-" & Day(Fromdate) & " 00:00:00.000' and hasexc='Yes' order by transdate", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    If r!transactiontype = "Inv" Then
        OpeningBalance = OpeningBalance + r!rands
    Else
        OpeningBalance = OpeningBalance - r!rands
    End If
    r.MoveNext
Loop
r.Close

xHour = Hour(Time)
xMin = Minute(Time)
xSec = Second(Time)

If Dir("\\" & servername & "\" & sharename & "\Creditors", vbDirectory) = "" Then
MkDir "\\" & servername & "\" & sharename & "\Creditors"
End If

Open "\\" & servername & "\" & sharename & "\Creditors\" & Day(Date) & "#" & MonthName(Month(Date)) & "#" & Year(Date) & xHour & xMin & xSec & ".doc" For Output As #2

With r
.Open "select * from exporterdetails where exporter='" & creditor & "'", c, adOpenDynamic, adLockOptimistic
End With
pbody = pbody + "<html>"
pbody = pbody + "<head>"
pbody = pbody + "<title>Untitled Document</title>"
pbody = pbody + "<meta http-equiv=Content-Type content=text/html; charset=iso-8859-1>"
pbody = pbody + "</head>"

pbody = pbody + "<body bgcolor=#FFFFFF text=#000000>"
pbody = pbody + "<div align=center>"
pbody = pbody + "  <p><b><font size=4>Malls Tiles</font></b></p>"
pbody = pbody + "  <p><b><font size=4>Creditors Account from " & Fromdate.Value & "</font></b></p>"

pbody = pbody + "  <table width=100% border=1 cellspacing=1 cellpadding=0>"

pbody = pbody + "    <tr> "
pbody = pbody + "      <td width=18% bgcolor=#000000> "
pbody = pbody + "        <div align=left><font color=#FFFFFF><b><font size=2>Account "
pbody = pbody + "Number:</font></b></font></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=82%> "
pbody = pbody + "        <div align=left><b><font color=#000000><font "
pbody = pbody + "size=2>" & r!ExpId & "</font></font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "    </tr>"
pbody = pbody + "    <tr> "

pbody = pbody + "      <td width=18% bgcolor=#000000> "
pbody = pbody + "        <div align=left><font color=#FFFFFF><b><font size=2>Account / Creditor "
pbody = pbody + "          Name:</font></b></font></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=82%> "
pbody = pbody + "        <div align=left><b><font color=#000000><font "
pbody = pbody + "size=2>" & UCase(r!exporter) & "</font></font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "    </tr>"

pbody = pbody + "    <tr> "
pbody = pbody + "      <td width=18% bgcolor=#000000> "
pbody = pbody + "        <div align=left><font color=#FFFFFF><b><font "
pbody = pbody + "size=2>Address:</font></b></font></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=82%> "
pbody = pbody + "        <div align=left><b><font color=#000000><font "
pbody = pbody + "size=2>" & r!Address & "</font></font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "    </tr>"

pbody = pbody + "  </table>"
r.Close
pbody = pbody + "  <hr>"

pbody = pbody + "  <hr>"

pbody = pbody + "  <table width=100% border=1 cellspacing=1 cellpadding=0>"
pbody = pbody + "    <tr bgcolor=#000000> "

pbody = pbody + "      <td width=4%> "
pbody = pbody + "        <div align=center><font color=#FFFFFF><b><font "
pbody = pbody + "size=2>Date</font></b></font></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=20%> "
pbody = pbody + "        <div align=center><b><font size=2 color=#FFFFFF>Transaction</font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=11%> "
pbody = pbody + "        <div align=center><b><font size=2><font color=#000000><font "
pbody = pbody + "color=#FFFFFF>Inv "
pbody = pbody + "          /Chq/CN No</font></font></font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=16%> "
pbody = pbody + "        <div align=center><b><font size=2><font color=#000000><font "
pbody = pbody + "color=#FFFFFF>File "
pbody = pbody + "          Number </font></font></font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=8%> "
pbody = pbody + "        <div align=center><b><font size=2><font color=#000000><font "
pbody = pbody + "color=#FFFFFF>Foreign "
pbody = pbody + "          Amt </font></font></font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=5%> "
pbody = pbody + "        <div align=center><b><font size=2><font color=#000000><font "
pbody = pbody + "color=#FFFFFF>Exch. "
pbody = pbody + "          Rate </font></font></font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=5%> "
pbody = pbody + "        <div align=center><b><font size=2><font color=#000000><font "
pbody = pbody + "color=#FFFFFF>Terms</font></font></font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=7%> "
pbody = pbody + "        <div align=center><b><font size=2><font color=#000000><font "
pbody = pbody + "color=#FFFFFF>Payment "
pbody = pbody + "          Due </font></font></font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=8%> "
pbody = pbody + "        <div align=center><b><font size=2><font color=#000000><font "
pbody = pbody + "color=#FFFFFF>Debit</font></font></font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=16%> "
pbody = pbody + "        <div align=center><b><font size=2><font color=#000000><font "
pbody = pbody + "color=#FFFFFF>Credit</font></font></font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "    </tr>"

'//////////////////////////////////////////////////START WITH ACCOUNT OPENING BALANCE////////////////
pbody = pbody + "    <tr> "
pbody = pbody + "      <td width=4%> "
pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif></font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=20%> "
pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>Opening Balance</font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=11%> "
pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif></font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=16%> "
pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif></font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=8%> "
pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif></font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=5%> "
pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif></font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=5%> "
pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif></font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=7%> "
pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif></font></b></div>"
pbody = pbody + "      </td>"
If OpeningBalance > 0 Then
pbody = pbody + "      <td width=8%> "
pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>&nbsp</font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=16%> "
pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>R" & OpeningBalance & "</font></b></div>"
pbody = pbody + "      </td>"
CreditBalance = CreditBalance + CDbl(OpeningBalance)
Else
pbody = pbody + "      <td width=8%> "
pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>R" & OpeningBalance & "</font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=16%> "
pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>&nbsp</font></b></div>"
pbody = pbody + "      </td>"
DebitBalance = DebitBalance + CDbl(OpeningBalance)
End If

pbody = pbody + "    </tr>"

'/////////////////CLOSE OPENING BALANCE///////////////////
With r
.Open "select * from creditors where exporter='" & creditor & "'and  transdate>='" & Year(Fromdate) & "-" & Month(Fromdate) & "-" & Day(Fromdate) & " 00:00:00.000' and hasexc='Yes' order by transdate", c, adOpenDynamic, adLockOptimistic
End With

Do While r.EOF = False
'///////////////START ACCOUNT LINES/////////////////////
    pbody = pbody + "    <tr> "
    pbody = pbody + "      <td width=4%> "
    pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>" & r!transdate & "</font></b></div>"
    pbody = pbody + "      </td>"
    If r!transactiontype = "Inv" Then
        pbody = pbody + "      <td width=20%> "
        pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>Invoice</font></b></div>"
        pbody = pbody + "      </td>"
    ElseIf r!transactiontype = "Chq" Then
        pbody = pbody + "      <td width=20%> "
        pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>Payment</font></b></div>"
        pbody = pbody + "      </td>"
    ElseIf r!transactiontype = "CN" Then
        pbody = pbody + "      <td width=20%> "
        pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>Credit Note</font></b></div>"
        pbody = pbody + "      </td>"
    End If
    If r!transactiontype = "Inv" Then
        pbody = pbody + "      <td width=11%> "
        pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>" & r!invno & "</font></b></div>"
        pbody = pbody + "      </td>"
    Else
        pbody = pbody + "      <td width=11%> "
        pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>" & r!chqno & "</font></b></div>"
        pbody = pbody + "      </td>"
    End If
    If r!transactiontype = "Inv" Then
        pbody = pbody + "      <td width=16%> "
        pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>" & r!FileNumber & "</font></b></div>"
        pbody = pbody + "      </td>"
    Else
        pbody = pbody + "      <td width=16%> "
        pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>" & r!FileNumber & " - " & r!invno & "</font></b></div>"
        pbody = pbody + "      </td>"
    End If
    pbody = pbody + "      <td width=8%> "
    pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>" & r!fa & "</font></b></div>"
    pbody = pbody + "      </td>"
    pbody = pbody + "      <td width=5%> "
    pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>" & r!exc & "</font></b></div>"
    pbody = pbody + "      </td>"
    pbody = pbody + "      <td width=5%> "
    pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>" & r!pay & "</font></b></div>"
    pbody = pbody + "      </td>"
    If r!transactiontype = "Inv" Then
        pbody = pbody + "      <td width=7%> "
        pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>" & r!paymentdate & "</font></b></div>"
        pbody = pbody + "      </td>"
    Else
        pbody = pbody + "      <td width=7%> "
        pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>&nbsp</font></b></div>"
        pbody = pbody + "      </td>"
    End If
    If r!transactiontype = "Inv" Then
        pbody = pbody + "      <td width=8%> "
        pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>&nbsp</font></b></div>"
        pbody = pbody + "      </td>"
        pbody = pbody + "      <td width=16%> "
        pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>R" & r!rands & "</font></b></div>"
        pbody = pbody + "      </td>"
        CreditBalance = CreditBalance + CDbl(r!rands)
    Else
        pbody = pbody + "      <td width=8%> "
        pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>R" & r!rands & "</font></b></div>"
        pbody = pbody + "      </td>"
        pbody = pbody + "      <td width=16%> "
        pbody = pbody + "        <div align=left><b><font size=1 face=Arial, Helvetica, sans-serif>&nbsp</font></b></div>"
        pbody = pbody + "      </td>"
        DebitBalance = DebitBalance + CDbl(r!rands)
    End If
    pbody = pbody + "    </tr>"
    r.MoveNext
Loop
r.Close

'////////////////CLOSE ACCOUNT LINES///////////////////////////////////

'//////////////////TOTAL//////////////////////////////////////////////
pbody = pbody + "    <tr> "
pbody = pbody + "      <td width=4%>&nbsp;</td>"
pbody = pbody + "      <td width=20%>&nbsp;</td>"
pbody = pbody + "      <td width=11%>&nbsp;</td>"
pbody = pbody + "      <td width=16%>&nbsp;</td>"
pbody = pbody + "      <td width=8%>&nbsp;</td>"
pbody = pbody + "      <td width=5%>&nbsp;</td>"
pbody = pbody + "      <td width=5%>&nbsp;</td>"
pbody = pbody + "      <td width=7% bgcolor=#000000> "
pbody = pbody + "        <div align=center><b><font size=2 color=#FFFFFF>Total "
pbody = pbody + "Balance</font></b></div>"
pbody = pbody + "      </td>"
If CreditBalance > DebitBalance Then
    pbody = pbody + "      <td width=8%> "
    pbody = pbody + "        <div align=left><b><b><font size=2>&nbsp</font></b></b></div>"
    pbody = pbody + "      </td>"
    pbody = pbody + "      <td width=16%> "
    pbody = pbody + "        <div align=left><b><b><font size=2>R" & Format(CreditBalance - DebitBalance, "####.00") & "</font></b></b></div>"
    pbody = pbody + "      </td>"
Else
    pbody = pbody + "      <td width=8%> "
    pbody = pbody + "        <div align=left><b><b><font size=2>R" & Format(DebitBalance - CreditBalance, "####.00") & "</font></b></b></div>"
    pbody = pbody + "      </td>"
    pbody = pbody + "      <td width=16%> "
    pbody = pbody + "        <div align=left><b><b><font size=2>&nbsp</font></b></b></div>"
    pbody = pbody + "      </td>"
End If

pbody = pbody + "    </tr>"
'///////////////////////////////////////////
pbody = pbody + "  </table>"
pbody = pbody + "  <p align=left>&nbsp;</p>"
pbody = pbody + "</div>"
pbody = pbody + "</body>"
pbody = pbody + "</html>"
Print #2, pbody
Close #2
    
    MAPISession1.SignOn 'sign on
    If MAPISession1.SessionID <> 0 Then 'signed on
        With MAPIMessages1
        .SessionID = MAPISession1.SessionID
        .Compose
        .AttachmentName = "Creditors Account" 'attachment name
        .AttachmentPathName = "\\" & servername & "\" & sharename & "\Creditors\" & Day(Date) & "#" & MonthName(Month(Date)) & "#" & Year(Date) & xHour & xMin & xSec & ".doc"
        .RecipAddress = txtemail.Text 'set the receiver's email To the one they specified (again, text box or a default address)
        .MsgSubject = "Creditors Account" 'set the subject
        .MsgNoteText = "Regards. Please find Creditors Account attached." 'message text
        .Send True 'don't display a dialog saying it was sent
        End With
    End If
Exit Sub
nomail:
MsgBox "E Mail could not be sent - " & Err.Description & ". Please try again later!", vbExclamation, "E mail Error!"

End Sub


