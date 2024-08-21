VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRepSalesSummary 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Summary Report"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&View Report"
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
      MICON           =   "frmRepSalesSummary.frx":0000
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
      Caption         =   "Date Criteria"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox c1 
         Caption         =   "All Dates"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dFROM 
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   19791873
         CurrentDate     =   38483
      End
      Begin MSComCtl2.DTPicker dTO 
         Height          =   255
         Left            =   2400
         TabIndex        =   3
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   19791873
         CurrentDate     =   38483
      End
      Begin VB.Label Label2 
         Caption         =   "To:"
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "From:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   615
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   1560
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
      MICON           =   "frmRepSalesSummary.frx":001C
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
Attribute VB_Name = "frmRepSalesSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub c1_Click()
If c1.Value Then
    dFROM.Enabled = False

    dTO.Enabled = False
Else
    dFROM.Enabled = True
    dTO.Enabled = True
End If

End Sub

Private Sub Form_Load()
dFROM.Value = Format(Date, "DD/MM/YYYY")
dTO.Value = Format(Date, "DD/MM/YYYY")
End Sub

Private Sub LaVolpeButton1_Click()
Dim pbody As String
Dim r As New Recordset
Dim r1 As New Recordset
Dim r2 As New Recordset
Dim r3 As New Recordset
Dim r7 As New Recordset
Dim wd As New Word.Application
Dim xHour As String
Dim xMin As String
Dim xSec As String
Dim CurrentDate As Date
Dim tQTY As Double
Dim TDisc As Double
Dim TVAT As Double
Dim TTotal As Double
Dim sPRICE As Double
Dim GTQTY As Double
Dim GTDISC As Double
Dim GTVAT As Double
Dim GTTOTAL As Double

xHour = Hour(Time)
xMin = Minute(Time)
xSec = Second(Time)

If Dir(App.Path & "\Reports", vbDirectory) = "" Then
MkDir App.Path & "\Reports"
End If

Open App.Path & "\Reports\" & Day(Date) & "#" & MonthName(Month(Date)) & "#" & Year(Date) & xHour & xMin & xSec & ".doc" For Output As #2




pbody = pbody + "<html>"
pbody = pbody + "<head>"
pbody = pbody + "<title>Product Sales Report</title>"
pbody = pbody + "<meta http-equiv=Content-Type content=text/html; charset=iso-8859-1>"
pbody = pbody + "</head>"

pbody = pbody + "<body bgcolor=#FFFFFF text=#000000>"
pbody = pbody + "<div align=center>"
pbody = pbody + "  <p><b><font face=Times New Roman, Times, serif><u>" & CompName & "</u></font></b></p>"

pbody = pbody + "  <p><b><font face=Times New Roman, Times, serif><u>PRODUCT SALES "

pbody = pbody + "SUMMARY</u></font></b></p>"


pbody = pbody + "  <p align=left><b><font face=Times New Roman, Times, serif size=2>Date "
pbody = pbody + "    : " & Format(Date, "DD/MM/YYYY") & "</font></b></p>"
pbody = pbody + "  <p align=left><b><font face=Times New Roman, Times, serif size=2>Time "
pbody = pbody + "    : " & Time & "</font></b></p>"
If c1.Value = False Then
    pbody = pbody + "  <p align=left><b><font face=Times New Roman, Times, serif size=2>SELECTED "
    pbody = pbody + "    DATES : " & Format(dFROM.Value, "DD/MM/YYYY") & " - " & Format(dTO.Value, "DD/MM/YYYY") & "</font></b></p>"
Else
    pbody = pbody + "  <p align=left><b><font face=Times New Roman, Times, serif size=2>SELECTED "
    pbody = pbody + "    DATES : ALL</font></b></p>"
End If
With r7
.Open "select * from department order by department", c, adOpenDynamic, adLockOptimistic
End With
Do While r7.EOF = False
GTTOTAL = 0
GTVAT = 0
GTDISC = 0
GTQTY = 0

    pbody = pbody + "  <p align=left><b><font face=Times New Roman, Times, serif size=2>"
    pbody = pbody + "Department : " & r7!department & "</font></b></p>"


pbody = pbody + "  <table width=100% border=1 cellspacing=1 cellpadding=0>"
pbody = pbody + "    <tr bgcolor=#000000> "
pbody = pbody + "      <td width=11%> "
pbody = pbody + "        <div align=left><b><font color=#FFFFFF>Stock Code</font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=32%> "
pbody = pbody + "        <div align=left><b><b><font size=2><font "

pbody = pbody + "color=#FFFFFF>Description</font></font></b></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=6%> "
pbody = pbody + "        <div align=center><b><b><font size=2><font "

pbody = pbody + "color=#FFFFFF>Qty</font></font></b></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=8%> "
pbody = pbody + "        <div align=center><b><b><font size=2><font "

pbody = pbody + "color=#FFFFFF>Price</font></font></b></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=11%> "
pbody = pbody + "        <div align=center><b><b><font size=2><font "

pbody = pbody + "color=#FFFFFF>Discount</font></font></b></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=7%> "
pbody = pbody + "        <div align=center><b><b><font size=2><font "

pbody = pbody + "color=#FFFFFF>VAT</font></font></b></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=7%> "
pbody = pbody + "        <div align=center><b><b><font size=2><font "

pbody = pbody + "color=#FFFFFF>Total</font></font></b></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=18%> "
pbody = pbody + "        <div align=center><b><b><font size=2><font "

pbody = pbody + "color=#FFFFFF>Onhand</font></font></b></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "    </tr>"
If c1.Value = 1 Then
    With r
    .Open "select distinct itemcodemain from sales where department='" & r7!department & "' order by itemcodemain", c, adOpenDynamic, adLockOptimistic
    End With
Else
    With r
    .Open "select distinct(itemcodemain) from sales where saledate>='" & Year(dFROM.Value) & "-" & Month(dFROM.Value) & "-" & Day(dFROM.Value) & " 00:00:00.000' and saledate<='" & Year(dTO.Value) & "-" & Month(dTO.Value) & "-" & Day(dTO.Value) & " 00:00:00.000' and department='" & r7!department & "' order by itemcodemain", c, adOpenDynamic, adLockOptimistic
    End With
End If
Do While r.EOF = False
    If c1.Value = 1 Then
        With r3
        .Open "select distinct(unitprice) from sales where itemcodemain='" & r!itemcodemain & "' and department='" & r7!department & "' order by unitprice", c, adOpenDynamic, adLockOptimistic
        End With
    Else
        With r3
        .Open "select distinct(unitprice) from sales where saledate>='" & Year(dFROM.Value) & "-" & Month(dFROM.Value) & "-" & Day(dFROM.Value) & " 00:00:00.000' and saledate<='" & Year(dTO.Value) & "-" & Month(dTO.Value) & "-" & Day(dTO.Value) & " 00:00:00.000' and itemcodemain='" & r!itemcodemain & "' and department='" & r7!department & "' order by unitprice", c, adOpenDynamic, adLockOptimistic
        End With
    End If
    Do While r3.EOF = False
    
        If c1.Value = 1 Then
            With r1
            .Open "select * from sales where itemcodemain='" & r!itemcodemain & "' and unitprice='" & r3!unitprice & "' and department='" & r7!department & "'", c, adOpenDynamic, adLockOptimistic
            End With
        Else
            With r1
            .Open "select * from sales where saledate>='" & Year(dFROM.Value) & "-" & Month(dFROM.Value) & "-" & Day(dFROM.Value) & " 00:00:00.000' and saledate<='" & Year(dTO.Value) & "-" & Month(dTO.Value) & "-" & Day(dTO.Value) & " 00:00:00.000' and itemcodemain='" & r!itemcodemain & "' and unitprice='" & r3!unitprice & "' and department='" & r7!department & "'", c, adOpenDynamic, adLockOptimistic
            End With
        End If
    
        With r2
        .Open "select * from stock where stockcodemain='" & r!itemcodemain & "' and department='" & r7!department & "'", c, adOpenDynamic, adLockOptimistic
        End With
        tQTY = 0
        TDisc = 0
        TVAT = 0
        TTotal = 0
        sPRICE = Format(r3!unitprice, "#####.00")
        Do While r1.EOF = False
            tQTY = Format(tQTY + CDbl(r1!QTY), "####.00")
            TDisc = Format(TDisc + CDbl(r1!totdisc), "####.00")
            TVAT = Format(TVAT + CDbl(r1!vat), "####.00")
            TTotal = Format(TTotal + CDbl(r1!total), "####.00")
            r1.MoveNext
        Loop
        r1.Close
        
    
        pbody = pbody + "    <tr> "
        pbody = pbody + "      <td width=11%> "
        pbody = pbody + "        <div align=left><font face=Arial, Helvetica, sans-serif><font face=Arial, "
        pbody = pbody + "Helvetica, sans-serif><font face=Arial, Helvetica, sans-serif><font "
        pbody = pbody + "size=2>" & r!itemcodemain & "</font></font></font></font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=32%> "
        pbody = pbody + "        <div align=left><font face=Arial, Helvetica, sans-serif><font face=Arial, "
        pbody = pbody + "Helvetica, sans-serif><font face=Arial, Helvetica, sans-serif><font "
        pbody = pbody + "size=2>" & UCase(r2!stockdesc) & "</font></font></font></font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=6%>"
        pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
        pbody = pbody + "size=2>" & tQTY & "</font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=8%>"
        pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
        pbody = pbody + "size=2>" & Format(sPRICE, "####.00") & "</font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=11%>"
        pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
        pbody = pbody + "size=2>" & Format(TDisc, "####.00") & "</font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=7%>"
        pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
        pbody = pbody + "size=2>" & Format(TVAT, "####.00") & "</font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=7%>"
        pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
        pbody = pbody + "size=2>" & Format(TTotal, "####.00") & "</font></div>"
        pbody = pbody + "      </td>"
        r2.Close
        With r2
        .Open "select * from stock where stockcodemain='" & r!itemcodemain & "' and department='" & r7!department & "'", c, adOpenDynamic, adLockOptimistic
        End With
        
        pbody = pbody + "      <td width=18%>"
        pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
        pbody = pbody + "size=2>" & r2!QTY & "</font></div>"
        pbody = pbody + "      </td>"
        r2.Close
        pbody = pbody + "    </tr>"
        r3.MoveNext
        GTQTY = GTQTY + tQTY
        GTDISC = GTDISC + TDisc
        GTVAT = GTVAT + TVAT
        GTTOTAL = GTTOTAL + TTotal
        
    Loop
    r3.Close
    r.MoveNext
Loop
r.Close


        pbody = pbody + "    <tr> "
        pbody = pbody + "      <td width=11%> "
        pbody = pbody + "        <div align=left><font face=Arial, Helvetica, sans-serif><font face=Arial, "
        pbody = pbody + "Helvetica, sans-serif><font face=Arial, Helvetica, sans-serif><font "
        pbody = pbody + "size=2></font></font></font></font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=32%> "
        pbody = pbody + "        <div align=left><font face=Arial, Helvetica, sans-serif><font face=Arial, "
        pbody = pbody + "Helvetica, sans-serif><font face=Arial, Helvetica, sans-serif><font "
        pbody = pbody + "size=2></font></font></font></font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=6%>"
        pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
        pbody = pbody + "size=2><b>" & GTQTY & "</b></font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=8%>"
        pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
        pbody = pbody + "size=2></font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=11%>"
        pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
        pbody = pbody + "size=2><b>" & Format(GTDISC, "####.00") & "</b></font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=7%>"
        pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
        pbody = pbody + "size=2><b>" & Format(GTVAT, "####.00") & "</b></font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=7%>"
        pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
        pbody = pbody + "size=2><b>" & Format(GTTOTAL, "####.00") & "</b></font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=18%>"
        pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
        pbody = pbody + "size=2></font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "    </tr>"




pbody = pbody + "  </table>"
r7.MoveNext
Loop
r7.Close

pbody = pbody + "  <p align=center><b><font size=2 face=Arial, Helvetica, sans-serif>**END "
pbody = pbody + "    OF REPORT**</font></b></p>"
pbody = pbody + "</div>"
pbody = pbody + "</body>"
pbody = pbody + "</html>"


Print #2, pbody
Close #2
    
    wd.Documents.Open App.Path & "\Reports\" & Day(Date) & "#" & MonthName(Month(Date)) & "#" & Year(Date) & xHour & xMin & xSec & ".doc"
    wd.Visible = True

End Sub

Private Sub LaVolpeButton2_Click()
Unload Me

End Sub
