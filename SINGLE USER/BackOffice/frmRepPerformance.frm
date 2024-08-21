VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRepPerformance 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Performance Report"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   2160
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
      MICON           =   "frmRepPerformance.frx":0000
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
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4455
      Begin VB.ComboBox cCrit 
         Height          =   315
         ItemData        =   "frmRepPerformance.frx":001C
         Left            =   960
         List            =   "frmRepPerformance.frx":002C
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CheckBox c1 
         BackColor       =   &H80000016&
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Order By:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
      Top             =   2160
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
      MICON           =   "frmRepPerformance.frx":004D
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
Attribute VB_Name = "frmRepPerformance"
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
cCrit.Text = "Sales"

End Sub

Private Sub LaVolpeButton1_Click()
Dim pbody As String
Dim r As New Recordset
Dim r1 As New Recordset
Dim r2 As New Recordset
Dim r3 As New Recordset
Dim r4 As New Recordset
Dim r5 As New Recordset
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
If Len(cCrit.Text) = 0 Then
    MsgBox "Please select order by criteria!", vbExclamation, "Status"
    cCrit.SetFocus
    Exit Sub
End If

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

pbody = pbody + "  <p><b><font face=Times New Roman, Times, serif><u>SALES "

pbody = pbody + "PERFORMANCE</u></font></b></p>"


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
If cCrit.Text = "Sales" Then
    pbody = pbody + "  <p align=left><b><font face=Times New Roman, Times, serif size=2>Order By "
    pbody = pbody + "    : Sales</font></b></p>"
End If
If cCrit.Text = "GP" Then
    pbody = pbody + "  <p align=left><b><font face=Times New Roman, Times, serif size=2>Order By "
    pbody = pbody + "    : Gross Profit</font></b></p>"
End If
If cCrit.Text = "Qty Sold" Then
    pbody = pbody + "  <p align=left><b><font face=Times New Roman, Times, serif size=2>Order By "
    pbody = pbody + "    : Quantity Sold</font></b></p>"
End If
If cCrit.Text = "Profit" Then
    pbody = pbody + "  <p align=left><b><font face=Times New Roman, Times, serif size=2>Order By "
    pbody = pbody + "    : Profit</font></b></p>"
End If
With r7
.Open "select * from department order by department", c, adOpenDynamic, adLockOptimistic
End With
Do While r7.EOF = False
    pbody = pbody + "  <p align=left><b><font face=Times New Roman, Times, serif size=2>"
    pbody = pbody + "Department : " & r7!department & "</font></b></p>"


pbody = pbody + "  <table width=100% border=1 cellspacing=1 cellpadding=0>"
pbody = pbody + "    <tr bgcolor=#000000> "
pbody = pbody + "      <td width=10%> "
pbody = pbody + "        <div align=left><b><font color=#FFFFFF>Stock Code</font></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=20%> "
pbody = pbody + "        <div align=left><b><b><font size=2><font "

pbody = pbody + "color=#FFFFFF>Description</font></font></b></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=10%> "
pbody = pbody + "        <div align=center><b><b><font size=2><font "

pbody = pbody + "color=#FFFFFF>Cost Excl</font></font></b></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=10%> "
pbody = pbody + "        <div align=center><b><b><font size=2><font "

pbody = pbody + "color=#FFFFFF>Onhand</font></font></b></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=11%> "
pbody = pbody + "        <div align=center><b><b><font size=2><font "

pbody = pbody + "color=#FFFFFF>Sales Qty</font></font></b></b></div>"
pbody = pbody + "      </td>"
pbody = pbody + "      <td width=10%> "
pbody = pbody + "        <div align=center><b><b><font size=2><font "

pbody = pbody + "color=#FFFFFF>Cost</font></font></b></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=10%> "
pbody = pbody + "        <div align=center><b><b><font size=2><font "
pbody = pbody + "color=#FFFFFF>Sales</font></font></b></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=10%> "
pbody = pbody + "        <div align=center><b><b><font size=2><font "
pbody = pbody + "color=#FFFFFF>Profit</font></font></b></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=10%> "
pbody = pbody + "        <div align=center><b><b><font size=2><font "
pbody = pbody + "color=#FFFFFF>GP %</font></font></b></b></div>"
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
            With r1
            .Open "select * from sales where itemcodemain='" & r!itemcodemain & "' and department='" & r7!department & "'", c, adOpenDynamic, adLockOptimistic
            End With
        Else
            With r1
            .Open "select * from sales where saledate>='" & Year(dFROM.Value) & "-" & Month(dFROM.Value) & "-" & Day(dFROM.Value) & " 00:00:00.000' and saledate<='" & Year(dTO.Value) & "-" & Month(dTO.Value) & "-" & Day(dTO.Value) & " 00:00:00.000' and itemcodemain='" & r!itemcodemain & "' and department='" & r7!department & "'", c, adOpenDynamic, adLockOptimistic
            End With
        End If
    
        With r2
        .Open "select * from stock where stockcodemain='" & r!itemcodemain & "' and department='" & r7!department & "'", c, adOpenDynamic, adLockOptimistic
        End With
        With r4
        .Open "select * from stockpurchasehistory where stockcodemain='" & r!itemcodemain & "' and department='" & r7!department & "'", c, adOpenDynamic, adLockOptimistic
        End With
        
        tQTY = 0
        TDisc = 0
        TVAT = 0
        TTotal = 0
        
        Do While r1.EOF = False
            tQTY = Format(tQTY + CDbl(r1!QTY), "####.00")
            TDisc = Format(TDisc + CDbl(r1!totdisc), "####.00")
            TVAT = Format(TVAT + CDbl(r1!vat), "####.00")
            TTotal = Format(TTotal + CDbl(r1!total), "####.00")
            r1.MoveNext
        Loop
        r1.Close
        
        With r5
        .Open "select * from repperformance", c, adOpenDynamic, adLockOptimistic
        End With
        r5.AddNew
        r5!itemcode = r!itemcodemain
        r5!descr = r2!stockdesc
        r5!costexcl = r4!costofItemEXC
        r5!onhand = r2!QTY
        r5!saleqty = tQTY
        r5!cost = Format(CDbl(r5!costexcl) * CDbl(r5!saleqty), "#####.00")
        r5!sales = Format(TTotal - TVAT, "#####.00")
        r5!profit = Format(CDbl(r5!sales) - CDbl(r5!cost), "#####.00")
        r5!gp = Round((CDbl(r5!profit) / CDbl(r5!sales)) * 100, 2)
        r5.Update
        r5.Close
        r2.Close
        r4.Close
        
        r.MoveNext
        
        
Loop
r.Close
If cCrit.Text = "Sales" Then

    With r
    .Open "select * from repperformance order by sales desc", c, adOpenDynamic, adLockOptimistic
    End With
    
End If
If cCrit.Text = "GP" Then

    With r
    .Open "select * from repperformance order by gp desc", c, adOpenDynamic, adLockOptimistic
    End With
    
End If
If cCrit.Text = "Qty Sold" Then

    With r
    .Open "select * from repperformance order by saleqty desc", c, adOpenDynamic, adLockOptimistic
    End With
    
End If
If cCrit.Text = "Profit" Then

    With r
    .Open "select * from repperformance order by profit desc", c, adOpenDynamic, adLockOptimistic
    End With
    
End If

Do While r.EOF = False

'/////////////
        pbody = pbody + "    <tr> "
        pbody = pbody + "      <td width=10%> "
        pbody = pbody + "        <div align=left><font face=Arial, Helvetica, sans-serif><font face=Arial, "
        pbody = pbody + "Helvetica, sans-serif><font face=Arial, Helvetica, sans-serif><font "
        pbody = pbody + "size=2>" & r!itemcode & "</font></font></font></font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=20%> "
        pbody = pbody + "        <div align=left><font face=Arial, Helvetica, sans-serif><font face=Arial, "
        pbody = pbody + "Helvetica, sans-serif><font face=Arial, Helvetica, sans-serif><font "
        pbody = pbody + "size=2>" & UCase(r!descr) & "</font></font></font></font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=10%>"
        pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
        pbody = pbody + "size=2>" & Format(r!costexcl, "#####.00") & "</font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=10%>"
        pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
        pbody = pbody + "size=2>" & r!onhand & "</font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=10%>"
        pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
        pbody = pbody + "size=2>" & r!saleqty & "</font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=10%>"
        pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
        pbody = pbody + "size=2>" & Format(r!cost, "####.00") & "</font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=10%>"
        pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
        pbody = pbody + "size=2>" & Format(r!sales, "####.00") & "</font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=10%>"
        pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
        pbody = pbody + "size=2>" & Format(r!profit, "#####.00") & "</font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=10%>"
        pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
        pbody = pbody + "size=2>" & r!gp & "</font></div>"
        pbody = pbody + "      </td>"
       
        pbody = pbody + "    </tr>"
        
        r.MoveNext
Loop
r.Close
With r
.Open "delete from repperformance", c, adOpenDynamic, adLockOptimistic
End With


'/////////////





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
