VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRepSales 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales  Report"
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
      MICON           =   "frmRepSales.frx":0000
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
         Format          =   125698049
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
         Format          =   125698049
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
      MICON           =   "frmRepSales.frx":001C
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
Attribute VB_Name = "frmRepSales"
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
Dim wd As New Word.Application
Dim xHour As String
Dim xMin As String
Dim xSec As String
Dim CurrentDate As Date
Dim TTotalEXC As Double
Dim TTotalINC As Double
Dim TVAT As Double
Dim TDisc As Double
Dim GTTotalEXC As Double
Dim GTDISC As Double
Dim GTVAT As Double
Dim GTTOTALINC As Double

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




pbody = pbody + "  <p><b><font face=Times New Roman, Times, serif><u>SALES "

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

 '************************************TOTAL SALES
pbody = pbody + "  <p align=left><b><font face=Times New Roman, Times, serif size=2>"
    pbody = pbody + "Total Sales</font></b></p>"
    
    pbody = pbody + "  <table width=75% border=1 cellspacing=1 cellpadding=0>"
    pbody = pbody + "    <tr bgcolor=#000000> "
    pbody = pbody + "      <td width=17%> "
    pbody = pbody + "        <div align=left><b><font color=#FFFFFF>Date</font></b></div>"
    pbody = pbody + "      </td>"
    pbody = pbody + "      <td width=17%> "
    pbody = pbody + "        <div align=left><b><b><font size=2><font "
    
    pbody = pbody + "color=#FFFFFF>TOTAL Excl.</font></font></b></b></div>"
    pbody = pbody + "      </td>"
    pbody = pbody + "      <td width=17%> "
    pbody = pbody + "        <div align=center><b><b><font size=2><font "
    
    pbody = pbody + "color=#FFFFFF>V.A.T</font></font></b></b></div>"
    pbody = pbody + "      </td>"
    pbody = pbody + "      <td width=17%> "
    pbody = pbody + "        <div align=center><b><b><font size=2><font "
    
    pbody = pbody + "color=#FFFFFF>TOTAL Incl.</font></font></b></b></div>"
    pbody = pbody + "      </td>"
    pbody = pbody + "      <td width=17%> "
    pbody = pbody + "        <div align=center><b><b><font size=2><font "
    
    pbody = pbody + "color=#FFFFFF>Discount</font></font></b></b></div>"
    pbody = pbody + "      </td>"
    pbody = pbody + "      <td width=17%> "
    pbody = pbody + "        <div align=center><b><b><font size=2><font "
    
    pbody = pbody + "color=#FFFFFF>Cash Taken</font></font></b></b></div>"
    pbody = pbody + "      </td>"
    
    pbody = pbody + "    </tr>"
    If c1.Value = 1 Then
        With r
        .Open "select distinct(saledate) from sales order by saledate", c, adOpenDynamic, adLockOptimistic
        End With
    Else
        With r
        .Open "select distinct(saledate) from sales where saledate>='" & Year(dFROM.Value) & "-" & Month(dFROM.Value) & "-" & Day(dFROM.Value) & " 00:00:00.000' and saledate<='" & Year(dTO.Value) & "-" & Month(dTO.Value) & "-" & Day(dTO.Value) & " 00:00:00.000' order by saledate", c, adOpenDynamic, adLockOptimistic
        End With
    End If
    Do While r.EOF = False
        
        With r1
        .Open "select * from sales where saledate='" & Year(r!saledate) & "-" & Month(r!saledate) & "-" & Day(r!saledate) & " 00:00:00.000'", c, adOpenDynamic, adLockOptimistic
        End With
        
            TTotalEXC = 0
            TDisc = 0
            TVAT = 0
            TTotalINC = 0
            
        Do While r1.EOF = False
            TTotalEXC = Format(TTotalEXC + (CDbl(r1!Total) - CDbl(r1!vat)), "####.00")
            TDisc = Format(TDisc + CDbl(r1!TOTDISC), "####.00")
            TVAT = Format(TVAT + CDbl(r1!vat), "####.00")
            TTotalINC = Format(TTotalINC + CDbl(r1!Total), "####.00")
            r1.MoveNext
        Loop
        r1.Close
            
        
            pbody = pbody + "    <tr> "
            
            pbody = pbody + "      <td width=17%> "
            pbody = pbody + "        <div align=left><font face=Arial, Helvetica, sans-serif><font face=Arial, "
            pbody = pbody + "Helvetica, sans-serif><font face=Arial, Helvetica, sans-serif><font "
            pbody = pbody + "size=2>" & r!saledate & "</font></font></font></font></div>"
            pbody = pbody + "      </td>"
            
            pbody = pbody + "      <td width=17%> "
            pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif><font face=Arial, "
            pbody = pbody + "Helvetica, sans-serif><font face=Arial, Helvetica, sans-serif><font "
            pbody = pbody + "size=2>" & Format(TTotalEXC, "####.00") & "</font></font></font></font></div>"
            pbody = pbody + "      </td>"
            
            pbody = pbody + "      <td width=17%>"
            pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
            pbody = pbody + "size=2>" & Format(TVAT, "####.00") & "</font></div>"
            pbody = pbody + "      </td>"
            
            pbody = pbody + "      <td width=17%>"
            pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
            pbody = pbody + "size=2>" & Format(TTotalINC, "####.00") & "</font></div>"
            pbody = pbody + "      </td>"
            
            pbody = pbody + "      <td width=17%>"
            pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
            pbody = pbody + "size=2>" & Format(TDisc, "####.00") & "</font></div>"
            pbody = pbody + "      </td>"
            
            pbody = pbody + "      <td width=17%>"
            pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
            pbody = pbody + "size=2>" & Format(TTotalINC, "####.00") & "</font></div>"
            pbody = pbody + "      </td>"
            
            
            pbody = pbody + "    </tr>"
            r.MoveNext
            GTTotalEXC = GTTotalEXC + TTotalEXC
            GTDISC = GTDISC + TDisc
            GTVAT = GTVAT + TVAT
            GTTOTALINC = GTTOTALINC + TTotalINC
            
        Loop
        r.Close
    
    
            pbody = pbody + "    <tr> "
            pbody = pbody + "      <td width=17%> "
            pbody = pbody + "        <div align=left><font face=Arial, Helvetica, sans-serif><font face=Arial, "
            pbody = pbody + "Helvetica, sans-serif><font face=Arial, Helvetica, sans-serif><font "
            pbody = pbody + "size=2></font></font></font></font></div>"
            pbody = pbody + "      </td>"
            
            pbody = pbody + "      <td width=17%> "
            pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif><font face=Arial, "
            pbody = pbody + "Helvetica, sans-serif><font face=Arial, Helvetica, sans-serif><font "
            pbody = pbody + "size=2><b>" & Format(GTTotalEXC, "####.00") & "</b></font></font></font></font></div>"
            pbody = pbody + "      </td>"
            
            pbody = pbody + "      <td width=17%>"
            pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
            pbody = pbody + "size=2><b>" & Format(GTVAT, "####.00") & "</b></font></div>"
            pbody = pbody + "      </td>"
            
            pbody = pbody + "      <td width=17%>"
            pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
            pbody = pbody + "size=2><b>" & Format(GTTOTALINC, "####.00") & "<b></font></div>"
            pbody = pbody + "      </td>"
            
            pbody = pbody + "      <td width=17%>"
            pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
            pbody = pbody + "size=2><b>" & Format(GTDISC, "####.00") & "</b></font></div>"
            pbody = pbody + "      </td>"
            
            pbody = pbody + "      <td width=17%>"
            pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
            pbody = pbody + "size=2><b>" & Format(GTTOTALINC, "####.00") & "</b></font></div>"
            pbody = pbody + "      </td>"
            
            
            pbody = pbody + "    </tr>"
    
    
    GTTotalEXC = 0
    GTTOTALINC = 0
    GTDISC = 0
    GTVAT = 0
    
    
    pbody = pbody + "  </table>"

 
'****************************************END TOTAL SALES

With r3
.Open "select * from department order by department", c, adOpenDynamic, adLockOptimistic
End With
Do While r3.EOF = False
    pbody = pbody + "  <p align=left><b><font face=Times New Roman, Times, serif size=2>"
    pbody = pbody + "Department : " & r3!department & "</font></b></p>"
    
    pbody = pbody + "  <table width=75% border=1 cellspacing=1 cellpadding=0>"
    pbody = pbody + "    <tr bgcolor=#000000> "
    pbody = pbody + "      <td width=17%> "
    pbody = pbody + "        <div align=left><b><font color=#FFFFFF>Date</font></b></div>"
    pbody = pbody + "      </td>"
    pbody = pbody + "      <td width=17%> "
    pbody = pbody + "        <div align=left><b><b><font size=2><font "
    
    pbody = pbody + "color=#FFFFFF>TOTAL Excl.</font></font></b></b></div>"
    pbody = pbody + "      </td>"
    pbody = pbody + "      <td width=17%> "
    pbody = pbody + "        <div align=center><b><b><font size=2><font "
    
    pbody = pbody + "color=#FFFFFF>V.A.T</font></font></b></b></div>"
    pbody = pbody + "      </td>"
    pbody = pbody + "      <td width=17%> "
    pbody = pbody + "        <div align=center><b><b><font size=2><font "
    
    pbody = pbody + "color=#FFFFFF>TOTAL Incl.</font></font></b></b></div>"
    pbody = pbody + "      </td>"
    pbody = pbody + "      <td width=17%> "
    pbody = pbody + "        <div align=center><b><b><font size=2><font "
    
    pbody = pbody + "color=#FFFFFF>Discount</font></font></b></b></div>"
    pbody = pbody + "      </td>"
    pbody = pbody + "      <td width=17%> "
    pbody = pbody + "        <div align=center><b><b><font size=2><font "
    
    pbody = pbody + "color=#FFFFFF>Cash Taken</font></font></b></b></div>"
    pbody = pbody + "      </td>"
    
    pbody = pbody + "    </tr>"
    If c1.Value = 1 Then
        With r
        .Open "select distinct(saledate) from sales where department='" & r3!department & "' order by saledate", c, adOpenDynamic, adLockOptimistic
        End With
    Else
        With r
        .Open "select distinct(saledate) from sales where saledate>='" & Year(dFROM.Value) & "-" & Month(dFROM.Value) & "-" & Day(dFROM.Value) & " 00:00:00.000' and saledate<='" & Year(dTO.Value) & "-" & Month(dTO.Value) & "-" & Day(dTO.Value) & " 00:00:00.000' and department='" & r3!department & "' order by saledate", c, adOpenDynamic, adLockOptimistic
        End With
    End If
    Do While r.EOF = False
        
        With r1
        .Open "select * from sales where saledate='" & Year(r!saledate) & "-" & Month(r!saledate) & "-" & Day(r!saledate) & " 00:00:00.000' and department='" & r3!department & "'", c, adOpenDynamic, adLockOptimistic
        End With
        
            TTotalEXC = 0
            TDisc = 0
            TVAT = 0
            TTotalINC = 0
            
        Do While r1.EOF = False
            TTotalEXC = Format(TTotalEXC + (CDbl(r1!Total) - CDbl(r1!vat)), "####.00")
            TDisc = Format(TDisc + CDbl(r1!TOTDISC), "####.00")
            TVAT = Format(TVAT + CDbl(r1!vat), "####.00")
            TTotalINC = Format(TTotalINC + CDbl(r1!Total), "####.00")
            r1.MoveNext
        Loop
        r1.Close
            
        
            pbody = pbody + "    <tr> "
            
            pbody = pbody + "      <td width=17%> "
            pbody = pbody + "        <div align=left><font face=Arial, Helvetica, sans-serif><font face=Arial, "
            pbody = pbody + "Helvetica, sans-serif><font face=Arial, Helvetica, sans-serif><font "
            pbody = pbody + "size=2>" & r!saledate & "</font></font></font></font></div>"
            pbody = pbody + "      </td>"
            
            pbody = pbody + "      <td width=17%> "
            pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif><font face=Arial, "
            pbody = pbody + "Helvetica, sans-serif><font face=Arial, Helvetica, sans-serif><font "
            pbody = pbody + "size=2>" & Format(TTotalEXC, "####.00") & "</font></font></font></font></div>"
            pbody = pbody + "      </td>"
            
            pbody = pbody + "      <td width=17%>"
            pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
            pbody = pbody + "size=2>" & Format(TVAT, "####.00") & "</font></div>"
            pbody = pbody + "      </td>"
            
            pbody = pbody + "      <td width=17%>"
            pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
            pbody = pbody + "size=2>" & Format(TTotalINC, "####.00") & "</font></div>"
            pbody = pbody + "      </td>"
            
            pbody = pbody + "      <td width=17%>"
            pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
            pbody = pbody + "size=2>" & Format(TDisc, "####.00") & "</font></div>"
            pbody = pbody + "      </td>"
            
            pbody = pbody + "      <td width=17%>"
            pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
            pbody = pbody + "size=2>" & Format(TTotalINC, "####.00") & "</font></div>"
            pbody = pbody + "      </td>"
            
            
            pbody = pbody + "    </tr>"
            r.MoveNext
            GTTotalEXC = GTTotalEXC + TTotalEXC
            GTDISC = GTDISC + TDisc
            GTVAT = GTVAT + TVAT
            GTTOTALINC = GTTOTALINC + TTotalINC
            
        Loop
        r.Close
    
    
            pbody = pbody + "    <tr> "
            pbody = pbody + "      <td width=17%> "
            pbody = pbody + "        <div align=left><font face=Arial, Helvetica, sans-serif><font face=Arial, "
            pbody = pbody + "Helvetica, sans-serif><font face=Arial, Helvetica, sans-serif><font "
            pbody = pbody + "size=2></font></font></font></font></div>"
            pbody = pbody + "      </td>"
            
            pbody = pbody + "      <td width=17%> "
            pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif><font face=Arial, "
            pbody = pbody + "Helvetica, sans-serif><font face=Arial, Helvetica, sans-serif><font "
            pbody = pbody + "size=2><b>" & Format(GTTotalEXC, "####.00") & "</b></font></font></font></font></div>"
            pbody = pbody + "      </td>"
            
            pbody = pbody + "      <td width=17%>"
            pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
            pbody = pbody + "size=2><b>" & Format(GTVAT, "####.00") & "</b></font></div>"
            pbody = pbody + "      </td>"
            
            pbody = pbody + "      <td width=17%>"
            pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
            pbody = pbody + "size=2><b>" & Format(GTTOTALINC, "####.00") & "<b></font></div>"
            pbody = pbody + "      </td>"
            
            pbody = pbody + "      <td width=17%>"
            pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
            pbody = pbody + "size=2><b>" & Format(GTDISC, "####.00") & "</b></font></div>"
            pbody = pbody + "      </td>"
            
            pbody = pbody + "      <td width=17%>"
            pbody = pbody + "        <div align=right><font face=Arial, Helvetica, sans-serif "
            pbody = pbody + "size=2><b>" & Format(GTTOTALINC, "####.00") & "</b></font></div>"
            pbody = pbody + "      </td>"
            
            
            pbody = pbody + "    </tr>"
    
    
    GTTotalEXC = 0
    GTTOTALINC = 0
    GTDISC = 0
    GTVAT = 0
    
    
    pbody = pbody + "  </table>"
r3.MoveNext
Loop
r3.Close

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
