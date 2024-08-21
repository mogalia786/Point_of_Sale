VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmGRV 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Serial Tracking"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Select GRV#"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtsn 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   840
         Width           =   3135
      End
      Begin VB.CheckBox cSN 
         BackColor       =   &H80000016&
         Caption         =   "All GRV#'s"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "GRV#:"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&View Tracking"
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
      MICON           =   "frmGRV.frx":0000
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
      Left            =   1800
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "frmGRV.frx":001C
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
Attribute VB_Name = "frmGRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cSN_Click()
If cSN.Value = 1 Then
    txtsn.Text = ""
    txtsn.Enabled = False
Else
    txtsn.Enabled = True
    txtsn.SetFocus
End If

End Sub

Private Sub LaVolpeButton1_Click()
Dim r As New Recordset
Dim r2 As New Recordset
Dim r3 As New Recordset
Dim r4 As New Recordset
Dim r5 As New Recordset
Dim wd As New Word.Application
Dim xHour As String
Dim xMin As String
Dim xSec As String
Dim CurrentDate As Date
If cSN.Value = 0 Then
    If Len(txtsn) = 0 Then
        Exit Sub
    End If
End If

xHour = Hour(Time)
xMin = Minute(Time)
xSec = Second(Time)

If Dir(App.Path & "\Reports", vbDirectory) = "" Then
MkDir App.Path & "\Reports"
End If

Open App.Path & "\Reports\" & Day(Date) & "#" & MonthName(Month(Date)) & "#" & Year(Date) & xHour & xMin & xSec & ".doc" For Output As #2

pbody = pbody + "<html>"
pbody = pbody + "<head>"
pbody = pbody + "<title>Untitled Document</title>"
pbody = pbody + "<meta http-equiv=Content-Type content=text/html; charset=iso-8859-1>"
pbody = pbody + "</head>"

pbody = pbody + "<body bgcolor=#FFFFFF text=#000000>"
pbody = pbody + "<div align=center>"
pbody = pbody + "  <p><b><font face=Times New Roman, Times, serif><u>" & CompName & "</u></font></b></p>"

pbody = pbody + "  <p><b><u>GRV Report</u></b></p>"
pbody = pbody + "  <p align=left><b>Date: " & Now & "</b></p>"
If cSN.Value = 1 Then
pbody = pbody + "  <p align=left><b>GRV Criteria: ALL</b></p>"
Else
pbody = pbody + "  <p align=left><b>GRV Criteria: " & UCase(txtsn.Text) & "</b></p>"
End If
pbody = pbody + "  <table width=100% border=1 cellspacing=1 cellpadding=0>"
pbody = pbody + "    <tr bgcolor=#000000> "

pbody = pbody + "      <td width=5%> "
pbody = pbody + "        <div align=left><font color=#FFFFFF><b><font size=2 face=Arial, Helvetica, "
pbody = pbody + "sans-serif>GRV#</font></b></font></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=6%> "
pbody = pbody + "        <div align=left><b><font face=Arial, Helvetica, sans-serif size=2 "
pbody = pbody + "color=#FFFFFF>Date</font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=8%> "
pbody = pbody + "        <div align=left><b><font size=2 color=#FFFFFF><font color=#FFFFFF><font "
pbody = pbody + "face=Arial, Helvetica, sans-serif>S/Code</font></font></font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=8%> "
pbody = pbody + "        <div align=left><b><font size=2 color=#FFFFFF>B/Code</font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=14%> "
pbody = pbody + "        <div align=left><b><font size=2 color=#FFFFFF>Serial#</font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=20%> "
pbody = pbody + "        <div align=left><b><font size=2 color=#FFFFFF><font color=#FFFFFF><font "
pbody = pbody + "color=#FFFFFF>Supplier</font></font></font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=8%> "
pbody = pbody + "        <div align=left><b><font size=2 color=#FFFFFF><font color=#FFFFFF><font "
pbody = pbody + "color=#FFFFFF>Cost(Excl)</font></font></font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=8%> "
pbody = pbody + "        <div align=left><b><font size=2 color=#FFFFFF><font color=#FFFFFF><font "
pbody = pbody + "color=#FFFFFF>VAT(in)</font></font></font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=8%> "
pbody = pbody + "        <div align=left><b><font size=2 color=#FFFFFF><font color=#FFFFFF><font "
pbody = pbody + "color=#FFFFFF>Cost(Incl)</font></font></font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=8%> "
pbody = pbody + "        <div align=left><b><font size=2 color=#FFFFFF>Price(Incl)</font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=7%> "
pbody = pbody + "        <div align=left><b><font size=2 color=#FFFFFF><font color=#FFFFFF><font "
pbody = pbody + "color=#FFFFFF>Qty</font></font></font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "    </tr>"

If cSN.Value = 1 Then
    With r
    .Open "SELECT * from stockpurchasehistory order by pid", c, adOpenDynamic, adLockOptimistic
    End With
Else
    With r
    .Open "SELECT * from stockpurchasehistory where pid='" & txtsn.Text & "' order by pid", c, adOpenDynamic, adLockOptimistic
    End With
End If

Do While r.EOF = False


    pbody = pbody + "    <tr> "
    pbody = pbody + "<td width=5%><font size=1 face=Arial, Helvetica, sans-serif>" & r!pid & "</font></td>"
    
    pbody = pbody + "<td width=6%><font size=1 face=Arial, Helvetica, sans-serif>" & r!datepurchased & "</font></td>"
    
    pbody = pbody + "<td width=8%><font size=1 face=Arial, Helvetica, sans-serif>" & r!stockcodeMAIN & "</font></td>"
    
    pbody = pbody + "<td width=8%><font size=1 face=Arial, Helvetica, sans-serif>" & r!stockcode & "</font></td>"
    
    pbody = pbody + "<td width=14%><font size=1 face=Arial, Helvetica, sans-serif>" & r!serialnumber & "</font></td>"
    
    pbody = pbody + "<td width=20%><font size=1 face=Arial, Helvetica, sans-serif>" & r!supplier & "</font></td>"
    
    pbody = pbody + "<td width=8%><font size=1 face=Arial, Helvetica, sans-serif>" & r!costofItemEXC & "</font></td>"
    
    pbody = pbody + "<td width=8%><font size=1 face=Arial, Helvetica, sans-serif>" & r!vatinput & "</font></td>"
    
    pbody = pbody + "<td width=8%><font size=1 face=Arial, Helvetica, sans-serif>" & r!costofiteminc & "</font></td>"
    
    pbody = pbody + "<td width=8%><font size=1 face=Arial, Helvetica, sans-serif>" & r!sellingpriceINc & "</font></td>"
    
    pbody = pbody + "<td width=7%><font size=1 face=Arial, Helvetica, sans-serif>" & r!qtypurchased & "</font></td>"
    
    pbody = pbody + "    </tr>"
r.MoveNext
Loop
r.Close
pbody = pbody + "  </table>"
pbody = pbody + "  <p align=left>&nbsp;</p>"
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


