VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H80000016&
   Caption         =   "Carousel BackOffice"
   ClientHeight    =   3090
   ClientLeft      =   225
   ClientTop       =   1170
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1800
      Top             =   1320
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   2280
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   4022
      ButtonWidth     =   2196
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Log Off"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "View Supplier"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "View Stock"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add Stock"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sales Report"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Performance"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Serial Tracking"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stock on hand"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Product Sales"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Returns"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   17640
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":158A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2280
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnulogoff 
         Caption         =   "&log Off"
         Shortcut        =   ^{F12}
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuadmin 
      Caption         =   "&Administration"
      Begin VB.Menu mnucreatenewuser 
         Caption         =   "&Create New User Account"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuedituser 
         Caption         =   "&Edit / View User Account"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnusuper 
         Caption         =   "Create S&upervisor Account"
      End
      Begin VB.Menu mnuviewLoginLogsheet 
         Caption         =   "View &Login Logsheet"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuviewtrans 
         Caption         =   "View &Transaction Logsheet"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnucreatecostcode 
         Caption         =   "C&reate Cost Code"
      End
      Begin VB.Menu mnusetupmarkup 
         Caption         =   "&Setup Markup"
      End
      Begin VB.Menu mnutillsetup 
         Caption         =   "&Till Slip and Display Pole setup"
         Begin VB.Menu mnuhead 
            Caption         =   "&Setup Header and Footer on Till Slip"
         End
         Begin VB.Menu mnupole 
            Caption         =   "Setup &Pole Display greeting "
         End
         Begin VB.Menu mnuopendrawer 
            Caption         =   "Set &Till open drawer code"
         End
      End
      Begin VB.Menu mnutaxcode 
         Caption         =   "Create Ta&x Codes"
      End
      Begin VB.Menu mnucreatedep 
         Caption         =   "&Create Departments"
      End
      Begin VB.Menu deletesales 
         Caption         =   "Delete Sales"
      End
      Begin VB.Menu mnusetupcompany 
         Caption         =   "Setup Co&mpany"
      End
   End
   Begin VB.Menu mnustock 
      Caption         =   "&Stock"
      Begin VB.Menu mnucreatedisc 
         Caption         =   "Create &Discount Code"
      End
      Begin VB.Menu mnuassDiscCodetoStock 
         Caption         =   "&Assign Discount Code to Stock"
      End
      Begin VB.Menu mnuaddStock 
         Caption         =   "&Create/Add  Stock"
      End
      Begin VB.Menu mnusetstocktax 
         Caption         =   "Set/View Stock &Tax Codes"
      End
      Begin VB.Menu mnuviewStockDetails 
         Caption         =   "&View Stock Details"
      End
      Begin VB.Menu mnuAdjust 
         Caption         =   "A&djustments"
      End
      Begin VB.Menu delete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuaddSerial 
         Caption         =   "Add &Serial# to stock"
      End
      Begin VB.Menu mnupack 
         Caption         =   "Pack"
         Begin VB.Menu mnucreatepackexist 
            Caption         =   "Create pack from existing stock (using serial numbers)"
         End
         Begin VB.Menu mnucreatepacknewstock 
            Caption         =   "Create pack from existing stock (using barcodes)"
         End
         Begin VB.Menu mnusplitpack 
            Caption         =   "&Split Pack"
         End
      End
   End
   Begin VB.Menu mnureturn 
      Caption         =   "&Returns"
      Begin VB.Menu mnuviewreturns 
         Caption         =   "&View Returns"
      End
   End
   Begin VB.Menu mnusupplier 
      Caption         =   "&Supplier"
      Begin VB.Menu mnuaddsup 
         Caption         =   "&Add Supplier"
      End
      Begin VB.Menu mnuviewsupplier 
         Caption         =   "&View Supplier Details"
      End
      Begin VB.Menu mnueditsup 
         Caption         =   "&Edit Supplier details / Opening Balance"
      End
   End
   Begin VB.Menu mnudebtor 
      Caption         =   "&Debtors"
      Visible         =   0   'False
   End
   Begin VB.Menu mnucreditor 
      Caption         =   "&Creditors"
      Begin VB.Menu mnuviewcredacc 
         Caption         =   "&View Creditor's Account"
      End
      Begin VB.Menu mnupaycred 
         Caption         =   "&Make Payment to Creditor"
      End
      Begin VB.Menu mnucreditnotes 
         Caption         =   "&Credit Notes"
      End
      Begin VB.Menu mnuviewbalinv 
         Caption         =   "View &Balanced owed on Invoice"
      End
   End
   Begin VB.Menu mnureport 
      Caption         =   "&Report"
      Begin VB.Menu mnuprodsalessummary 
         Caption         =   "&Product Sales Summary"
      End
      Begin VB.Menu mnustoconhand 
         Caption         =   "&Stock on Hand"
      End
      Begin VB.Menu mnusalesrep 
         Caption         =   "S&ales Report"
      End
      Begin VB.Menu mnusalesperformers 
         Caption         =   "Sales &Performers"
      End
      Begin VB.Menu mnureturns 
         Caption         =   "&Returns"
      End
      Begin VB.Menu mnupayments 
         Caption         =   "Pa&yments"
      End
      Begin VB.Menu mnustockpurchase 
         Caption         =   "S&tock Purchases"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuserialtracking 
         Caption         =   "Serial Trac&king"
      End
      Begin VB.Menu mnugrv 
         Caption         =   "&GRV"
      End
      Begin VB.Menu mnusalesvoucher 
         Caption         =   "Print Sales &Voucher"
      End
   End
   Begin VB.Menu mnulabels 
      Caption         =   "Labels"
      Begin VB.Menu mnuprintlabels 
         Caption         =   "&Print Labels"
      End
      Begin VB.Menu mnucreatecostcodes 
         Caption         =   "&Create Cost Codes"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Delete_Click()
frmStockDelete.Show vbModal

End Sub

Private Sub deletesales_Click()
rmDeleteSales.Show vbModal
End Sub

Private Sub MDIForm_Load()
Dim r As New Recordset
Dim cs As New Connection
Dim rs As New Recordset
With cs
.Provider = "microsoft.jet.oledb.4.0"
.ConnectionString = App.Path & "/temp.mdb"
.Open
End With
With r
.Open "select * from company", c, adOpenDynamic, adLockOptimistic
End With
CompName = r!cname
r.Close

With rs
.Open "select * from readersetup", cs, adOpenDynamic, adLockOptimistic
End With
ReaderPort = rs!PortNo
rs.Close

    mnuadmin.Enabled = False
    'sale.Enabled = False
    'Card.Enabled = False
    'Exhibitor.Enabled = False
    'Comp.Enabled = False
    'Ticket.Enabled = False
    'Report.Enabled = False

With r
.Open "select * from logon where username='" & CurrentUser & "'", c, adOpenDynamic, adLockOptimistic
End With
If r!Admin = "Yes" Then
    mnuadmin.Enabled = True
    mnustock.Enabled = True
    mnureturn.Enabled = True
    mnusupplier.Enabled = True
    mnudebtor.Enabled = True
    mnucreditor.Enabled = True
    mnureport.Enabled = True
    Toolbar1.Buttons(1).Enabled = True
    Toolbar1.Buttons(2).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(4).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    Toolbar1.Buttons(6).Enabled = True
    Toolbar1.Buttons(7).Enabled = True
    Toolbar1.Buttons(8).Enabled = True
    Toolbar1.Buttons(9).Enabled = True
    Toolbar1.Buttons(10).Enabled = True
    Toolbar1.Buttons(11).Enabled = True

    Exit Sub
Else
    mnuadmin.Enabled = False

    If r!canstock = "Yes" Then
        mnustock.Enabled = True
        Toolbar1.Buttons(4).Enabled = True

    Else
        mnustock.Enabled = False
        Toolbar1.Buttons(4).Enabled = False
        Toolbar1.Buttons(5).Enabled = False

    End If
    If r!canreturn = "Yes" Then
        mnureturn.Enabled = True
            Toolbar1.Buttons(11).Enabled = True

    Else
        mnureturn.Enabled = False
        Toolbar1.Buttons(11).Enabled = False

    End If
    If r!candebtor = "Yes" Then
        mnudebtor.Enabled = True
    
    Else
        mnudebtor.Enabled = False
    End If
    If r!cancreditor = "Yes" Then
        mnucreditor.Enabled = True
            

    Else
        mnucreditor.Enabled = False
    End If
    If r!canreport = "Yes" Then
        mnureport.Enabled = True
        Toolbar1.Buttons(6).Enabled = True
        Toolbar1.Buttons(7).Enabled = True
        Toolbar1.Buttons(8).Enabled = True
        Toolbar1.Buttons(9).Enabled = True
        Toolbar1.Buttons(10).Enabled = True

    Else
        mnureport.Enabled = False
        Toolbar1.Buttons(6).Enabled = False
        Toolbar1.Buttons(7).Enabled = False
        Toolbar1.Buttons(8).Enabled = False
        Toolbar1.Buttons(9).Enabled = False
        Toolbar1.Buttons(10).Enabled = False
       


    End If
    If r!cansupplier = "Yes" Then
        mnusupplier.Enabled = True
        Toolbar1.Buttons(3).Enabled = True

    Else
        mnusupplier.Enabled = False
        Toolbar1.Buttons(3).Enabled = False

    End If
    
        
End If

r.Close

End Sub

Private Sub mnuaddexh_Click()
frmCreateExhib.Show vbModal

End Sub

Private Sub mnuasscardtoexh_Click()
frmCardAssign.Show vbModal

End Sub

Private Sub mnublock_Click()
frmCardBlock.Show vbModal

End Sub

Private Sub mnucreatecomp_Click()
frmCompCreate.Show vbModal

End Sub

Private Sub mnublank_Click()

End Sub

Private Sub mnuaddSerial_Click()
frmAddSerial.Show vbModal
End Sub

Private Sub mnuaddStock_Click()
frmStockAdd.Show vbModal

End Sub

Private Sub mnuaddsup_Click()
frmAddSupplier.Show vbModal

End Sub

Private Sub mnuAdjust_Click()
frmAdjustments.Show vbModal

End Sub

Private Sub mnuassDiscCodetoStock_Click()
frmStockDiscCodes.Show vbModal

End Sub

Private Sub mnucreatecostcode_Click()
frmCostCode.Show vbModal

End Sub

Private Sub mnucreatecostcodes_Click()
Dim r As New Recordset
Dim r10 As New Recordset
With r
.Open "select * from stockpurchasehistory order by stockcodemain, datepurchased", c, adOpenDynamic, adLockOptimistic
End With
With r10
.Open "select * from costcode", c, adOpenDynamic, adLockOptimistic
End With

Do While r.EOF = False
    lcp = Len(r!costofiteminc)
    cc = ""
    For ii = 1 To lcp
        
        If Mid(r!costofiteminc, ii, 1) = "1" Then
            cc = cc & r10!one
        End If
        If Mid(r!costofiteminc, ii, 1) = "2" Then
            cc = cc & r10!two
        End If
        If Mid(r!costofiteminc, ii, 1) = "3" Then
            cc = cc & r10!three
        End If
        If Mid(r!costofiteminc, ii, 1) = "4" Then
            cc = cc & r10!four
        End If
        If Mid(r!costofiteminc, ii, 1) = "5" Then
            cc = cc & r10!five
        End If
        If Mid(r!costofiteminc, ii, 1) = "6" Then
            cc = cc & r10!six
        End If
        If Mid(r!costofiteminc, ii, 1) = "7" Then
            cc = cc & r10!seven
        End If
        If Mid(r!costofiteminc, ii, 1) = "8" Then
            cc = cc & r10!eight
        End If
        If Mid(r!costofiteminc, ii, 1) = "9" Then
            cc = cc & r10!nine
        End If
        If Mid(r!costofiteminc, ii, 1) = "0" Then
            cc = cc & r10!zero
        End If
        If Mid(r!costofiteminc, ii, 1) = "." Then
            cc = cc & "/"
        End If
    Next ii
    
    
    Open "c:\LABELS\" & r!stockcodeMAIN & "\costprice.TXT" For Output As #2
    Print #2, cc
    Close #2

r.MoveNext
Loop
r.Close
MsgBox "Cost codes successfully created!", vbExclamation, "Status"

End Sub

Private Sub mnucreatedep_Click()
frmCreateDep.Show vbModal

End Sub

Private Sub mnucreatedisc_Click()
frmDiscountCode.Show vbModal

End Sub

Private Sub mnucreatenewuser_Click()
frmAddNewUser.Show
End Sub

Private Sub mnucrowdanal_Click()
frmCrowdRepSel.Show vbModal

End Sub

Private Sub mnucreatepackexist_Click()
frmCreatePackExisting.Show vbModal

End Sub

Private Sub mnucreatepacknewstock_Click()
frmCreatePackExisting2.Show vbModal

End Sub

Private Sub mnucreditnotes_Click()
frmCredCredNote.Show vbModal

End Sub

Private Sub mnueditsup_Click()
frmEditSupplier.Show vbModal

End Sub

Private Sub mnuedituser_Click()
frmaccedit.Show
End Sub

Private Sub mnuexit_Click()
End

End Sub

Private Sub mnuheadcount_Click()
frmHeadCount.Show vbModal

End Sub

Private Sub mnugrv_Click()
frmGRV.Show vbModal

End Sub

Private Sub mnuhead_Click()
frmTillSlipSetup.Show vbModal

End Sub

Private Sub mnulogoff_Click()
Unload Me
frmlogon.Show

End Sub

Private Sub mnumonsalesatTill_Click()
frmMonitorsales.Show vbModal

End Sub

Private Sub mnusales_Click()
frmSalesRepSel.Show vbModal

End Sub

Private Sub mnusettickprice_Click()
frmTicketPrice.Show vbModal

End Sub

Private Sub mnusetupcard_Click()
frmReaderSetup.Show vbModal

End Sub

Private Sub mnuviewcardtrans_Click()
frmCardTransactions.Show vbModal

End Sub

Private Sub mnuviewcompstatus_Click()
frmCompStatus.Show vbModal

End Sub

Private Sub mnuviewexh_Click()
frmViewExhibitor.Show vbModal

End Sub

Private Sub mnuopendrawer_Click()
frmDrawerCode.Show vbModal

End Sub

Private Sub mnupaycred_Click()
frmCredPay.Show vbModal

End Sub

Private Sub mnupayments_Click()
frmRepPayments.Show vbModal

End Sub

Private Sub mnupole_Click()
frmPoleDisplay.Show vbModal

End Sub

Private Sub mnuprintlabels_Click()
frmLabelPrint.Show vbModal

End Sub

Private Sub mnuprodsalessummary_Click()
frmRepSalesSummary.Show vbModal

End Sub

Private Sub mnureturns_Click()
frmRepReturns.Show vbModal

End Sub

Private Sub mnusalesperformers_Click()
frmRepPerformance.Show vbModal

End Sub

Private Sub mnusalesrep_Click()
frmRepSales.Show vbModal

End Sub

Private Sub mnusalesvoucher_Click()
frmTillSlipShow.Show vbModal

End Sub

Private Sub mnuserialtracking_Click()
frmSerialTrack.Show vbModal
End Sub

Private Sub mnusetstocktax_Click()
frmStocktaxCode.Show vbModal

End Sub

Private Sub mnusetupcompany_Click()
frmCompSetup.Show vbModal

End Sub

Private Sub mnusetupmarkup_Click()
frmMarkUp.Show vbModal

End Sub

Private Sub mnusplitpack_Click()
frmPackSplit.Show vbModal

End Sub

Private Sub mnustoconhand_Click()
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
pbody = pbody + "  <p><b><font face=Times New Roman, Times, serif><u>Stock on "

pbody = pbody + "Hand</u></font></b></p>"


pbody = pbody + "  <p align=left><b><font face=Times New Roman, Times, serif size=2>Date "
pbody = pbody + "    : " & Format(Date, "DD/MM/YYYY") & "</font></b></p>"
pbody = pbody + "  <p align=left><b><font face=Times New Roman, Times, serif size=2>Time "
pbody = pbody + "    : " & Time & "</font></b></p>"

With r7
.Open "select * from department order by department", c, adOpenDynamic, adLockOptimistic
End With
Do While r7.EOF = False
    pbody = pbody + "  <p align=left><b><font face=Times New Roman, Times, serif size=2>"
    pbody = pbody + "Department : " & r7!department & "</font></b></p>"

pbody = pbody + "  <table width=75% border=1 cellspacing=1 cellpadding=0>"
pbody = pbody + "    <tr bgcolor=#000000> "

pbody = pbody + "      <td width=20%> "
pbody = pbody + "        <div align=left><b><font color=#FFFFFF>Stock Code</font></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=60%> "
pbody = pbody + "        <div align=left><b><b><font size=2><font "
pbody = pbody + "color=#FFFFFF>Description</font></font></b></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "      <td width=20%> "
pbody = pbody + "        <div align=center><b><b><font size=2><font "
pbody = pbody + "color=#FFFFFF>On Hand</font></font></b></b></div>"
pbody = pbody + "      </td>"

pbody = pbody + "    </tr>"
With r
.Open "select * from stock where department='" & r7!department & "' order by stockcodemain", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
        pbody = pbody + "    <tr> "
        
        pbody = pbody + "      <td width=20%> "
        pbody = pbody + "        <div align=left><font face=Arial, Helvetica, sans-serif><font face=Arial, "
        pbody = pbody + "Helvetica, sans-serif><font face=Arial, Helvetica, sans-serif><font "
        pbody = pbody + "size=2>" & r!stockcodeMAIN & "</font></font></font></font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=60%> "
        pbody = pbody + "        <div align=left><font face=Arial, Helvetica, sans-serif><font face=Arial, "
        pbody = pbody + "Helvetica, sans-serif><font face=Arial, Helvetica, sans-serif><font "
        pbody = pbody + "size=2>" & UCase(r!stockdesc) & "</font></font></font></font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "      <td width=20%>"
        pbody = pbody + "        <div align=center><font face=Arial, Helvetica, sans-serif "
        pbody = pbody + "size=2>" & r!QTY & "</font></div>"
        pbody = pbody + "      </td>"
        
        pbody = pbody + "    </tr>"
        r.MoveNext
Loop
r.Close



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

Private Sub mnusuper_Click()
frmSupervisor.Show vbModal

End Sub

Private Sub mnutaxcode_Click()
frmTaxCodes.Show vbModal

End Sub

Private Sub mnuviewbalinv_Click()
frmCredViewBalInv.Show vbModal

End Sub

Private Sub mnuviewcredacc_Click()
frmCredSel.Show vbModal

End Sub

Private Sub mnuviewLoginLogsheet_Click()
frmLoginLogSheet.Show
End Sub

Private Sub mnuviewtotsales_Click()
frmSalesTOTAL.Show vbModal

End Sub

Private Sub mnuviewreturns_Click()
frmViewReturn.Show vbModal
End Sub

Private Sub mnuviewStockDetails_Click()
frmViewStock.Show vbModal

End Sub

Private Sub mnuviewsupplier_Click()
frmViewSupplier.Show vbModal

End Sub

Private Sub mnuviewtrans_Click()
frmTest.Show

End Sub

Private Sub Timer1_Timer()
frmMain.Caption = "Carousel Backoffice - " & CompName

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        mnuexit_Click
    Case 2
        mnulogoff_Click
    Case 3
     mnuviewsupplier_Click
    Case 4
        mnuviewStockDetails_Click
    Case 5
       mnuaddStock_Click
    Case 6
       mnusalesrep_Click
    Case 7
        mnusalesperformers_Click
    Case 8
        mnuserialtracking_Click
    Case 9
       mnustoconhand_Click
    Case 10
        mnuprodsalessummary_Click
    Case 11
        mnureturns_Click
        
End Select

End Sub
