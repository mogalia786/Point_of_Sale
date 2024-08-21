VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmPrinterSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printer Setup"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4110
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Update"
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
      MICON           =   "frmPrinterSetup.frx":0000
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
   Begin VB.Frame f2 
      Caption         =   "Port Settings"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
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
      TabIndex        =   3
      Top             =   1440
      Width           =   3495
      Begin VB.ComboBox cPortno 
         Height          =   315
         ItemData        =   "frmPrinterSetup.frx":001C
         Left            =   1080
         List            =   "frmPrinterSetup.frx":0032
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Printer Connection"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3495
      Begin VB.OptionButton optSer 
         Caption         =   "Serial (RS232)"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton optPar 
         Caption         =   "Parallel"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cancel / E&xit"
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
      MICON           =   "frmPrinterSetup.frx":0048
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
Attribute VB_Name = "frmPrinterSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If PrinterConn = "Parallel" Then
    optPar.Value = True
    f2.Enabled = False
Else
    optSer.Value = True
    f2.Enabled = True
    cPortno.Text = PortNo
End If

End Sub

Private Sub LaVolpeButton1_Click()
Dim r As New Recordset
Dim cs As New Connection
If optSer.Value = True Then
    If Len(cPortno.Text) = 0 Then
        MsgBox "Please select printer port number!", vbExclamation, "Status"
        cPortno.SetFocus
        Exit Sub
    End If
End If

With cs
.ConnectionString = App.Path & "/Temp.mdb"
.Provider = "microsoft.jet.oledb.4.0"
.Open
End With
With r
.Open "printercon", cs, adOpenDynamic, adLockOptimistic
End With
If optPar.Value = True Then
    r!printercon = "Parallel"
    r.Update
    PrinterConn = "Parallel"
Else
    r!printercon = "Serial"
    r.Update
    PrinterConn = "Serial"
End If
r.Close

With r
.Open "portno", cs, adOpenDynamic, adLockOptimistic
End With
If optSer.Value = True Then
    r!PortNo = cPortno.Text
    r.Update
    PortNo = cPortno.Text
End If
r.Close
MsgBox "Settings updated successfully!", vbInformation, "Status"
Unload Me



End Sub

Private Sub LaVolpeButton2_Click()
Unload Me

End Sub

Private Sub optPar_Click()
f2.Enabled = False

End Sub

Private Sub optSer_Click()
f2.Enabled = True
'cPortno.SetFocus
End Sub
