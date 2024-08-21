VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTest 
   BackColor       =   &H80000016&
   Caption         =   "Transaction Log Sheet"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmTransactionlogsheet.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lvw1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   -2147483626
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Username"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date of Event"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Time of Event"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Event"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Result"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "Date Criteria"
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
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5055
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000016&
         Caption         =   "Today only"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dt 
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   1200
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   393216
         Format          =   55246849
         CurrentDate     =   38159
      End
      Begin MSComCtl2.DTPicker dp 
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   393216
         Format          =   55246849
         CurrentDate     =   38159
      End
      Begin LVbuttons.LaVolpeButton Command1 
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&View Logsheet"
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
         FCOLO           =   33023
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmTransactionlogsheet.frx":0442
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
      Begin LVbuttons.LaVolpeButton Command2 
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
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
         COLTYPE         =   2
         BCOL            =   14872561
         FCOL            =   0
         FCOLO           =   33023
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmTransactionlogsheet.frx":045E
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Start date:"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Start date:"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
    dp.Enabled = False
    dt.Enabled = False
Else
    dp.Enabled = True
    dt.Enabled = True
End If

End Sub

Private Sub Command1_Click()
Dim r As New Recordset
Dim li As ListItem
Dim d As String
lvw1.ListItems.Clear
If Check1.Value = 1 Then
    With r
    .Open "select * from transactionlogsheet where dateofevent='" & Year(Date) & "-" & Month(Date) & "-" & Day(Date) & " 00:00:00.000'", c, adOpenDynamic, adLockOptimistic
    End With
Else

    With r
    .Open "select * from transactionlogsheet where dateofevent>='" & Year(dp.Value) & "-" & Month(dp.Value) & "-" & Day(dp.Value) & " 00:00:00.000' and dateofevent<='" & Year(dt.Value) & "-" & Month(dt.Value) & "-" & Day(dt.Value) & " 00:00:00.000'", c, adOpenDynamic, adLockOptimistic
    End With
    
End If

Do While r.EOF = False
    Set li = lvw1.ListItems.Add(, , r!username)
    With li
    .SubItems(1) = r!dateofevent
    .SubItems(2) = r!timeofevent
    .SubItems(3) = r!event
    .SubItems(4) = r!Result
    '.SubItems(5) = r!ObjectActedOn
    End With
    
    r.MoveNext
Loop
r.Close

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

