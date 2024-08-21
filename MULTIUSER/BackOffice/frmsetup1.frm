VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmsetup1 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Connection Setup"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   ControlBox      =   0   'False
   Icon            =   "frmsetup1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton Command1 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   3240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Ok"
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
      MICON           =   "frmsetup1.frx":0442
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
   Begin VB.TextBox txtSQLUserPwd 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2520
      Width           =   2550
   End
   Begin VB.TextBox txtSqlSN 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1905
      TabIndex        =   2
      Top             =   1470
      Width           =   2550
   End
   Begin VB.TextBox txtSQLUID 
      Appearance      =   0  'Flat
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1905
      TabIndex        =   3
      Top             =   2025
      Width           =   2550
   End
   Begin VB.TextBox txtShare 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1875
      TabIndex        =   1
      Top             =   915
      Width           =   2550
   End
   Begin VB.TextBox txtsn 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1875
      TabIndex        =   0
      Top             =   360
      Width           =   2550
   End
   Begin LVbuttons.LaVolpeButton command2 
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   3240
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
      FCOLO           =   33023
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmsetup1.frx":045E
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
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "SQL UserPwd"
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
      Height          =   405
      Left            =   720
      TabIndex        =   16
      Top             =   2535
      Width           =   1710
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   " eg. ABC"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   4440
      TabIndex        =   15
      Top             =   2580
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "SQL Server Name"
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
      Height          =   405
      Left            =   240
      TabIndex        =   14
      Top             =   1440
      Width           =   1710
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "SQL UserID"
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
      Height          =   405
      Left            =   765
      TabIndex        =   13
      Top             =   2040
      Width           =   1710
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "eg. XYZ"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   4470
      TabIndex        =   12
      Top             =   2085
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "eg. MyServer"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   4485
      TabIndex        =   11
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "eg. Computer1"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   4455
      TabIndex        =   10
      Top             =   450
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "eg. C"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   4440
      TabIndex        =   9
      Top             =   975
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Share Name"
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
      Height          =   405
      Left            =   735
      TabIndex        =   8
      Top             =   930
      Width           =   1710
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "IP Address/Server Name"
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
      Height          =   405
      Left            =   210
      TabIndex        =   7
      Top             =   330
      Width           =   1710
   End
End
Attribute VB_Name = "frmsetup1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim cs As New Connection
Dim r As New Recordset
On Error GoTo NoConnection

If Len(txtsn) > 0 And Len(txtShare) > 0 And Len(txtSqlSN) > 0 And Len(txtSQLUID) > 0 And Len(txtSQLUserPwd) > 0 Then
Me.MousePointer = 11
With cs
.ConnectionString = App.Path & "/Temp.mdb"
.Provider = "microsoft.jet.oledb.4.0"
.Open
End With

With r
.Open "server", cs, adOpenDynamic, adLockOptimistic
End With
r!servername = txtsn
r!sharename = txtShare
r!SqlServerName = txtSqlSN
r!SQLUserid = txtSQLUID
r!SQLUserpwd = txtSQLUserPwd

servername = txtsn
sharename = txtShare
SqlServerName = txtSqlSN
SQLUserid = txtSQLUID
SQLUserpwd = txtSQLUserPwd


r.Update
r.Close
cs.Close
'Set s = CreateObject("scripting.filesystemobject")


's.copyfile "\\" & servername & "\" & sharename & "\TestServer.txt", App.Path & "\TestServer.txt"
'Kill App.Path & "\TestServer.txt"


ConnectMe SqlServerName, SQLUserid, SQLUserpwd
'ConnectMe2 SqlServerName, SQLUserid, SQLUserpwd

Unload Me

frmlogon.Show





Else
MsgBox "Please enter a server name as well as share name!", vbExclamation, "error!"
End If
Me.MousePointer = 0

Exit Sub
NoConnection:
Me.MousePointer = 0
If Err.Number = 76 Then
    MsgBox "Share name " & sharename & " within Server " & servername & " does not exist!", vbCritical, "Cannot connect to Server!"
    txtsn = ""
    txtShare = ""
    txtsn.SetFocus
Else
    MsgBox Err.Description, vbCritical, "Cannot connect to SqlServer!"
End If

End Sub

Private Sub Command2_Click()
End

End Sub

Private Sub Form_Load()
Dim r As New Recordset
Dim cs As New Connection
With cs
.ConnectionString = App.Path & "\Temp.mdb"
.Provider = "microsoft.jet.oledb.4.0"
.Open
End With
With r
.Open "Server", cs, adOpenDynamic, adLockOptimistic
End With
If r.EOF = False Then
txtsn = r!servername
txtShare = r!sharename
txtSqlSN = r!SqlServerName
txtSQLUID = r!SQLUserid
txtSQLUserPwd = r!SQLUserpwd

r.Close
Else
r.Close
End If
cs.Close


End Sub

Private Sub LaVolpeButton1_Click()

End Sub
