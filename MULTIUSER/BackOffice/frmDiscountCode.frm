VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDiscountCode 
   BackColor       =   &H80000016&
   Caption         =   "Discount Codes"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H80000016&
      Caption         =   "Available Discount Codes"
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
      Height          =   2415
      Left            =   240
      TabIndex        =   11
      Top             =   3720
      Width           =   4455
      Begin MSComctlLib.ListView lvw1 
         Height          =   1935
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Discount(%)"
            Object.Width           =   1790
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Effective From"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Effective To"
            Object.Width           =   2293
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "New Discount Code"
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
      Height          =   3135
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   4455
      Begin MSComCtl2.DTPicker dFrom 
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   55246849
         CurrentDate     =   38460
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton1 
         Height          =   255
         Left            =   3240
         TabIndex        =   4
         Top             =   2640
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "&Add to List"
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
         BCOL            =   0
         FCOL            =   16777215
         FCOLO           =   33023
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmDiscountCode.frx":0000
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
      Begin VB.TextBox txtdisc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   1
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtcode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         Top             =   600
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   255
         Left            =   1320
         TabIndex        =   3
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   55246849
         CurrentDate     =   38460
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton2 
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   2640
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
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
         BCOL            =   0
         FCOL            =   16777215
         FCOLO           =   33023
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmDiscountCode.frx":001C
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
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valid to:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valid from:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Discount(%):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Code:"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmDiscountCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim li As ListItem
Dim r As New Recordset
dFrom.Value = Format(Date, "DD/MM/YYYY")
dTo.Value = Format(Date, "DD/MM/YYYY")
With r
.Open "Select * from DiscCodes order by DiscCode", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    Set li = lvw1.ListItems.Add(, , r!disccode)
    With li
    .SubItems(1) = r!discount
    .SubItems(2) = Format(r!Fromdate, "DD/MM/YYYY")
    .SubItems(3) = Format(r!todate, "DD/MM/YYYY")
    End With
    r.MoveNext
Loop
r.Close

End Sub

Private Sub LaVolpeButton1_Click()
Dim r As New Recordset
If Len(txtcode) = 0 Then
    MsgBox "Please enter Discount Code!", vbExclamation, "Status"
    txtcode.SetFocus
    Exit Sub
End If
If txtcode.Text = "0" Then
    MsgBox "Cannot use '0' as Discount code as '0' is the default value! Please enter another Discount Code.", vbExclamation, "Status"
    txtcode.Text = ""
    txtcode.SetFocus
    Exit Sub
End If
If Len(txtdisc) = 0 Then
    MsgBox "Please enter Discount!", vbExclamation, "Status"
    txtdisc.SetFocus
    Exit Sub
End If

With r
.Open "select * from disccodes where disccode='" & txtcode.Text & "'", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = False Then
    res = MsgBox("Cannot create code! Code already exist. Do you wish to overwrite code?", vbYesNo + vbQuestion, "Status")
    If res = vbNo Then
        txtcode.Text = ""
        txtcode.SetFocus
        Exit Sub
    Else
        r!disccode = UCase(txtcode.Text)
        r!discount = txtdisc
        r!Fromdate = Format(dFrom.Value, "DD/MM/YYYY")
        r!todate = Format(dTo.Value, "DD/MM/YYYY")
        r.Update
        r.Close
        With r
        .Open "select * from stockdiscount where disccode='" & txtcode.Text & "'", c, adOpenDynamic, adLockOptimistic
        End With
        Do While r.EOF = False
            r!disc = txtdisc
            r!Fromdate = Format(dFrom.Value, "DD/MM/YYYY")
            r!todate = Format(dTo.Value, "DD/MM/YYYY")
            r.Update
            r.MoveNext
        Loop
        r.Close
        
        dFrom.Value = Format(Date, "DD/MM/YYYY")
        dTo.Value = Format(Date, "DD/MM/YYYY")
        txtcode.Text = ""
        txtdisc.Text = ""
        lvw1.ListItems.Clear
        With r
        .Open "Select * from DiscCodes order by DiscCode", c, adOpenDynamic, adLockOptimistic
        End With
        Do While r.EOF = False
            Set li = lvw1.ListItems.Add(, , r!disccode)
            With li
            .SubItems(1) = r!discount
            .SubItems(2) = Format(r!Fromdate, "DD/MM/YYYY")
            .SubItems(3) = Format(r!todate, "DD/MM/YYYY")
            End With
            r.MoveNext
        Loop
        r.Close
        MsgBox "Discount Code successfully created!", vbInformation, "Status"
        txtcode.SetFocus
        Exit Sub
    End If
End If
r.Close
With r
.Open "select * from disccodes", c, adOpenDynamic, adLockOptimistic
End With
    r.AddNew
    r!disccode = UCase(txtcode.Text)
    r!discount = txtdisc
    r!Fromdate = Format(dFrom.Value, "DD/MM/YYYY")
    r!todate = Format(dTo.Value, "DD/MM/YYYY")
    r.Update
    r.Close
    dFrom.Value = Format(Date, "DD/MM/YYYY")
    dTo.Value = Format(Date, "DD/MM/YYYY")
    txtcode.Text = ""
    txtdisc.Text = ""
    lvw1.ListItems.Clear
    With r
    .Open "Select * from DiscCodes order by DiscCode", c, adOpenDynamic, adLockOptimistic
    End With
    Do While r.EOF = False
        Set li = lvw1.ListItems.Add(, , r!disccode)
        With li
        .SubItems(1) = r!discount
        .SubItems(2) = Format(r!Fromdate, "DD/MM/YYYY")
        .SubItems(3) = Format(r!todate, "DD/MM/YYYY")
        End With
        r.MoveNext
    Loop
    r.Close
    MsgBox "Discount Code successfully created!", vbInformation, "Status"
        
    txtcode.SetFocus









End Sub

Private Sub LaVolpeButton2_Click()
Unload Me

End Sub

Private Sub txtdisc_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    A = InStr(1, txtdisc, ".")
    If A > 0 Then
        KeyAscii = 0
    End If
End If

If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Or KeyAscii = 46 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If

End Sub
