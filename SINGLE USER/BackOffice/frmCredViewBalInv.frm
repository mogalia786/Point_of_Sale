VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmCredViewBalInv 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Balance owed on invoice"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   3360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BackColor       =   &H80000016&
      Caption         =   "Balance due on Invoice"
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
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   3135
      Begin VB.Label lBAL 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF80FF&
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000016&
      Caption         =   "Invoice"
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
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3135
      Begin VB.ComboBox cINV 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "Select Creditor"
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
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   3135
      Begin VB.ComboBox cCred 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2775
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   3120
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "frmCredViewBalInv.frx":0000
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
Attribute VB_Name = "frmCredViewBalInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cCred_Click()
Dim r As New Recordset
cINV.Clear
With r
.Open "select * from creditorsinvoice where creditor='" & cCred.Text & "' order by invno", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    cINV.AddItem r!invno
    r.MoveNext
Loop
r.Close

    
End Sub

Private Sub cINV_Click()
Dim r As New Recordset
Dim TBALA As Double
lBAL = ""
With r
.Open "select * from creditorsinvoice where invno='" & cINV.Text & "' and creditor='" & cCred.Text & "'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    TBALA = TBALA + r!tendered
    r.MoveNext
Loop
r.Close
With r
.Open "select * from creditorspayment where invno='" & cINV.Text & "' and creditor='" & cCred.Text & "'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    TBALA = TBALA - r!amount
    r.MoveNext
Loop
r.Close

With r
.Open "select * from creditorscreditnote where invno='" & cINV.Text & "' and creditor='" & cCred.Text & "'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    TBALA = TBALA - r!amount
    r.MoveNext
Loop
r.Close

lBAL = "R" & Format(TBALA, "#####.00")

End Sub

Private Sub Form_Load()
Dim r As New Recordset
With r
.Open "select * from supplier order by supplier", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    cCred.AddItem r!supplier
    r.MoveNext
Loop
r.Close

End Sub

Private Sub LaVolpeButton1_Click()
Dim r As New Recordset
If Len(cCred.Text) = 0 Then
    MsgBox "Please select creditor!", vbExclamation, "Status"
    cCred.SetFocus
    Exit Sub
End If
If Len(cINV.Text) = 0 Then
    MsgBox "Please select invoice number!", vbExclamation, "Status"
    cINV.SetFocus
    Exit Sub
End If
If Len(lBAL) = 0 Then
    MsgBox "Please select invoice!", vbExclamation, "status"
    cINV.SetFocus
    Exit Sub
End If
If Len(txtCHQ.Text) = 0 Then
    txtCHQ = "N/A"
End If
If Len(txtAMT.Text) = 0 Then
    MsgBox "Please enter amount paid!", vbExclamation, "Status"
    txtAMT.SetFocus
    Exit Sub
End If
If CDbl(txtAMT) > CDbl(lBAL) Then
    MsgBox "Amount tendered cannot be higher than amount owed!", vbInformation, "Status"
    txtAMT.Text = ""
    txtAMT.SetFocus
    Exit Sub
End If
With r
.Open "select * from creditorspayment", c, adOpenDynamic, adLockOptimistic
End With
r.AddNew
r!paymentdate = Format(pDate.Value, "DD/MM/YYYY")
r!creditor = UCase(cCred.Text)
r!chqno = txtCHQ.Text
r!invno = cINV.Text
r!amount = Format(txtAMT.Text, "#####.00")
r.Update
r.Close
res = MsgBox("Payment updated successfully! Do you wish to make another payment>", vbYesNo + vbQuestion, "Make another payment?")
If res = vbYes Then
    cCred.ListIndex = -1
    cINV.ListIndex = -1
    lBAL = ""
    txtCHQ = ""
    txtAMT.Text = ""
    cCred.SetFocus
Else
    Unload Me
End If

End Sub

Private Sub LaVolpeButton2_Click()
Unload Me

End Sub

Private Sub txtAMT_LostFocus()
If Len(txtAMT.Text) > 0 Then
    If Len(lBAL) = 0 Then
        MsgBox "Please select Invoice Number!", vbInformation, "Status"
        cINV.SetFocus
        Exit Sub
    End If
    If CDbl(txtAMT) > CDbl(lBAL) Then
        MsgBox "Amount tendered cannot be higher than amount owed!", vbInformation, "Status"
        txtAMT.Text = ""
        txtAMT.SetFocus
        Exit Sub
    End If
End If

End Sub
