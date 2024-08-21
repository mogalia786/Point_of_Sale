VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmStockEdit 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change stock details"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   2
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
      COLTYPE         =   2
      BCOL            =   14872561
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmStockEdit.frx":0000
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
   Begin VB.Frame Frame4 
      BackColor       =   &H80000016&
      Caption         =   "Stock Details"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.ComboBox cDEP 
         Height          =   315
         ItemData        =   "frmStockEdit.frx":001C
         Left            =   1440
         List            =   "frmStockEdit.frx":001E
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox cTaxCode 
         Height          =   315
         ItemData        =   "frmStockEdit.frx":0020
         Left            =   1440
         List            =   "frmStockEdit.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtSP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtdesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Department:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Code:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Selling Price:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lSC 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Code:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1320
         Width           =   1095
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   2
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
      MICON           =   "frmStockEdit.frx":0024
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
Attribute VB_Name = "frmStockEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim r As New Recordset
With r
.Open "select * from taxcode order by taxcode", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    cTaxCode.AddItem r!taxcode
    r.MoveNext
Loop
r.Close
With r
.Open "select * from department order by department", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    cDEP.AddItem r!department
    r.MoveNext
Loop
r.Close
cDEP.Text = frmViewStock.lDEP.Caption

lSC = frmViewStock.lSC
cTaxCode.Text = frmViewStock.lTAX
txtdesc.Text = frmViewStock.ldesc
txtSP.Text = frmViewStock.lsp

End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

End Sub

Private Sub LaVolpeButton1_Click()
Dim r As New Recordset
If Len(txtdesc.Text) = 0 Then
    MsgBox "Please enter stock description!", vbExclamation, "Status"
    txtdesc.SetFocus
    Exit Sub
End If
If Len(cTaxCode.Text) = 0 Then
    MsgBox "Please select taxcode!", vbExclamation, "Status"
    cTaxCode.SetFocus
    Exit Sub
End If
If Len(cDEP.Text) = 0 Then
    MsgBox "Please select department!", vbExclamation, "Status"
    cDEP.SetFocus
    Exit Sub
End If
If Len(txtSP.Text) = 0 Then
    MsgBox "Please enter selling price!", vbExclamation, "Status"
    txtSP.SetFocus
    Exit Sub
End If

res = MsgBox("Are you sure you wish to update details?", vbYesNo + vbQuestion, "Update?")
If res = vbNo Then
    frmViewStock.lvw1_Click
    lSC = frmViewStock.lSC
    cTaxCode.Text = frmViewStock.lTAX
    txtdesc.Text = frmViewStock.ldesc
    txtSP.Text = frmViewStock.lsp
    Exit Sub
End If


With r
.Open "select * from stock where stockcodemain='" & lSC.Caption & "'", c, adOpenDynamic, adLockOptimistic
End With
r!stockdesc = UCase(txtdesc.Text)
r!taxcode = cTaxCode.Text
r!department = cDEP.Text
r!unitprice = Format(txtSP.Text, "#####.00")
r.Update
r.Close
With r
.Open "select * from sales where itemcodemain='" & lSC.Caption & "'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
r!department = cDEP.Text
r.Update
r.MoveNext
Loop
r.Close
With r
.Open "select * from returns where stockcode='" & lSC.Caption & "'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
r!department = cDEP.Text
r.Update
r.MoveNext
Loop
r.Close





With r
.Open "select * from stockpurchasehistory where stockcodemain='" & lSC.Caption & "'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
r!department = cDEP.Text
r!sellingpriceINc = Format(txtSP.Text, "#####.00")
r.Update
r.MoveNext
Loop
r.Close
MsgBox "Record successfully updated!", vbInformation, "Status"
frmViewStock.lvw1_Click
Unload Me


End Sub

Private Sub LaVolpeButton2_Click()
Unload Me

End Sub

Private Sub txtSP_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    A = InStr(1, txtSP, ".")
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
