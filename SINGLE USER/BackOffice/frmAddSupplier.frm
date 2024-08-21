VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmAddSupplier 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Supplier"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   5040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Add Supplier"
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
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmAddSupplier.frx":0000
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
      Caption         =   "Supplier Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   4815
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   4815
      Begin VB.TextBox lBAL 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox tSUP 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   0
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox tADD1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox tADD2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox tADD3 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox tADD4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox tTEL 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox tFAX 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   3240
         Width           =   1935
      End
      Begin VB.ComboBox cTerms 
         Height          =   315
         ItemData        =   "frmAddSupplier.frx":001C
         Left            =   1800
         List            =   "frmAddSupplier.frx":006B
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Balance:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   20
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address1:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address2:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address3:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Address4:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tel:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Terms:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   4200
         Width           =   1575
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   5040
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
      COLTYPE         =   1
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmAddSupplier.frx":00E6
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
Attribute VB_Name = "frmAddSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LaVolpeButton1_Click()
Dim r As New Recordset
Dim r1 As New Recordset
Dim r2 As New Recordset
If Len(tSUP) = 0 Then
    MsgBox "Please enter Supplier!", vbInformation, "Status"
    tSUP.SetFocus
    Exit Sub
End If
If Len(tADD1) = 0 Then
    tADD1.Text = "N/A"
End If
If Len(tADD2) = 0 Then
    tADD2.Text = "N/A"
End If
If Len(tADD3) = 0 Then
    tADD3 = "N/A"
End If
If Len(tADD4) = 0 Then
    tADD4.Text = "N/A"
End If
If Len(tTEL) = 0 Then
    tTEL.Text = "N/A"
End If
If Len(tFAX) = 0 Then
    tFAX.Text = "N/A"
End If
If Len(cTerms.Text) = 0 Then
    MsgBox "Please select terms!", vbInformation, "Status"
    cTerms.SetFocus
    Exit Sub
End If
If Len(lBAL) = 0 Then
    lBAL = "0.00"
End If

With r
.Open "select * from supplier where supplier='" & tSUP & "'", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = True Then
    With r1
    .Open "select * from suppliercode", c, adOpenDynamic, adLockOptimistic
    End With
    With r2
    .Open "select * from ledgeraccount", c, adOpenDynamic, adLockOptimistic
    End With
    r2.AddNew
    r2!account = UCase(tSUP)
    r2!accountno = "C" & CDbl(r1!suppliercode) + 1
    r2.Update
    r2.Close
    r1!suppliercode = CDbl(r1!suppliercode) + 1
    r1.Update
    r1.Close
    
    r.AddNew
    r!supplier = UCase(tSUP)
    r!add1 = tADD1
    r!add2 = tADD2
    r!add3 = tADD3
    r!add4 = tADD4
    r!tel = tTEL
    r!fax = tFAX
    r!terms = cTerms.Text
    r!OpeningBalance = Format(lBAL, "#####.00")
    r.Update
Else
    r!supplier = UCase(tSUP)
    r!add1 = tADD1
    r!add2 = tADD2
    r!add3 = tADD3
    r!add4 = tADD4
    r!tel = tTEL
    r!fax = tFAX
    r!terms = cTerms.Text
    r!OpeningBalance = Format(lBAL, "#####.00")
    r.Update
End If
r.Close
res = MsgBox("Supplier successfully added! Do you wish to add another?", vbYesNo + vbQuestion, "Add Another?")
If res = vbYes Then
    tSUP = ""
    tADD1.Text = ""
    tADD2 = ""
    tADD3 = ""
    tADD4 = ""
    tTEL = ""
    tFAX = ""
    cTerms.ListIndex = -1
    lBAL = ""
    tSUP.SetFocus
Else
    Unload Me
End If

End Sub

Private Sub LaVolpeButton2_Click()
Unload Me

End Sub

Private Sub lBAL_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    A = InStr(1, lBAL, ".")
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
