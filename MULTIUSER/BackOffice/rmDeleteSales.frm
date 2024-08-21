VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rmDeleteSales 
   Caption         =   "Delete Sales"
   ClientHeight    =   1725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
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
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4455
      Begin MSComCtl2.DTPicker dFROM 
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   97583105
         CurrentDate     =   38483
      End
      Begin MSComCtl2.DTPicker dTO 
         Height          =   255
         Left            =   3000
         TabIndex        =   3
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   97583105
         CurrentDate     =   38483
      End
      Begin VB.Label Label1 
         Caption         =   "From:"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "To:"
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   480
         Width           =   255
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Delete"
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
      MICON           =   "rmDeleteSales.frx":0000
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
      Left            =   1320
      TabIndex        =   6
      Top             =   1200
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
      MICON           =   "rmDeleteSales.frx":001C
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
Attribute VB_Name = "rmDeleteSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
dFROM.Value = Format(Date, "DD/MM/YYYY")
dTO.Value = Format(Date, "DD/MM/YYYY")
End Sub


       
 

Private Sub LaVolpeButton1_Click()
delete = MsgBox("Are you sure you wish delete sales?", vbYesNo, "Delete")
If delete = vbYes Then
Dim r As New Recordset

 With r
        .Open "delete from sales where saledate>='" & Year(dFROM.Value) & "-" & Month(dFROM.Value) & "-" & Day(dFROM.Value) & " 00:00:00.000' and saledate<='" & Year(dTO.Value) & "-" & Month(dTO.Value) & "-" & Day(dTO.Value) & " 00:00:00.000' ", c, adOpenDynamic, adLockOptimistic
        End With
        MsgBox " Sales Deleted Successfully!", vbInformation, "Sales Deleted"
        Else
        Exit Sub
        End If
        
End Sub

Private Sub LaVolpeButton2_Click()

Unload Me

End Sub

