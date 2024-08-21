VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmViewStock 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Details"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
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
      Height          =   2775
      Left            =   120
      TabIndex        =   14
      Top             =   5040
      Width           =   7215
      Begin VB.Label Label11 
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
         Left            =   4200
         TabIndex        =   42
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label lDEP 
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
         Left            =   5400
         TabIndex        =   41
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label labyrt5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Adjustments:"
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
         Left            =   0
         TabIndex        =   40
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label lAdj 
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
         TabIndex        =   39
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Qty Packed:"
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
         Left            =   0
         TabIndex        =   38
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Qty Unpacked:"
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
         Left            =   3720
         TabIndex        =   37
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lPACKED 
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
         TabIndex        =   36
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label lUNPACKED 
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
         Left            =   5400
         TabIndex        =   35
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label lPdate 
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
         Left            =   5400
         TabIndex        =   34
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label lTOTPurch 
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
         Left            =   5400
         TabIndex        =   33
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lTAX 
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
         Left            =   5400
         TabIndex        =   32
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lSUP 
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
         Left            =   5400
         TabIndex        =   31
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Last Purchase Date:"
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
         Left            =   3720
         TabIndex        =   30
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Purchased:"
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
         Left            =   3600
         TabIndex        =   29
         Top             =   960
         Width           =   1695
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
         Left            =   4200
         TabIndex        =   28
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Last Supplier:"
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
         Left            =   4200
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lsp 
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
         Left            =   5400
         TabIndex        =   26
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label lcp 
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
         TabIndex        =   25
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label lsold 
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
         TabIndex        =   24
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label lqty 
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
         TabIndex        =   23
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label ldesc 
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
         TabIndex        =   22
         Top             =   600
         Width           =   2655
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
         TabIndex        =   21
         Top             =   240
         Width           =   2655
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
         Left            =   4200
         TabIndex        =   20
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ave. Cost Price:"
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
         Left            =   0
         TabIndex        =   19
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Qty Sold:"
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
         TabIndex        =   18
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Qty on hand:"
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
         TabIndex        =   17
         Top             =   960
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
         TabIndex        =   16
         Top             =   600
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
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000016&
      Caption         =   "Search..."
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
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   4335
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtBarcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label1 
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
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode:"
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
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000016&
      Caption         =   "Criteria (Left Click on Item to view details / Right Click to change details)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   7215
      Begin MSComctlLib.ListView lvw1 
         Height          =   1575
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Stock Code"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "Criteria"
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
      TabIndex        =   1
      Top             =   120
      Width           =   7215
      Begin VB.OptionButton OptStockCode 
         BackColor       =   &H80000016&
         Caption         =   "Stock Code"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optDesc 
         BackColor       =   &H80000016&
         Caption         =   "Item Description"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optBarcode 
         BackColor       =   &H80000016&
         Caption         =   "Barcode"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   2040
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
      MICON           =   "frmViewStock.frx":0000
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
   Begin VB.Image Image1 
      Height          =   975
      Left            =   5280
      Picture         =   "frmViewStock.frx":001C
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "frmViewStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LaVolpeButton1_Click()
Unload Me

End Sub

Public Sub lvw1_Click()
Dim r As New Recordset
Dim r1 As New Recordset
Dim r2 As New Recordset
Dim QTYSOLD As Double
Dim qtyOnhand As Double
Dim NumofPurchase As Double
Dim TotalCP As Double
Dim TotalPurchased As Double
Dim QtyPacked As Double
Dim QtyUnPacked As Double
Dim mDIFF As Double
Dim TotAdj As Double
If lvw1.ListItems.Count = 0 Then Exit Sub
If Len(lvw1.SelectedItem) = 0 Then Exit Sub
lSC = ""
ldesc = ""
lTAX = ""
lsp = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lcp = ""
lPACKED = ""
lUNPACKED = ""
lAdj = ""
lDEP = ""

lSC = lvw1.SelectedItem
With r
.Open "select * from stock where stockcodemain='" & lvw1.SelectedItem & "'", c, adOpenDynamic, adLockOptimistic
End With

lDEP = r!department

ldesc.Caption = UCase(r!stockdesc)
lTAX = r!taxcode
lsp = Format(r!unitprice, "#####.00")
r.Close
With r
.Open "select * from stock where stockcodemain='" & lvw1.SelectedItem & "'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    qtyOnhand = qtyOnhand + CDbl(r!QTY)
    r.MoveNext
Loop


lqty = qtyOnhand
r.Close

With r
.Open "select * from sales where itemcodemain='" & lvw1.SelectedItem & "'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    QTYSOLD = QTYSOLD + CDbl(r!QTY)
    r.MoveNext
Loop
r.Close
lsold = QTYSOLD
With r
.Open "select * from stockpurchasehistory where stockcodemain='" & lvw1.SelectedItem & "' order by datepurchased desc", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = False Then
    lSUP = UCase(r!supplier)
    lPdate = r!datepurchased
Else
    lSUP = "Internal"
    lPdate = "N/A"
End If

r.Close

With r
.Open "select * from stockpurchasehistory where stockcodemain='" & lvw1.SelectedItem & "' order by datepurchased desc", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    NumofPurchase = NumofPurchase + 1
    TotalCP = TotalCP + r!costofiteminc
    r.MoveNext
Loop
r.Close
lcp = Format(TotalCP / NumofPurchase, "#####.00")
With r
.Open "select * from stockpurchasehistory where stockcodemain='" & lvw1.SelectedItem & "' order by datepurchased desc", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    TotalPurchased = TotalPurchased + CDbl(r!qtypurchased)
    r.MoveNext
Loop
lTOTPurch = TotalPurchased


r.Close
TotAdj = 0
With r
.Open "select * from stockadjustment where stockcode='" & lvw1.SelectedItem & "'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    TotAdj = TotAdj + r!adjustedby
    r.MoveNext
Loop
r.Close

If CDbl(lTOTPurch) <> CDbl(lqty) + CDbl(lsold) Then
    With r
    .Open "select * from packslist where pstockcode='" & lvw1.SelectedItem & "'", c, adOpenDynamic, adLockOptimistic
    End With
    If r.EOF = False Then
        QtyUnPacked = CDbl(lTOTPurch) - CDbl(lqty) - CDbl(lsold) + TotAdj
        QtyPacked = 0
    Else
        QtyPacked = CDbl(lTOTPurch) - CDbl(lqty) - CDbl(lsold) + TotAdj
        QtyUnPacked = 0
    End If
    r.Close
    lPACKED = QtyPacked
    lUNPACKED = QtyUnPacked
Else
    lPACKED = "0"
    lUNPACKED = "0"
End If

lAdj.Caption = TotAdj

End Sub

Private Sub Option1_Click()

End Sub

Private Sub Option2_Click()

End Sub

Private Sub lvw1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim r As New Recordset
Dim r1 As New Recordset
Dim r2 As New Recordset
Dim QTYSOLD As Double
Dim qtyOnhand As Double
Dim NumofPurchase As Double
Dim TotalCP As Double
Dim TotalPurchased As Double
Dim QtyPacked As Double
Dim QtyUnPacked As Double
Dim mDIFF As Double
Dim TotAdj As Double
If lvw1.ListItems.Count = 0 Then Exit Sub
If Len(lvw1.SelectedItem) = 0 Then Exit Sub
lSC = ""
ldesc = ""
lTAX = ""
lsp = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lcp = ""
lPACKED = ""
lUNPACKED = ""
lAdj = ""

lSC = lvw1.SelectedItem
With r
.Open "select * from stock where stockcodemain='" & lvw1.SelectedItem & "'", c, adOpenDynamic, adLockOptimistic
End With


ldesc.Caption = UCase(r!stockdesc)
lTAX = r!taxcode
lsp = Format(r!unitprice, "#####.00")
r.Close
With r
.Open "select * from stock where stockcodemain='" & lvw1.SelectedItem & "'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    qtyOnhand = qtyOnhand + CDbl(r!QTY)
    r.MoveNext
Loop


lqty = qtyOnhand
r.Close

With r
.Open "select * from sales where itemcodemain='" & lvw1.SelectedItem & "'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    QTYSOLD = QTYSOLD + CDbl(r!QTY)
    r.MoveNext
Loop
r.Close
lsold = QTYSOLD
With r
.Open "select * from stockpurchasehistory where stockcodemain='" & lvw1.SelectedItem & "' order by datepurchased desc", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = False Then
    lSUP = UCase(r!supplier)
    lPdate = r!datepurchased
Else
    lSUP = "Internal"
    lPdate = "N/A"
End If

r.Close

With r
.Open "select * from stockpurchasehistory where stockcodemain='" & lvw1.SelectedItem & "' order by datepurchased desc", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    NumofPurchase = NumofPurchase + 1
    TotalCP = TotalCP + r!costofiteminc
    r.MoveNext
Loop
r.Close
lcp = Format(TotalCP / NumofPurchase, "#####.00")
With r
.Open "select * from stockpurchasehistory where stockcodemain='" & lvw1.SelectedItem & "' order by datepurchased desc", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    TotalPurchased = TotalPurchased + CDbl(r!qtypurchased)
    r.MoveNext
Loop
lTOTPurch = TotalPurchased


r.Close
TotAdj = 0
With r
.Open "select * from stockadjustment where stockcode='" & lvw1.SelectedItem & "'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    TotAdj = TotAdj + r!adjustedby
    r.MoveNext
Loop
r.Close

If CDbl(lTOTPurch) <> CDbl(lqty) + CDbl(lsold) Then
    With r
    .Open "select * from packslist where pstockcode='" & lvw1.SelectedItem & "'", c, adOpenDynamic, adLockOptimistic
    End With
    If r.EOF = False Then
        QtyUnPacked = CDbl(lTOTPurch) - CDbl(lqty) - CDbl(lsold) + TotAdj
        QtyPacked = 0
    Else
        QtyPacked = CDbl(lTOTPurch) - CDbl(lqty) - CDbl(lsold) + TotAdj
        QtyUnPacked = 0
    End If
    r.Close
    lPACKED = QtyPacked
    lUNPACKED = QtyUnPacked
Else
    lPACKED = "0"
    lUNPACKED = "0"
End If

lAdj.Caption = TotAdj

End Sub

Private Sub lvw1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim r As New Recordset
Dim r1 As New Recordset
Dim r2 As New Recordset
Dim QTYSOLD As Double
Dim qtyOnhand As Double
Dim NumofPurchase As Double
Dim TotalCP As Double
Dim TotalPurchased As Double
Dim QtyPacked As Double
Dim QtyUnPacked As Double
Dim mDIFF As Double
Dim TotAdj As Double
If lvw1.ListItems.Count = 0 Then Exit Sub
If Len(lvw1.SelectedItem) = 0 Then Exit Sub
lSC = ""
ldesc = ""
lTAX = ""
lsp = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lcp = ""
lPACKED = ""
lUNPACKED = ""
lAdj = ""

lSC = lvw1.SelectedItem
With r
.Open "select * from stock where stockcodemain='" & lvw1.SelectedItem & "'", c, adOpenDynamic, adLockOptimistic
End With


ldesc.Caption = UCase(r!stockdesc)
lTAX = r!taxcode
lsp = Format(r!unitprice, "#####.00")
r.Close
With r
.Open "select * from stock where stockcodemain='" & lvw1.SelectedItem & "'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    qtyOnhand = qtyOnhand + CDbl(r!QTY)
    r.MoveNext
Loop


lqty = qtyOnhand
r.Close

With r
.Open "select * from sales where itemcodemain='" & lvw1.SelectedItem & "'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    QTYSOLD = QTYSOLD + CDbl(r!QTY)
    r.MoveNext
Loop
r.Close
lsold = QTYSOLD
With r
.Open "select * from stockpurchasehistory where stockcodemain='" & lvw1.SelectedItem & "' order by datepurchased desc", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = False Then
    lSUP = UCase(r!supplier)
    lPdate = r!datepurchased
Else
    lSUP = "Internal"
    lPdate = "N/A"
End If

r.Close

With r
.Open "select * from stockpurchasehistory where stockcodemain='" & lvw1.SelectedItem & "' order by datepurchased desc", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    NumofPurchase = NumofPurchase + 1
    TotalCP = TotalCP + r!costofiteminc
    r.MoveNext
Loop
r.Close
lcp = Format(TotalCP / NumofPurchase, "#####.00")
With r
.Open "select * from stockpurchasehistory where stockcodemain='" & lvw1.SelectedItem & "' order by datepurchased desc", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    TotalPurchased = TotalPurchased + CDbl(r!qtypurchased)
    r.MoveNext
Loop
lTOTPurch = TotalPurchased


r.Close
TotAdj = 0
With r
.Open "select * from stockadjustment where stockcode='" & lvw1.SelectedItem & "'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    TotAdj = TotAdj + r!adjustedby
    r.MoveNext
Loop
r.Close

If CDbl(lTOTPurch) <> CDbl(lqty) + CDbl(lsold) Then
    With r
    .Open "select * from packslist where pstockcode='" & lvw1.SelectedItem & "'", c, adOpenDynamic, adLockOptimistic
    End With
    If r.EOF = False Then
        QtyUnPacked = CDbl(lTOTPurch) - CDbl(lqty) - CDbl(lsold) + TotAdj
        QtyPacked = 0
    Else
        QtyPacked = CDbl(lTOTPurch) - CDbl(lqty) - CDbl(lsold) + TotAdj
        QtyUnPacked = 0
    End If
    r.Close
    lPACKED = QtyPacked
    lUNPACKED = QtyUnPacked
Else
    lPACKED = "0"
    lUNPACKED = "0"
End If

lAdj.Caption = TotAdj

End Sub


Private Sub lvw1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    If lvw1.ListItems.Count = 0 Then Exit Sub
    If Len(lvw1.SelectedItem) = 0 Then Exit Sub
    frmStockEdit.Show vbModal
End If

End Sub

Private Sub optBarcode_Click()
txtcode.Enabled = False
txtdesc.Enabled = False
txtBarcode.Enabled = True
'txtsn.Enabled = False
lAdj = ""

txtcode = ""
txtdesc = ""
lvw1.ListItems.Clear
lSC = ""
ldesc = ""
lTAX = ""
lsp = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lcp = ""
lPACKED = ""
lUNPACKED = ""
txtBarcode.SetFocus

End Sub

Private Sub optDesc_Click()
txtdesc.Enabled = True
txtcode.Enabled = False
txtBarcode.Enabled = False
'txtsn.Enabled = False
txtcode = ""
txtBarcode = ""
lvw1.ListItems.Clear
lSC = ""
ldesc = ""
lTAX = ""
lsp = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lcp = ""
lPACKED = ""
lUNPACKED = ""
lAdj = ""

txtdesc.SetFocus

End Sub

Private Sub optSN_Click()
txtcode.Enabled = False
txtdesc.Enabled = False
txtBarcode.Enabled = False
txtsn.Enabled = True

txtsn.SetFocus

End Sub

Private Sub OptStockCode_Click()
txtcode.Enabled = True
txtdesc.Enabled = False
txtBarcode.Enabled = False
'txtsn.Enabled = False
txtBarcode = ""
txtdesc = ""
lvw1.ListItems.Clear
lSC = ""
ldesc = ""
lTAX = ""
lsp = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lcp = ""
lPACKED = ""
lUNPACKED = ""
lAdj = ""

txtcode.SetFocus

End Sub

Private Sub txtBarcode_KeyPress(KeyAscii As Integer)
Dim search As String
Dim rs As New Recordset
lvw1.ListItems.Clear
lSC = ""
ldesc = ""
lTAX = ""
lsp = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lcp = ""
lPACKED = ""
lUNPACKED = ""
lAdj = ""


If KeyAscii <> 8 Then
    search = "%" & txtBarcode.Text + Chr(KeyAscii) + "%"
Else
    If Len(txtBarcode.Text) > 1 Then
    search = "%" & Mid(txtBarcode.Text, 1, Len(txtBarcode) - 1) + "%"
    Else
    Exit Sub
    End If
End If


With rs
.Open "select distinct(stockcodemain) from stock where stockcode like '" & search & "' order by stockcodeMAIN", c, adOpenDynamic, adLockOptimistic
If .EOF = False Then
    Do While .EOF = False
        Set li = lvw1.ListItems.Add(, , !stockcodeMAIN)
        'lvw1.ListItems(1).ForeColor = vbBlack
        'lvw1.ListItems(1).Bold = False
        'lvw1.ColumnHeaders(1).Width = 1440

        With li
        '.SubItems(1) = rs!stockcode
        '.SubItems(2) = rs!serialnumber
        '.SubItems(3) = rs!stockdesc
        '.SubItems(4) = rs!taxcode
        '.SubItems(5) = rs!unitprice

        End With
        .MoveNext
    Loop
End If
.Close
End With

End Sub


Private Sub txtcode_KeyPress(KeyAscii As Integer)
Dim search As String
Dim rs As New Recordset
lvw1.ListItems.Clear
lSC = ""
ldesc = ""
lTAX = ""
lsp = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lcp = ""
lPACKED = ""
lUNPACKED = ""
lAdj = ""



If KeyAscii <> 8 Then
    search = "%" & txtcode.Text + Chr(KeyAscii) + "%"
Else
    If Len(txtcode.Text) > 1 Then
    search = "%" & Mid(txtcode.Text, 1, Len(txtcode) - 1) + "%"
    Else
    Exit Sub
    End If
End If


With rs
.Open "select distinct(stockcodemain) from stock where stockcodemain like '" & search & "' order by stockcodeMAIN", c, adOpenDynamic, adLockOptimistic
If .EOF = False Then
    Do While .EOF = False
        Set li = lvw1.ListItems.Add(, , !stockcodeMAIN)
        'lvw1.ListItems(1).ForeColor = vbBlack
        'lvw1.ListItems(1).Bold = False
        'lvw1.ColumnHeaders(1).Width = 1440

        With li
        '.SubItems(1) = rs!stockcode
        '.SubItems(2) = rs!serialnumber
        '.SubItems(3) = rs!stockdesc
        '.SubItems(4) = rs!taxcode
        '.SubItems(5) = rs!unitprice

        End With
        .MoveNext
    Loop
End If
.Close
End With

End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
Dim search As String
Dim rs As New Recordset
lvw1.ListItems.Clear
lSC = ""
ldesc = ""
lTAX = ""
lsp = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lcp = ""
lPACKED = ""
lUNPACKED = ""
lAdj = ""


If KeyAscii <> 8 Then
    search = "%" & txtdesc.Text + Chr(KeyAscii) + "%"
Else
    If Len(txtdesc.Text) > 1 Then
    search = "%" & Mid(txtdesc.Text, 1, Len(txtdesc) - 1) + "%"
    Else
    Exit Sub
    End If
End If


With rs
.Open "select distinct(stockcodemain) from stock where stockdesc like '" & search & "' order by stockcodemain", c, adOpenDynamic, adLockOptimistic
If .EOF = False Then
    Do While .EOF = False
        Set li = lvw1.ListItems.Add(, , !stockcodeMAIN)
        'lvw1.ListItems(1).ForeColor = vbBlack
        'lvw1.ListItems(1).Bold = False
        'lvw1.ColumnHeaders(1).Width = 1440

        With li
        '.SubItems(1) = rs!stockcode
        '.SubItems(2) = rs!serialnumber
        '.SubItems(3) = rs!stockdesc
        '.SubItems(4) = rs!TaxCode
        '.SubItems(5) = rs!unitprice
        End With
        .MoveNext
    Loop
End If
.Close
End With

End Sub

Private Sub txtsn_KeyPress(KeyAscii As Integer)
Dim search As String
Dim rs As New Recordset
lvw1.ListItems.Clear
lSC = ""
ldesc = ""
lTAX = ""
lsp = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""



If KeyAscii <> 8 Then
    search = "%" & txtsn.Text + Chr(KeyAscii) + "%"
Else
    If Len(txtsn.Text) > 1 Then
    search = "%" & Mid(txtsn.Text, 1, Len(txtsn) - 1) + "%"
    Else
    Exit Sub
    End If
End If


With rs
.Open "select * from stock where serialnumber like '" & search & "' order by stockcode", c, adOpenDynamic, adLockOptimistic
If .EOF = False Then
    Do While .EOF = False
        Set li = lvw1.ListItems.Add(, , !stockcodeMAIN)
        'lvw1.ListItems(1).ForeColor = vbBlack
        'lvw1.ListItems(1).Bold = False
        'lvw1.ColumnHeaders(1).Width = 1440

        With li
        '.SubItems(1) = rs!stockcode
        '.SubItems(2) = rs!serialnumber
        '.SubItems(3) = rs!stockdesc
        '.SubItems(4) = rs!taxcode
        '.SubItems(5) = rs!unitprice

        End With
        .MoveNext
    Loop
End If
.Close
End With

End Sub



