VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmStockAdd 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Stock"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin LVbuttons.LaVolpeButton LaVolpeButton7 
      Height          =   375
      Left            =   10200
      TabIndex        =   25
      Top             =   8040
      Width           =   1575
      _ExtentX        =   2778
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
      BCOL            =   13160660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmStockAdd.frx":0000
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
      Caption         =   "Items Lines"
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
      Height          =   2055
      Left            =   240
      TabIndex        =   45
      Top             =   5880
      Width           =   11535
      Begin MSComctlLib.ListView lvw1 
         Height          =   1695
         Left            =   120
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Stock Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Barcode"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Serial Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cost(Excl)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Input VAT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Cost(Incl.)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Selling Price(Incl)"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "TAX Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Dep"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   -360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStockAdd.frx":001C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000016&
      Caption         =   "Supplier"
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
      Height          =   4695
      Left            =   6480
      TabIndex        =   34
      Top             =   240
      Width           =   5295
      Begin VB.TextBox lBAL 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   3720
         Width           =   1935
      End
      Begin VB.ComboBox cTerms 
         Height          =   315
         ItemData        =   "frmStockAdd.frx":046E
         Left            =   1800
         List            =   "frmStockAdd.frx":04BD
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox tFAX 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox tTEL 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox tADD4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox tADD3 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox tADD2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox tADD1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox tSUP 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton2 
         Height          =   375
         Left            =   4200
         TabIndex        =   0
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Find"
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
         BCOL            =   16777215
         FCOL            =   0
         FCOLO           =   33023
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmStockAdd.frx":0538
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "1"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton6 
         Height          =   375
         Left            =   4200
         TabIndex        =   44
         Top             =   840
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Reset"
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
         BCOL            =   16777215
         FCOL            =   0
         FCOLO           =   33023
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmStockAdd.frx":0554
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "1"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.Label Label19 
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
         TabIndex        =   50
         Top             =   3720
         Width           =   1695
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
         TabIndex        =   42
         Top             =   4200
         Width           =   1575
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
         TabIndex        =   41
         Top             =   3240
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
         TabIndex        =   40
         Top             =   2760
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
         TabIndex        =   39
         Top             =   2280
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
         TabIndex        =   38
         Top             =   1800
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
         TabIndex        =   37
         Top             =   1320
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
         TabIndex        =   36
         Top             =   840
         Width           =   1335
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
         TabIndex        =   35
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "Stock Details"
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
      Height          =   5895
      Left            =   240
      TabIndex        =   28
      Top             =   0
      Width           =   6135
      Begin VB.ComboBox cDep 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   690
         Width           =   2535
      End
      Begin VB.TextBox tScanQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   22
         Top             =   5010
         Width           =   495
      End
      Begin VB.TextBox tSN 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   23
         Top             =   5370
         Width           =   2055
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton8 
         Height          =   375
         Left            =   4680
         TabIndex        =   51
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "R&eset"
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
         MICON           =   "frmStockAdd.frx":0570
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
      Begin VB.TextBox tINVNO 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox tCOST 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   19
         Top             =   3600
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00404040&
         Caption         =   "Calculate Selling Price using MarkUp"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   4080
         Width           =   4455
      End
      Begin VB.TextBox tSell 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   21
         Top             =   4440
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker cDates 
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   127270913
         CurrentDate     =   38461
      End
      Begin VB.ComboBox cTaxCode 
         Height          =   315
         ItemData        =   "frmStockAdd.frx":058C
         Left            =   1680
         List            =   "frmStockAdd.frx":058E
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   3240
         Width           =   1215
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton1 
         Height          =   375
         Left            =   3480
         TabIndex        =   14
         Top             =   1560
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Find"
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
         BCOL            =   16777215
         FCOL            =   0
         FCOLO           =   33023
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmStockAdd.frx":0590
         ALIGN           =   1
         IMGLST          =   "ImageList1"
         IMGICON         =   "1"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.TextBox tDESC 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   2520
         Width           =   4215
      End
      Begin VB.TextBox tBC 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   16
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox tSC 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox tQTY 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton9 
         Height          =   255
         Left            =   4320
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "&New"
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
         BCOL            =   16777215
         FCOL            =   0
         FCOLO           =   33023
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmStockAdd.frx":05AC
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "1"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Department:"
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
         Left            =   240
         TabIndex        =   60
         Top             =   690
         Width           =   1335
      End
      Begin VB.Label lGRV 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4680
         TabIndex        =   59
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "GRV No:"
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
         Left            =   3240
         TabIndex        =   58
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label L22 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Scan Qty Only if you are scanning serial number!! If not leave Scan Qty blank."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   3840
         TabIndex        =   57
         Top             =   4965
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lScan 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         TabIndex        =   56
         Top             =   5010
         Width           =   615
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Scanned:"
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
         Left            =   2280
         TabIndex        =   55
         Top             =   5010
         Width           =   855
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Scan Qty:"
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
         TabIndex        =   54
         Top             =   5010
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Number:"
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
         TabIndex        =   53
         Top             =   5400
         Width           =   1575
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Inv No.:"
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
         Left            =   3240
         TabIndex        =   49
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1. Cost price per Item (Excl.):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   48
         Top             =   3600
         Width           =   3015
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "2. Recommended Selling price (Incl.):"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   47
         Top             =   4440
         Width           =   3735
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Date:"
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
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Code:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
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
         Left            =   240
         TabIndex        =   32
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode:"
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
         Left            =   240
         TabIndex        =   31
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Code:"
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
         Left            =   240
         TabIndex        =   30
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity:"
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
         Left            =   240
         TabIndex        =   29
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   1740
         Left            =   120
         Top             =   3120
         Width           =   5775
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   900
         Left            =   120
         Top             =   4920
         Width           =   5775
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton3 
      Height          =   495
      Left            =   6840
      TabIndex        =   24
      Top             =   5280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Add Item Line"
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
      MICON           =   "frmStockAdd.frx":05C8
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton4 
      Height          =   495
      Left            =   8400
      TabIndex        =   26
      Top             =   5280
      Width           =   1455
      _ExtentX        =   2566
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
      BCOL            =   0
      FCOL            =   16777215
      FCOLO           =   33023
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmStockAdd.frx":05E4
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
   Begin LVbuttons.LaVolpeButton LaVolpeButton5 
      Height          =   495
      Left            =   9960
      TabIndex        =   27
      Top             =   5280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Reset"
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
      BCOL            =   16777215
      FCOL            =   0
      FCOLO           =   33023
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmStockAdd.frx":0600
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "1"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Label lTQty 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty - 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   52
      Top             =   8160
      Width           =   1935
   End
   Begin VB.Menu mnurc 
      Caption         =   "RightClick"
      Visible         =   0   'False
      Begin VB.Menu mnuedit 
         Caption         =   "Edit Item"
      End
      Begin VB.Menu mnudel 
         Caption         =   "&Delete Item"
      End
   End
End
Attribute VB_Name = "frmStockAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Dim r As New Recordset
Dim r1 As New Recordset
Dim nnCINC As Double
If Len(cTaxCode.Text) = 0 Then
    MsgBox "Please select tax code!"
    cTaxCode.SetFocus
    Exit Sub
End If

If Len(tCOST.Text) = 0 Then
    If Check1.Value Then
        tSell.Text = ""
    Else
        tSell.Text = SP
    End If
    
Exit Sub
End If
If Check1.Value Then
    With r
    .Open "select * from markup", c, adOpenDynamic, adLockOptimistic
    End With
    With r1
    .Open "select * from taxcode where taxcode='" & cTaxCode.Text & "'", c, adOpenDynamic, adLockOptimistic
    End With
    x1tax = CDbl(r1!tax)
    r1.Close
    'nnCINC = CDbl(tCOST) + (CDbl(tCOST) * x1tax / 100)
    nnCINC = CDbl(tCOST)
    
    tSell.Text = nnCINC * (CDbl(r!markup) + 100) / 100
    tSell.Text = CDbl(tSell.Text) + (CDbl(tSell) * x1tax / 100)
    tSell = Format(tSell, "#####.00")
    
    
Else
   tSell = Format(SP, "#####.00")
End If

    
    
End Sub

Private Sub cTaxCode_Click()
Dim r As New Recordset
Dim r1 As New Recordset
Dim nnCINC As Double

If Len(tCOST.Text) = 0 Then
    If Check1.Value Then
        tSell.Text = ""
    Else
        tSell.Text = SP
    End If
    
Exit Sub
End If
If Check1.Value Then
    With r
    .Open "select * from markup", c, adOpenDynamic, adLockOptimistic
    End With
    With r1
    .Open "select * from taxcode where taxcode='" & cTaxCode.Text & "'", c, adOpenDynamic, adLockOptimistic
    End With
    If r1.EOF = False Then
    x1tax = CDbl(r1!tax)
    r1.Close
    nnCINC = CDbl(tCOST)
    
    tSell.Text = nnCINC * (CDbl(r!markup) + 100) / 100
    tSell.Text = CDbl(tSell.Text) + (CDbl(tSell) * x1tax / 100)
    tSell = Format(tSell, "#####.00")
    Else
    tsel = ""
    End If
Else
   tSell = Format(SP, "#####.00")
End If


End Sub


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
SP = 0
cDates.Value = Format(Date, "DD/MM/YYYY")
cTaxCode.Text = "1"
With r
.Open "select * from grv", c, adOpenDynamic, adLockOptimistic
End With
lGRV = "GRV" & r!grvno
r.Close
With r
.Open "select * from department order by department", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    cDep.AddItem r!department
    r.MoveNext
Loop
r.Close

End Sub

Private Sub LaVolpeButton1_Click()
SP = 0
frmItemSearch.Show vbModal

End Sub

Private Sub LaVolpeButton2_Click()
frmSupplierSearch.Show vbModal

End Sub

Public Sub LaVolpeButton3_Click()
Dim r As New Recordset
Dim r1 As New Recordset
Dim li As ListItem
Dim xTAX As Double
Dim TotLQty As Double

If Len(tScanQty) > 0 And Len(tSN) = 0 Then Exit Sub
If Len(cDep.Text) = 0 Then
    MsgBox "Please select department!", vbInformation, "Status"
    cDep.SetFocus
    Exit Sub
End If

If Len(tQTY) = 0 Then
    MsgBox "Please enter quantity!", vbInformation, "Status"
    tQTY.SetFocus
    Exit Sub
End If
If Len(tSC) = 0 Then
    MsgBox "Please enter Stock Code!", vbInformation, "Status"
    tSC.SetFocus
    Exit Sub
End If
If Len(tBC) = 0 Then
    tBC = "N/A"
End If
If Len(tDESC) = 0 Then
    MsgBox "Please enter Description!", vbInformation, "Status"
    tDESC.SetFocus
    Exit Sub
End If
If Len(tCOST) = 0 Then
    MsgBox "Please enter CostPrice!", vbInformation, "Status"
    tCOST.SetFocus
    Exit Sub
End If
If Len(cTaxCode.Text) = 0 Then
    MsgBox "Please select Tax Code!", vbInformation, "Status"
    cTaxCode.SetFocus
    Exit Sub
End If
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
If Len(tSell.Text) = 0 Then
    MsgBox "Please enter selling price!", vbExclamation, "Status"
    tSell.SetFocus
    Exit Sub
End If
If Len(tSN.Text) > 0 And CDbl(tQTY.Text) > 1 Then
    MsgBox "Qty cannot exceed 1 whenever a serial number exist!", vbExclamation, "Status"
    tQTY.Text = ""
    tQTY.SetFocus
    Exit Sub
End If
If Len(tSN.Text) > 0 And tSN <> "N/A" Then
    With r
    .Open "select * from serialnumber where serialnumber='" & tSN.Text & "'", c, adOpenDynamic, adLockOptimistic
    End With
    If r.EOF = False Then
        MsgBox "Cannot add item to list as there already exist an item with same serial number!", vbExclamation, "Status"
        tSN.Text = ""
        tSN.SetFocus
        Exit Sub
    End If
    r.Close
End If
If Len(tSN) = 0 Then
    tSN = "N/A"
End If
If tSN <> "N/A" Then
For i = 1 To lvw1.ListItems.Count
    If tSN.Text = lvw1.ListItems(i).SubItems(3) Then
        MsgBox "Cannot add item to list as there already exist an item with same serial number!", vbExclamation, "Status"
        tSN = ""
        tSN.SetFocus
        Exit Sub
    End If
Next i
End If
If Len(tScanQty) > 0 Then
    If CDbl(lScan) = CDbl(tScanQty) Then
        MsgBox "Cannot scan item as total qty has been achieved!", vbExclamationm, "Status"
        tSN.Text = ""
        Exit Sub
    End If
    lScan = CDbl(lScan) + 1


End If

SP = 0
Set li = lvw1.ListItems.Add(, , tQTY.Text)
With li
    .SubItems(1) = UCase(tSC.Text)
    .SubItems(2) = tBC.Text
    .SubItems(3) = tSN.Text
    .SubItems(4) = UCase(tDESC.Text)
    .SubItems(5) = Format(tCOST.Text, "#####.00")
    With r1
    .Open "select * from taxcode where taxcode='" & cTaxCode.Text & "'", c, adOpenDynamic, adLockOptimistic
    End With
    xTAX = r1!tax
    r1.Close
    .SubItems(6) = Format(tCOST.Text, "#####.00") * (xTAX / 100)
    .SubItems(7) = Format((Format(tCOST.Text, "#####.00")) + (Format(tCOST.Text, "#####.00") * (xTAX / 100)), "#####.00")
    .SubItems(8) = Format(tSell.Text, "#####.00")
    .SubItems(9) = cTaxCode.Text
    .SubItems(10) = cDep.Text
    
End With
TotLQty = 0
For i = 1 To lvw1.ListItems.Count
    TotLQty = TotLQty + CDbl(lvw1.ListItems(i))
Next i
lTQty.Caption = "Total Qty - " & TotLQty

    
'res = MsgBox("Do you wish to clear serial number and leave all other fields as they are?", vbYesNo + vbQuestion, "Status")
'If res = vbNo Then
 '   tQTY.Text = ""
 '   tSC = ""
 '   tBC = ""
 '   tSN = ""
 '   tDESC = ""
 '   tCOST = ""
 '   tSell.Text = ""
 '   Check1.Value = False
 '   cTaxCode.ListIndex = -1
    
  '  tQTY.SetFocus
'Else
If tSN = "N/A" Then
    SP = 0
    tQTY.Text = ""
    tSC = ""
    tBC = ""
    tSN = ""
    tDESC = ""
    'cTaxCode.ListIndex = -1
    tCOST = ""
    tSell.Text = ""
    tSN = ""
    tScanQty.Text = ""
    tQTY.SetFocus
Else

    
    tSN = ""
    tSN.SetFocus
End If
If lScan = tScanQty Then
    SP = 0
    tQTY.Text = ""
    tSC = ""
    tBC = ""
    tSN = ""
    tDESC = ""
    'cTaxCode.ListIndex = -1
    tCOST = ""
    tSell.Text = ""
    lScan = "0"
    tScanQty.Text = ""
    tQTY.SetFocus
End If

'End If



    





End Sub

Private Sub LaVolpeButton4_Click()
SP = 0
Unload Me

End Sub

Private Sub LaVolpeButton5_Click()
tQTY.Text = ""
tSC = ""
tBC = ""
tSN = ""
tDESC = ""
cTaxCode.ListIndex = -1
tCOST = ""
tSell.Text = ""
SP = 0
Check1.Value = False
lvw1.ListItems.Clear
End Sub

Private Sub LaVolpeButton6_Click()
tSUP.Text = ""
tADD1 = ""
tADD2 = ""
tADD3 = ""
tADD4 = ""
tTEL = ""
tFAX = ""
cTerms.ListIndex = -1

End Sub

Private Sub LaVolpeButton7_Click()
Dim r As New Recordset
Dim r1 As New Recordset
Dim r2 As New Recordset
Dim r10 As New Recordset
Dim xxTOTAL As Double
Dim INVTOTAL As Double

If Len(tINVNO.Text) = 0 Then
    MsgBox "Please enter invoice number!", vbExclamation, "Status"
    tINVNO.SetFocus
    Exit Sub
End If
If lvw1.ListItems.Count = 0 Then Exit Sub
If cTerms <> "Cash" Then
    INVTOTAL = 0
    For i = 1 To lvw1.ListItems.Count
        INVTOTAL = INVTOTAL + (CDbl(lvw1.ListItems(i).SubItems(7)) * CDbl(lvw1.ListItems(i)))
    Next i
    With r
    .Open "select * from creditorsinvoice", c, adOpenDynamic, adLockOptimistic
    End With
    r.AddNew
    r!invdate = Format(cDates.Value, "DD/MM/YYYY")
    r!invno = UCase(tINVNO.Text)
    r!creditor = UCase(tSUP.Text)
    r!terms = cTerms.Text
    r!paymentdate = Format(DateAdd("d", cTerms.Text, cDates.Value), "DD/MM/YYYY")
    r!tendered = Format(INVTOTAL, "#####.00")
    r!paid = "No"
    r.Update
    r.Close
Else
    INVTOTAL = 0
    For i = 1 To lvw1.ListItems.Count
        INVTOTAL = INVTOTAL + (CDbl(lvw1.ListItems(i).SubItems(7)) * CDbl(lvw1.ListItems(i)))
    Next i
    With r
    .Open "select * from creditorsinvoice", c, adOpenDynamic, adLockOptimistic
    End With
    r.AddNew
    r!invdate = Format(cDates.Value, "DD/MM/YYYY")
    r!invno = UCase(tINVNO.Text)
    r!creditor = UCase(tSUP.Text)
    r!terms = cTerms.Text
    r!paymentdate = Format(cDates.Value, "DD/MM/YYYY")
    r!tendered = Format(INVTOTAL, "#####.00")
    r!paid = "Yes"
    r.Update
    r.Close
End If
With r
.Open "select * from creditorspayment", c, adOpenDynamic, adLockOptimistic
End With
r.AddNew
r!paymentdate = Format(cDates.Value, "DD/MM/YYYY")
r!creditor = UCase(tSUP.Text)
r!chqno = "N/A"
r!invno = tINVNO.Text
r!amount = Format(INVTOTAL, "#####.00")
r.Update
r.Close



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
For i = 1 To lvw1.ListItems.Count
    With r
    .Open "select * from stock where stockcodemain='" & lvw1.ListItems(i).SubItems(1) & "' and stockcode='" & lvw1.ListItems(i).SubItems(2) & "'", c, adOpenDynamic, adLockOptimistic
    End With
    If r.EOF = False Then
        r!QTY = CDbl(r!QTY) + CDbl(lvw1.ListItems(i))
        r!taxcode = lvw1.ListItems(i).SubItems(9)
        r!unitprice = Format(lvw1.ListItems(i).SubItems(8), "#####.00")
        r.Update
    Else
        r.AddNew
        r!stockcodeMAIN = UCase(lvw1.ListItems(i).SubItems(1))
        r!stockcode = lvw1.ListItems(i).SubItems(2)
        'r!serialnumber = lvw1.ListItems(i).SubItems(3)
        r!stockdesc = UCase(lvw1.ListItems(i).SubItems(4))
        r!QTY = lvw1.ListItems(i)
        r!unitprice = lvw1.ListItems(i).SubItems(8)
        r!taxcode = lvw1.ListItems(i).SubItems(9)
        r!department = lvw1.ListItems(i).SubItems(10)
        r.Update
    End If
    r.Close
    
    With r
    .Open "select * from stockpurchasehistory", c, adOpenDynamic, adLockOptimistic
    End With
    r.AddNew
    r!pid = lGRV
    r!stockcodeMAIN = UCase(lvw1.ListItems(i).SubItems(1))
    r!stockcode = lvw1.ListItems(i).SubItems(2)
    r!serialnumber = lvw1.ListItems(i).SubItems(3)
    r!supplier = UCase(tSUP)
    r!costofItemEXC = Format(lvw1.ListItems(i).SubItems(5), "#####.00")
    r!vatinput = Format(lvw1.ListItems(i).SubItems(6), "#####.00")
    r!costofiteminc = Format(lvw1.ListItems(i).SubItems(7), "#####.00")
    r!sellingpriceINc = Format(lvw1.ListItems(i).SubItems(8), "#####.00")
    r!datepurchased = Format(cDates.Value, "DD/MM/YYYY")
    r!qtypurchased = lvw1.ListItems(i)
    r!department = lvw1.ListItems(i).SubItems(10)

    r.Update
    r.Close
    'If Dir("\\" & servername & "\\" & sharename & "\LABELS", vbDirectory) = "" Then
    'MkDir "\\" & servername & "\\" & sharename & "\LABELS"
    'End If

    'If Dir("\\" & servername & "\\" & sharename & "\LABELS\" & UCase(lvw1.ListItems(i).SubItems(1)), vbDirectory) = "" Then
    'MkDir "\\" & servername & "\\" & sharename & "\LABELS\" & UCase(lvw1.ListItems(i).SubItems(1))
    'End If
    'Open "\\" & servername & "\\" & sharename & "\LABELS\" & UCase(lvw1.ListItems(i).SubItems(1)) & "\barcode.TXT" For Output As #2
    'Print #2, lvw1.ListItems(i).SubItems(2)
    'Close #2
    'Open "\\" & servername & "\\" & sharename & "\LABELS\" & UCase(lvw1.ListItems(i).SubItems(1)) & "\price.TXT" For Output As #2
    'Print #2, lvw1.ListItems(i).SubItems(8)
    'Close #2
    'Open "\\" & servername & "\\" & sharename & "\LABELS\" & UCase(lvw1.ListItems(i).SubItems(1)) & "\description.TXT" For Output As #2
    'Print #2, lvw1.ListItems(i).SubItems(4)
    'Close #2
    'Open "\\" & servername & "\\" & sharename & "\LABELS\" & UCase(lvw1.ListItems(i).SubItems(1)) & "\stockcode.TXT" For Output As #2
    'Print #2, lvw1.ListItems(i).SubItems(1)
    'Close #2

    With r10
    .Open "select * from costcode", c, adOpenDynamic, adLockOptimistic
    End With
    lcp = Len(lvw1.ListItems(i).SubItems(7))
    For ii = 1 To lcp
        If Mid(lvw1.ListItems(i).SubItems(7), ii, 1) = "1" Then
            cc = cc & r10!one
        End If
        If Mid(lvw1.ListItems(i).SubItems(7), ii, 1) = "2" Then
            cc = cc & r10!two
        End If
        If Mid(lvw1.ListItems(i).SubItems(7), ii, 1) = "3" Then
            cc = cc & r10!three
        End If
        If Mid(lvw1.ListItems(i).SubItems(7), ii, 1) = "4" Then
            cc = cc & r10!four
        End If
        If Mid(lvw1.ListItems(i).SubItems(7), ii, 1) = "5" Then
            cc = cc & r10!five
        End If
        If Mid(lvw1.ListItems(i).SubItems(7), ii, 1) = "6" Then
            cc = cc & r10!six
        End If
        If Mid(lvw1.ListItems(i).SubItems(7), ii, 1) = "7" Then
            cc = cc & r10!seven
        End If
        If Mid(lvw1.ListItems(i).SubItems(7), ii, 1) = "8" Then
            cc = cc & r10!eight
        End If
        If Mid(lvw1.ListItems(i).SubItems(7), ii, 1) = "9" Then
            cc = cc & r10!nine
        End If
        If Mid(lvw1.ListItems(i).SubItems(7), ii, 1) = "0" Then
            cc = cc & r10!zero
        End If
        If Mid(lvw1.ListItems(i).SubItems(7), ii, 1) = "." Then
            cc = cc & "/"
        End If
    Next ii
    r10.Close
    
    'Open "\\" & servername & "\\" & sharename & "\LABELS\" & UCase(lvw1.ListItems(i).SubItems(1)) & "\costprice.TXT" For Output As #2
    'Print #2, cc
    'Close #2

Next i

For i = 1 To lvw1.ListItems.Count
    If lvw1.ListItems(i).SubItems(3) <> "N/A" Then
        With r
        .Open "select * from serialnumber", c, adOpenDynamic, adLockOptimistic
        End With
        r.AddNew
        r!stockcode = UCase(lvw1.ListItems(i).SubItems(1))
        r!serialnumber = lvw1.ListItems(i).SubItems(3)
        r.Update
        r.Close
    End If
    
Next i


If cTerms.Text <> "Cash" Then
    With r
    .Open "select * from creditorsjournal", c, adOpenDynamic, adLockOptimistic
    End With
    r.AddNew
    r!documentno = tINVNO
    r!dt = Format(cDates.Value, "DD/MM/YYYY")
    r!details = UCase(tSUP.Text)
    With r1
    .Open "select * from ledgeraccount where account='" & tSUP.Text & "'", c, adOpenDynamic, adLockOptimistic
    End With
    r!folio = r1!accountno
    r1.Close
    xxTOTAL = 0
    For i = 1 To lvw1.ListItems.Count
        xxTOTAL = xxTOTAL + (lvw1.ListItems(i) * CDbl(lvw1.ListItems(i).SubItems(7)))
    Next i
    r!total = Format(xxTOTAL, "#####.00")
    r!inventory = Format(xxTOTAL, "#####.00")
    r!equipment = "0"
    r!Stationery = "0"
    r.Update
    r.Close
Else
    With r
    .Open "select * from pettycashjournal", c, adOpenDynamic, adLockOptimistic
    End With
    r.AddNew
    r!Document = tINVNO.Text
    r!dt = Format(cDates.Value, "DD/MM/YYYY")
    r!details = UCase(tSUP.Text)
    With r1
    .Open "select * from ledgeraccount where account='" & tSUP.Text & "'", c, adOpenDynamic, adLockOptimistic
    End With
    r!folio = r1!accountno
    r1.Close
    xxTOTAL = 0
    For i = 1 To lvw1.ListItems.Count
        xxTOTAL = xxTOTAL + (lvw1.ListItems(i) * CDbl(lvw1.ListItems(i).SubItems(7)))
    Next i
    r!amount = Format(xxTOTAL, "#####.00")
    r!salaries = "0"
    r!repairs = "0"
    r!sundry = "0"
    r!inventory = Format(xxTOTAL, "#####.00")
    r!debtorscontrol = "0"
    r!creditorscontrol = "0"
    r!telephone = "0"
    r!donation = "0"
    r.Update
    r.Close
    
End If

    With r
    .Open "select * from grv", c, adOpenDynamic, adLockOptimistic
    End With
    r!grvno = CInt(r!grvno) + 1
    r.Update
    r.Close
    
res = MsgBox("Invoice successfully added! Do you wish to add another?", vbYesNo + vbQuestion, "Add Another?")
If res = vbNo Then
    SP = 0
    Unload Me
Else
lvw1.ListItems.Clear
tINVNO.Text = ""
cDep.ListIndex = -1


    SP = 0
    tQTY.Text = ""
    tSC = ""
    tBC = ""
    tSN = ""
    tDESC = ""
    'cTaxCode.ListIndex = -1
    tCOST = ""
    tSell.Text = ""
    tSUP.Text = ""
    tADD1.Text = ""
    tADD2.Text = ""
    tADD3.Text = ""
    tADD4.Text = ""
    tTEL.Text = ""
    tFAX.Text = ""
    With r
    .Open "select * from grv", c, adOpenDynamic, adLockOptimistic
    End With
    lGRV = "GRV" & r!grvno
    r.Close
    
    cTaxCode.Text = "1"
    Check1.Value = False
    tINVNO.SetFocus
End If

End Sub

Private Sub LaVolpeButton8_Click()
tQTY.Text = ""
cDep.ListIndex = -1

tSC = ""
tBC = ""
tSN = ""
tDESC = ""
'cTaxCode.ListIndex = -1
tCOST = ""
tSell.Text = ""
SP = 0
Check1.Value = False
lScan = "0"
tScanQty = ""

cDep.SetFocus
'lvw1.ListItems.Clear

End Sub

Private Sub LaVolpeButton9_Click()
Dim r As New Recordset
frmCreateDep.Show vbModal
cDep.Clear
With r
.Open "select * from department order by department", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    cDep.AddItem r!department
    r.MoveNext
Loop
r.Close

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


Private Sub lvw1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TotLQty As Double
If lvw1.ListItems.Count = 0 Then Exit Sub
If Len(lvw1.SelectedItem) = 0 Then Exit Sub

If Button = vbRightButton Then

Me.PopupMenu mnurc

End If

End Sub


Private Sub mnudel_Click()
    res = MsgBox("Are you sure you wish to remove Item?", vbYesNo + vbQuestion, "Remove Item?")
    If res = vbNo Then Exit Sub
    lvw1.ListItems.Remove lvw1.SelectedItem.Index
    TotLQty = 0
    For i = 1 To lvw1.ListItems.Count
        TotLQty = TotLQty + CDbl(lvw1.ListItems(i))
    Next i
    lTQty.Caption = "Total Qty - " & TotLQty

End Sub

Private Sub mnuedit_Click()
    res = MsgBox("Are you sure you wish to edit Item?", vbYesNo + vbQuestion, "Remove Item?")
    If res = vbNo Then Exit Sub
    tQTY = lvw1.SelectedItem
    tSC = lvw1.SelectedItem.SubItems(1)
    tBC = lvw1.SelectedItem.SubItems(2)
    tSN = lvw1.SelectedItem.SubItems(3)
    tDESC = lvw1.SelectedItem.SubItems(4)
    ttaxcode = lvw1.SelectedItem.SubItems(9)
    tCOST = lvw1.SelectedItem.SubItems(5)
    tSell = lvw1.SelectedItem.SubItems(8)
    cDep.Text = lvw1.SelectedItem.SubItems(10)
    lScan = "0"
tScanQty = ""

    
    lvw1.ListItems.Remove lvw1.SelectedItem.Index
    TotLQty = 0
    For i = 1 To lvw1.ListItems.Count
        TotLQty = TotLQty + CDbl(lvw1.ListItems(i))
    Next i
    lTQty.Caption = "Total Qty - " & TotLQty

End Sub

Private Sub tBC_Change()
lScan = "0"
tScanQty = ""

End Sub

Private Sub tCOST_Change()
Dim r As New Recordset
Dim r1 As New Recordset
Dim nnCINC As Double
If Len(cTaxCode.Text) = 0 Then
    MsgBox "Please select tax code!"
    cTaxCode.SetFocus
    Exit Sub
End If

If Len(tCOST.Text) = 0 Then
    If Check1.Value Then
        tSell.Text = ""
    Else
        tSell.Text = SP
    End If
    
Exit Sub
End If
If Check1.Value Then
    With r
    .Open "select * from markup", c, adOpenDynamic, adLockOptimistic
    End With
    With r1
    .Open "select * from taxcode where taxcode='" & cTaxCode.Text & "'", c, adOpenDynamic, adLockOptimistic
    End With
    x1tax = CDbl(r1!tax)
    r1.Close
    nnCINC = CDbl(tCOST)
    
    tSell.Text = nnCINC * (CDbl(r!markup) + 100) / 100
    tSell.Text = CDbl(tSell.Text) + (CDbl(tSell) * x1tax / 100)
    tSell = Format(tSell, "#####.00")
    
    
Else
   tSell = Format(SP, "#####.00")
End If

    

End Sub

Private Sub tCOST_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    A = InStr(1, tCOST, ".")
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


Private Sub tINVNO_Change()
lScan = "0"
tScanQty = ""

End Sub

Private Sub tQTY_Change()
lScan = "0"
tScanQty = ""

End Sub

Private Sub tQTY_KeyPress(KeyAscii As Integer)

If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If

End Sub


Private Sub tSC_Change()
lScan = "0"
tScanQty = ""

End Sub

Private Sub tSC_KeyPress(KeyAscii As Integer)
Dim rs As New Recordset
Dim r As New Recordset
If KeyAscii <> 8 Then
    search = tSC.Text + Chr(KeyAscii)
Else
    If Len(tSC.Text) > 1 Then
    search = Mid(tSC.Text, 1, Len(tSC) - 1)
    Else
    Exit Sub
    End If
End If


With rs
.Open "select * from stock where stockcodemain = '" & search & "'", c, adOpenDynamic, adLockOptimistic
End With
If rs.EOF = False Then
    tBC.Text = rs!stockcode
    tDESC.Text = rs!stockdesc
    cTaxCode.Text = rs!taxcode
    
    With r
    .Open "select * from stockpurchasehistory where stockcodemain='" & search & "' order by pid desc", c, adOpenDynamic, adLockOptimistic
    End With
    If r.EOF = False Then
    tCOST.Text = r!costofItemEXC
    End If
    r.Close
    tSell.Text = rs!unitprice
Else
    tBC.Text = ""
    tDESC.Text = ""
    cTaxCode.Text = "1"
    
    tCOST.Text = ""
    tSell.Text = ""
End If

    
End Sub

Private Sub tScanQty_GotFocus()
L22.Visible = True

End Sub

Private Sub tScanQty_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If

End Sub

Private Sub tScanQty_LostFocus()
L22.Visible = False

End Sub

Private Sub tSell_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    A = InStr(1, tSell, ".")
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


Private Sub tSN_GotFocus()
L22.Visible = True
End Sub

Private Sub tSN_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    'res = MsgBox("Update entry without updating any other fields?", vbYesNo + vbQuestion, "Update?")
    'If res = vbYes Then
    If Len(tSN.Text) = 0 Then Exit Sub
        Me.LaVolpeButton3_Click
    'End If
End If


End Sub


Private Sub tSN_LostFocus()
L22.Visible = False
End Sub


