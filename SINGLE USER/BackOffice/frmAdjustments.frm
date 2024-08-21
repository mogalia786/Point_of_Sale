VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAdjustments 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Adjustments"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H80000016&
      Caption         =   "Add Serial Numbers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   41
      Top             =   5760
      Width           =   7215
      Begin LVbuttons.LaVolpeButton LaVolpeButton3 
         Height          =   255
         Left            =   3600
         TabIndex        =   44
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "&Add"
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
         MICON           =   "frmAdjustments.frx":0000
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
      Begin VB.ListBox lSn 
         Appearance      =   0  'Flat
         Height          =   1200
         Left            =   4560
         TabIndex        =   43
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtsn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton4 
         Height          =   255
         Left            =   3600
         TabIndex        =   45
         Top             =   960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "&Remove"
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
         MICON           =   "frmAdjustments.frx":001C
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
         Caption         =   "Serial Number"
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
         Left            =   -120
         TabIndex        =   42
         Top             =   720
         Width           =   1455
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton1 
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   7440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Adjust Stock"
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
      MICON           =   "frmAdjustments.frx":0038
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
      Height          =   2415
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   7215
      Begin VB.TextBox txtAdj 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Adjusted:"
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
         TabIndex        =   40
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label ladj 
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
         TabIndex        =   39
         Top             =   2040
         Width           =   2655
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
         TabIndex        =   38
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
         TabIndex        =   37
         Top             =   600
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
         TabIndex        =   36
         Top             =   960
         Width           =   1095
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
         TabIndex        =   35
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Adjust stock to:"
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
         TabIndex        =   34
         Top             =   2040
         Width           =   1335
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
         TabIndex        =   33
         Top             =   240
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
         TabIndex        =   32
         Top             =   600
         Width           =   2655
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
         TabIndex        =   31
         Top             =   960
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
         TabIndex        =   30
         Top             =   1320
         Width           =   2175
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
         TabIndex        =   29
         Top             =   240
         Width           =   1095
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
         TabIndex        =   27
         Top             =   960
         Width           =   1695
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
         TabIndex        =   26
         Top             =   1320
         Width           =   1575
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
         TabIndex        =   25
         Top             =   240
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
         TabIndex        =   24
         Top             =   600
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
         TabIndex        =   23
         Top             =   960
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
         TabIndex        =   22
         Top             =   1320
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
         TabIndex        =   21
         Top             =   1680
         Width           =   2655
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
         TabIndex        =   20
         Top             =   1680
         Width           =   2655
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
         TabIndex        =   19
         Top             =   1680
         Width           =   1575
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
         TabIndex        =   18
         Top             =   1680
         Width           =   1335
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
      TabIndex        =   16
      Top             =   0
      Width           =   7215
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
      Begin VB.OptionButton optDesc 
         BackColor       =   &H80000016&
         Caption         =   "Item Description"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton OptStockCode 
         BackColor       =   &H80000016&
         Caption         =   "Stock Code"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000016&
      Caption         =   "Criteria (Click on Item to view details)"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   7215
      Begin MSComctlLib.ListView lvw1 
         Height          =   1215
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   2143
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
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   7215
      Begin VB.TextBox txtBarcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   5520
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtcode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
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
         Left            =   4320
         TabIndex        =   14
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
         Left            =   2280
         TabIndex        =   13
         Top             =   360
         Width           =   1095
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
         Left            =   0
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton2 
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   7440
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
      MICON           =   "frmAdjustments.frx":0054
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
Attribute VB_Name = "frmAdjustments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LaVolpeButton1_Click()
Dim r As New Recordset
Dim r10 As New Recordset
Dim r11 As New Recordset
Dim Adj As Double
If Len(lSC) = 0 Then
    MsgBox "Please select stock to adjust!", vbExclamation, "Status"
    lvw1.SetFocus
    Exit Sub
End If
If Len(txtAdj) = 0 Then
    MsgBox "Please select value to adjust stock to!", vbExclamation, "Status"
    txtAdj.SetFocus
    Exit Sub
End If
res = MsgBox("Are you sure you wish to adjust stock level?", vbYesNo + vbQuestion, "Adjust Stock?")
If res = vbNo Then Exit Sub

With r
.Open "select * from stock where stockcodemain='" & lSC & "'", c, adOpenDynamic, adLockOptimistic
End With
With r10
.Open "select * from serialnumber where stockcode='" & lSC & "'", c, adOpenDynamic, adLockOptimistic
End With

Adj = 0
Adj = CDbl(txtAdj) - r!qty
If r10.EOF = False Then
    If Adj > 0 Then
        If lSn.ListCount <> Adj Then
            res = MsgBox("Cannot adjust as serial numbers do not match adjustment value! This stock code requires you to add serial numbers in order to adjust! However do you wish to continue?", vbYesNo + vbQuestion, "Status")
            If res = vbNo Then
                Exit Sub
            End If
        End If
    Else
        lSn.Clear
    End If
    If lSn.ListCount > 0 Then
        With r11
        .Open "select * from serialnumber", c, adOpenDynamic, adLockOptimistic
        End With
        For i = 0 To lSn.ListCount - 1
            r11.AddNew
            r11!stockcode = lSC
            r11!serialnumber = lSn.List(i)
            r11.Update
        Next i
        r11.Close
    End If
    
End If
r10.Close


r!qty = CDbl(txtAdj)
r.Update
r.Close
With r
.Open "select * from stockadjustment", c, adOpenDynamic, adLockOptimistic
End With
r.AddNew
r!adjdate = Format(Date, "DD/MM/YYYY")
r!stockcode = lSC.Caption
r!adjustedto = txtAdj
r!adjustedby = Adj
r!adjustedfrom = CDbl(txtAdj) - Adj
r.Update
r.Close
res = MsgBox("Stock succesfully adjusted! Do you wish to adjust another?", vbYesNo + vbQuestion, "Adjust another?")
If res = vbYes Then
lvw1.ListItems.Clear
txtcode = ""
txtBarcode = ""
lSC = ""
ldesc = ""
lTAX = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lPACKED = ""
lUNPACKED = ""
txtAdj.Text = ""
txtDesc.Text = ""
lSn.Clear
txtsn.Text = ""

If optDesc.Value = True Then
txtBarcode.Enabled = False
txtcode.Enabled = False
txtDesc.Enabled = True
txtDesc.SetFocus
ElseIf OptStockCode.Value = True Then
txtBarcode.Enabled = False
txtcode.Enabled = True
txtDesc.Enabled = False

txtcode.SetFocus
ElseIf optBarcode.Value = True Then
txtBarcode.Enabled = True
txtcode.Enabled = False
txtDesc.Enabled = False

txtBarcode.SetFocus
End If

Else
Unload Me
End If



End Sub

Private Sub LaVolpeButton2_Click()
Unload Me

End Sub


Public Sub LaVolpeButton3_Click()
Dim r As New Recordset
If Len(txtsn) = 0 Then Exit Sub
If lvw1.ListItems.Count = 0 Then
    MsgBox "Please select stock code of item you wish to add serial#!", vbExclamation, "Status"
    txtsn.Text = ""
    If optDesc.Value = True Then
        txtBarcode.Enabled = False
        txtcode.Enabled = False
        txtDesc.Enabled = True
        txtDesc.SetFocus
    ElseIf OptStockCode.Value = True Then
        txtBarcode.Enabled = False
        txtcode.Enabled = True
        txtDesc.Enabled = False
        
        txtcode.SetFocus
    ElseIf optBarcode.Value = True Then
        txtBarcode.Enabled = True
        txtcode.Enabled = False
        txtDesc.Enabled = False
        
        txtBarcode.SetFocus
    End If
    Exit Sub
End If
If Len(lvw1.SelectedItem) = 0 Then
    MsgBox "Please select stock code of item you wish to add serial#!", vbExclamation, "Status"
    txtsn.Text = ""
    If optDesc.Value = True Then
        txtBarcode.Enabled = False
        txtcode.Enabled = False
        txtDesc.Enabled = True
        txtDesc.SetFocus
    ElseIf OptStockCode.Value = True Then
        txtBarcode.Enabled = False
        txtcode.Enabled = True
        txtDesc.Enabled = False
        
        txtcode.SetFocus
    ElseIf optBarcode.Value = True Then
        txtBarcode.Enabled = True
        txtcode.Enabled = False
        txtDesc.Enabled = False
        
        txtBarcode.SetFocus
    End If
    Exit Sub
End If
If lSn.ListCount > 0 Then
For i = 0 To lSn.ListCount - 1
    If lSn.List(i) = txtsn.Text Then
        MsgBox "Cannot add serial number as serial number already exist!", vbExclamation, "Status"
        txtsn.Text = ""
        txtsn.SetFocus
        Exit Sub
    End If
Next i
End If
With r
.Open "select * from serialnumber where serialnumber='" & txtsn & "'", c, adOpenDynamic, adLockOptimistic
End With
If r.EOF = True Then
    lSn.AddItem txtsn.Text
    txtsn.Text = ""
    txtsn.SetFocus
Else
    MsgBox "Cannot add serial number! This serial number already exists.", vbExclamation, "Status"
    txtsn.Text = ""
    txtsn.SetFocus
End If
r.Close

End Sub

Private Sub LaVolpeButton4_Click()
If lSn.ListCount = 0 Then Exit Sub
    If Len(lSn.List(lSn.ListIndex)) = 0 Then Exit Sub
    lSn.RemoveItem (lSn.ListIndex)
    txtsn.SetFocus
End Sub

Private Sub lvw1_Click()
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
'lsp = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lPACKED = ""
lUNPACKED = ""
ladj = ""

lSC = lvw1.SelectedItem



With r
.Open "select * from stock where stockcodemain='" & lvw1.SelectedItem & "'", c, adOpenDynamic, adLockOptimistic
End With


ldesc.Caption = UCase(r!stockdesc)
lTAX = r!taxcode
'lsp = Format(r!unitprice, "#####.00")
r.Close
With r
.Open "select * from stock where stockcodemain='" & lvw1.SelectedItem & "'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    qtyOnhand = qtyOnhand + CDbl(r!qty)
    r.MoveNext
Loop


lqty = qtyOnhand
r.Close

With r
.Open "select * from sales where itemcodemain='" & lvw1.SelectedItem & "'", c, adOpenDynamic, adLockOptimistic
End With
Do While r.EOF = False
    QTYSOLD = QTYSOLD + CDbl(r!qty)
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
'lcp = Format(TotalCP / NumofPurchase, "#####.00")
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

ladj.Caption = TotAdj


End Sub

Private Sub Option1_Click()

End Sub

Private Sub Option2_Click()

End Sub

Private Sub optBarcode_Click()
txtcode.Enabled = False
txtDesc.Enabled = False
txtBarcode.Enabled = True
'txtsn.Enabled = False

txtcode = ""
txtDesc = ""
lvw1.ListItems.Clear
lSC = ""
ldesc = ""
lTAX = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lPACKED = ""
lUNPACKED = ""
ladj = ""
txtBarcode.SetFocus

End Sub

Private Sub optDesc_Click()
txtDesc.Enabled = True
txtcode.Enabled = False
txtBarcode.Enabled = False
'txtsn.Enabled = False
txtcode = ""
txtBarcode = ""
lvw1.ListItems.Clear
lSC = ""
ldesc = ""
lTAX = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lPACKED = ""
lUNPACKED = ""
ladj = ""
txtDesc.SetFocus

End Sub

Private Sub optSN_Click()
txtcode.Enabled = False
txtDesc.Enabled = False
txtBarcode.Enabled = False
txtsn.Enabled = True

txtsn.SetFocus

End Sub

Private Sub OptStockCode_Click()
txtcode.Enabled = True
txtDesc.Enabled = False
txtBarcode.Enabled = False
'txtsn.Enabled = False
txtBarcode = ""
txtDesc = ""
lvw1.ListItems.Clear
lSC = ""
ldesc = ""
lTAX = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lPACKED = ""
lUNPACKED = ""
ladj = ""
txtcode.SetFocus

End Sub

Private Sub txtAdj_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If

End Sub


Private Sub txtBarcode_KeyPress(KeyAscii As Integer)
Dim search As String
Dim rs As New Recordset
lvw1.ListItems.Clear
lSC = ""
ldesc = ""
lTAX = ""
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lPACKED = ""
lUNPACKED = ""
ladj = ""

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
        Set li = lvw1.ListItems.Add(, , !stockcodemain)
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
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lPACKED = ""
lUNPACKED = ""
ladj = ""


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
        Set li = lvw1.ListItems.Add(, , !stockcodemain)
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
lqty = ""
lTOTPurch = ""
lsold = ""
lSUP = ""
lPdate = ""
lPACKED = ""
lUNPACKED = ""
ladj = ""

If KeyAscii <> 8 Then
    search = "%" & txtDesc.Text + Chr(KeyAscii) + "%"
Else
    If Len(txtDesc.Text) > 1 Then
    search = "%" & Mid(txtDesc.Text, 1, Len(txtDesc) - 1) + "%"
    Else
    Exit Sub
    End If
End If


With rs
.Open "select distinct(stockcodemain) from stock where stockdesc like '" & search & "' order by stockcodemain", c, adOpenDynamic, adLockOptimistic
If .EOF = False Then
    Do While .EOF = False
        Set li = lvw1.ListItems.Add(, , !stockcodemain)
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

Private Sub txtsn_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    LaVolpeButton3_Click
End If

End Sub


