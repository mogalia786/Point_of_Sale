VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCreatePackNew 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Pack from new stock"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "Parent Item Details"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   5535
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   6135
      Begin VB.TextBox tQTY 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         TabIndex        =   35
         Text            =   "1"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox tSC 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   34
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox tBC 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   33
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox tSN 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   32
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox tDESC 
         Appearance      =   0  'Flat
         Height          =   975
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   2640
         Width           =   4215
      End
      Begin VB.ComboBox cTaxCode 
         Height          =   315
         ItemData        =   "frmCreatePackNew.frx":0000
         Left            =   1680
         List            =   "frmCreatePackNew.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   3720
         Width           =   1215
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
         TabIndex        =   27
         Top             =   5040
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000016&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   26
         Top             =   4680
         Width           =   4455
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
         TabIndex        =   25
         Top             =   4200
         Width           =   1815
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
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker cDates 
         Height          =   375
         Left            =   1680
         TabIndex        =   28
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   55246849
         CurrentDate     =   38461
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton1 
         Height          =   375
         Left            =   3480
         TabIndex        =   30
         Top             =   1200
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
         MICON           =   "frmCreatePackNew.frx":0004
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
         TabIndex        =   46
         Top             =   720
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
         TabIndex        =   45
         Top             =   1200
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
         TabIndex        =   44
         Top             =   1680
         Width           =   1335
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
         TabIndex        =   43
         Top             =   2160
         Width           =   1575
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
         TabIndex        =   42
         Top             =   2520
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   3720
         Width           =   1335
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
         TabIndex        =   40
         Top             =   240
         Width           =   1335
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   39
         Top             =   5040
         Width           =   3735
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   38
         Top             =   4200
         Width           =   3015
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
         TabIndex        =   37
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "F1 to Update"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3480
         TabIndex        =   36
         Top             =   2280
         Width           =   1095
      End
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
      ForeColor       =   &H00400000&
      Height          =   4815
      Left            =   6360
      TabIndex        =   2
      Top             =   120
      Width           =   5295
      Begin VB.TextBox tSUP 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox tADD1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox tADD2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox tADD3 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox tADD4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox tTEL 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox tFAX 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   3240
         Width           =   1935
      End
      Begin VB.ComboBox cTerms 
         Height          =   315
         ItemData        =   "frmCreatePackNew.frx":0020
         Left            =   1800
         List            =   "frmCreatePackNew.frx":006C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox lBAL 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   3720
         Width           =   1935
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton2 
         Height          =   375
         Left            =   4200
         TabIndex        =   12
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
         MICON           =   "frmCreatePackNew.frx":00E4
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
      Begin LVbuttons.LaVolpeButton LaVolpeButton6 
         Height          =   375
         Left            =   4200
         TabIndex        =   13
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
         MICON           =   "frmCreatePackNew.frx":0100
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   4200
         Width           =   1575
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
         TabIndex        =   14
         Top             =   3720
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000016&
      Caption         =   "Child Items"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   5760
      Width           =   11535
      Begin VB.TextBox tCSC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   56
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox tCBC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   54
         Top             =   1080
         Width           =   3495
      End
      Begin LVbuttons.LaVolpeButton CmdAdd 
         Height          =   495
         Left            =   6360
         TabIndex        =   52
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
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
         COLTYPE         =   1
         BCOL            =   13160660
         FCOL            =   0
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmCreatePackNew.frx":011C
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
      Begin VB.TextBox txtChild 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   50
         Top             =   1680
         Width           =   3495
      End
      Begin VB.ListBox l1 
         Appearance      =   0  'Flat
         Height          =   2175
         Left            =   7920
         TabIndex        =   49
         Top             =   240
         Width           =   3495
      End
      Begin LVbuttons.LaVolpeButton LaVolpeButton8 
         Height          =   495
         Left            =   6360
         TabIndex        =   53
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "&Remove "
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
         MICON           =   "frmCreatePackNew.frx":0138
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
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Child Item Stock Code:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   57
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Child Item Barcode:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   55
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         Height          =   1455
         Left            =   6240
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Child Item Serial Number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   51
         Top             =   1680
         Width           =   2415
      End
   End
   Begin LVbuttons.LaVolpeButton LaVolpeButton7 
      Height          =   495
      Left            =   6600
      TabIndex        =   0
      Top             =   5160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
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
      MICON           =   "frmCreatePackNew.frx":0154
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
      Left            =   8280
      TabIndex        =   47
      Top             =   5160
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
      MICON           =   "frmCreatePackNew.frx":0170
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
      Left            =   9840
      TabIndex        =   48
      Top             =   5160
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
      MICON           =   "frmCreatePackNew.frx":018C
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
End
Attribute VB_Name = "frmCreatePackNew"
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

Public Sub CmdAdd_Click()
If Len(tCSC) = 0 Then
    MsgBox "Please enter child stock code!", vbExclamation, "status"
    tCSC.SetFocus
    Exit Sub
End If
If Len(tCBC) = 0 Then
    MsgBox "Please enter child Barcode!", vbExclamation, "status"
    tCBC.SetFocus
    Exit Sub
End If
If Len(txtChild.Text) = 0 Then Exit Sub
If l1.ListCount > 0 Then
    For i = 0 To l1.ListCount - 1
        If l1.List(i) = txtChild Then
            Exit Sub
        End If
    Next i
End If
l1.AddItem txtChild
txtChild.Text = ""
tCSC.Enabled = False
tCBC.Enabled = False
txtChild.SetFocus

    

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

End Sub

Private Sub LaVolpeButton1_Click()
SP = 0
frmItemSearchPack.Show vbModal

End Sub

Private Sub LaVolpeButton2_Click()
frmSupplierSearchPack.Show vbModal

End Sub

Public Sub LaVolpeButton3_Click()
Dim r As New Recordset
Dim r1 As New Recordset
Dim li As ListItem
Dim xTAX As Double

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
If Len(tSN) = 0 Then
    tSN = "N/A"
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
If Len(tSN.Text) > 0 Then
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
For i = 1 To lvw1.ListItems.Count
    If tSN.Text = lvw1.ListItems(i).SubItems(3) Then
        MsgBox "Cannot add item to list as there already exist an item with same serial number!", vbExclamation, "Status"
        tSN = ""
        tSN.SetFocus
        Exit Sub
    End If
Next i

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
End With
res = MsgBox("Do you wish to clear serial number and leave all other fields as they are?", vbYesNo + vbQuestion, "Status")
If res = vbNo Then
    tQTY.Text = ""
    tSC = ""
    tBC = ""
    tSN = ""
    tDESC = ""
    tCOST = ""
    tSell.Text = ""
    Check1.Value = False
    cTaxCode.ListIndex = -1
    
    tQTY.SetFocus
Else
    tSN = ""
    tSN.SetFocus
End If



    





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
Dim xxTOTAL As Double
If Len(tINVNO.Text) = 0 Then
    MsgBox "Please enter invoice number!", vbExclamation, "Status"
    tINVNO.SetFocus
    Exit Sub
End If
If l1.ListCount = 0 Then
    MsgBox "Please enter Child Items!", vbExclamation, "Status"
    txtChild.SetFocus
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
If Len(tSN) = 0 Then
    MsgBox "Please enter serial number!", vbExclamation, "status"
    tSN.SetFocus
    Exit Sub
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
    With r
    .Open "select * from stock where stockcodemain='" & tSC & "' and stockcode='" & tBC & "'", c, adOpenDynamic, adLockOptimistic
    End With
    If r.EOF = False Then
        r!QTY = CDbl(r!QTY) + 1
        r!taxcode = cTaxCode.Text
        r!unitprice = Format(tSell.Text, "#####.00")
        r.Update
    Else
        r.AddNew
        r!stockcodeMAIN = UCase(tSC)
        r!stockcode = tBC
        'r!serialnumber = lvw1.ListItems(i).SubItems(3)
        r!stockdesc = UCase(tDESC)
        r!QTY = "1"
        r!unitprice = Format(tSell, "#####.00")
        r!taxcode = cTaxCode.Text
        r.Update
    End If
    r.Close
    
    With r
    .Open "select * from stockpurchasehistory", c, adOpenDynamic, adLockOptimistic
    End With
    r.AddNew
    r!stockcodeMAIN = UCase(tSC)
    r!stockcode = tBC
    r!serialnumber = tSN
    r!supplier = UCase(tSUP)
    r!costofItemEXC = Format(tCOST, "#####.00")
    With r1
    .Open "select * from taxcode where taxcode='" & cTaxCode.Text & "'", c, adOpenDynamic, adLockOptimistic
    End With
    
    r!vatinput = Format(CDbl(tCOST) * (CDbl(r1!tax) / 100), "#####.00")
    r!costofiteminc = Format(CDbl(r!vatinput) + CDbl(tCOST), "#####.00")
    xxTOTAL = 0
    xxTOTAL = Format(r!costofiteminc, "#####.00")
    r1.Close
    r!sellingpriceINc = Format(tSell, "#####.00")
    r!datepurchased = Format(cDates.Value, "DD/MM/YYYY")
    r!qtypurchased = "1"
    r.Update
    r.Close


'For i = 1 To lvw1.ListItems.Count
    If tSN <> "N/A" Then
        With r
        .Open "select * from serialnumber", c, adOpenDynamic, adLockOptimistic
        End With
        r.AddNew
        r!stockcode = UCase(tSC)
        r!serialnumber = tSN
        r.Update
        r.Close
    End If
'Next i


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
    'xxTOTAL = 0
    'For i = 1 To lvw1.ListItems.Count
    '    xxTOTAL = xxTOTAL + (lvw1.ListItems(i) * CDbl(lvw1.ListItems(i).SubItems(7)))
    'Next i

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
    'xxTOTAL = 0
    'For i = 1 To lvw1.ListItems.Count
    '    xxTOTAL = xxTOTAL + (lvw1.ListItems(i) * CDbl(lvw1.ListItems(i).SubItems(7)))
    'Next i
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
.Open "select * from packS", c, adOpenDynamic, adLockOptimistic
End With
For i = 0 To l1.ListCount - 1
    r.AddNew
    r!PSTOCKCODE = UCase(tSC)
    r!pbarcode = tBC
    r!pserialnumber = tSN
    r!cstockcode = UCase(tCSC.Text)
    r!cbarcode = tCBC.Text
    r!cserialnumber = l1.List(i)
    r.Update
Next i
    
res = MsgBox("Invoice successfully added! Do you wish to add another?", vbYesNo + vbQuestion, "Add Another?")
If res = vbNo Then
    SP = 0
    Unload Me
Else
lvw1.ListItems.Clear
    SP = 0
    tQTY.Text = ""
    tSC = ""
    tBC = ""
    tSN = ""
    tDESC = ""
    cTaxCode.ListIndex = -1
    tCOST = ""
    tSell.Text = ""
    tSUP.Text = ""
    tADD1.Text = ""
    tADD2.Text = ""
    tADD3.Text = ""
    tADD4.Text = ""
    tTEL.Text = ""
    tFAX.Text = ""
    cTerms.ListIndex = -1
    Check1.Value = False
    tQTY.SetFocus
End If

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


Private Sub tQTY_KeyPress(KeyAscii As Integer)

If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If

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


Private Sub tSN_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    res = MsgBox("Update entry without updating any other fields?", vbYesNo + vbQuestion, "Update?")
    If res = vbYes Then
        Me.LaVolpeButton3_Click
    End If
End If


End Sub



Private Sub txtChild_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Me.CmdAdd_Click
End If

End Sub
