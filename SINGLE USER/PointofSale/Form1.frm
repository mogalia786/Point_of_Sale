VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin LVbuttons.LaVolpeButton PC 
      Height          =   255
      Left            =   9000
      TabIndex        =   34
      Top             =   6780
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BTYPE           =   2
      TX              =   "Price Change ON"
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
      BCOL            =   65280
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Form1.frx":0000
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
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   5400
      Top             =   3480
   End
   Begin VB.TextBox txtDisc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4440
      TabIndex        =   28
      Text            =   "0.00"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSCommLib.MSComm ms1 
      Left            =   5760
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00000000&
      Caption         =   "Returns"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   9000
      TabIndex        =   26
      Top             =   7920
      Width           =   1095
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "F8"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      Caption         =   "Discount"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   7800
      TabIndex        =   24
      Top             =   7920
      Width           =   1095
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "F7"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5400
      Top             =   4320
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00000000&
      Caption         =   "Reset / New Sale"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   10200
      TabIndex        =   19
      Top             =   7920
      Width           =   1695
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "End"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   480
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      Caption         =   "Sale"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   6600
      TabIndex        =   17
      Top             =   7920
      Width           =   1095
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "F5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox txtqty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   960
      MaxLength       =   3
      TabIndex        =   15
      Text            =   "1"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "Qty Change"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   1320
      TabIndex        =   13
      Top             =   7920
      Width           =   1095
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "F1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   0
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
   End
   Begin RichTextLib.RichTextBox T1 
      Height          =   5175
      Left            =   480
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1560
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   9128
      _Version        =   393217
      BackColor       =   16777215
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Form1.frx":001C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2880
      TabIndex        =   8
      Top             =   120
      Width           =   9135
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Sale Adjustment"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   4920
      TabIndex        =   6
      Top             =   7920
      Width           =   1575
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "F4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   480
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Item Search"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   3600
      TabIndex        =   4
      Top             =   7920
      Width           =   1215
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "F3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Tender"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   2520
      TabIndex        =   2
      Top             =   7920
      Width           =   975
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "F2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Crit. Toggle"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   7920
      Width           =   1095
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "F9"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Line Line3 
      X1              =   2400
      X2              =   2400
      Y1              =   6720
      Y2              =   7080
   End
   Begin VB.Line Line2 
      X1              =   6840
      X2              =   6840
      Y1              =   6720
      Y2              =   7080
   End
   Begin VB.Line Line1 
      X1              =   10680
      X2              =   10680
      Y1              =   6720
      Y2              =   7080
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "F11 - ON / OFF Toggle"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   6960
      TabIndex        =   35
      Top             =   6800
      Width           =   1935
   End
   Begin VB.Label lSALE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sale"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2400
      TabIndex        =   33
      Top             =   6720
      Width           =   4455
   End
   Begin VB.Label lCRIT 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Barcode"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   32
      Top             =   200
      Width           =   2895
   End
   Begin VB.Label lD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Disc-0%"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   615
      Left            =   9600
      TabIndex        =   31
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Lq 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Qty-1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   7080
      TabIndex        =   30
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label ldisc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Disc(%)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2640
      TabIndex        =   29
      Top             =   7200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lTeller 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10200
      TabIndex        =   23
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label lDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   0
      Top             =   6720
      Width           =   12015
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   61
      Left            =   11640
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   62
      Left            =   11640
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   63
      Left            =   11640
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   66
      Left            =   11640
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   64
      Left            =   11640
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   65
      Left            =   11640
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   71
      Left            =   11640
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   67
      Left            =   11640
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   68
      Left            =   11640
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   69
      Left            =   11640
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   70
      Left            =   11640
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   59
      Left            =   120
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   58
      Left            =   120
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   57
      Left            =   120
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   56
      Left            =   120
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   60
      Left            =   120
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   54
      Left            =   120
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   53
      Left            =   120
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   55
      Left            =   120
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   52
      Left            =   120
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   51
      Left            =   120
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   50
      Left            =   120
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label txtdesc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scan Item Code"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   0
      TabIndex        =   21
      Top             =   960
      Width           =   7095
   End
   Begin VB.Label lqty 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   -120
      TabIndex        =   16
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      Top             =   7800
      Width           =   12015
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   9000
      TabIndex        =   12
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label lTOTAL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   9480
      TabIndex        =   11
      Top             =   7200
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Due:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6240
      TabIndex        =   10
      Top             =   7200
      Width           =   2895
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   5175
      Left            =   0
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   5175
      Left            =   11520
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   9000
      Top             =   7200
      Width           =   2895
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Trans As String
Dim PTrim As Boolean
Public PCOn As Boolean



Private Sub Form_Load()
Dim cs As New Connection
Dim r As New Recordset
lineN = 1
ConnectMe2
PCOn = True
PTrim = True
lDate = Now
lTeller = "Teller: " & CurrentUser
Trans = "Sale"
With cs
.ConnectionString = App.Path & "\Temp.mdb"
.Provider = "microsoft.jet.oledb.4.0"
.Open
End With
With r
.Open "Till", cs, adOpenDynamic, adLockOptimistic
End With
TillId = r!TillId
r.Close
'ms1.PortOpen = True
With r
.Open "select * from poledisplay", c, adOpenDynamic, adLockOptimistic
End With
        ms1.PortOpen = True
        ms1.Output = Chr(27) + "@"
        ms1.Output = CStr(r!Line1) & vbCrLf
        ms1.Output = CStr(r!Line2)
    
        
        ms1.PortOpen = False
r.Close
End Sub

Private Sub LaVolpeButton1_Click()

End Sub

Public Sub PC_Click()
If PC.Caption = "Price Change ON" Then
    PC.Caption = "Price Change OFF"
    PC.BackColor = vbRed
    PCOn = False
Else
    PC.Caption = "Price Change ON"
    PC.BackColor = &HFF00&
    PCOn = True
End If
txtcode.SetFocus
End Sub

Private Sub T1_GotFocus()
txtcode.SetFocus
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_GotFocus()

End Sub

Private Sub Timer1_Timer()
lDate = Now
lTeller = "Teller: " & CurrentUser

End Sub

Private Sub Timer2_Timer()
lSALE.Caption = Trans

End Sub

Private Sub txtcode_KeyPress(KeyAscii As Integer)
'If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
'KeyAscii = KeyAscii
'Else
'KeyAscii = 0
'End If

End Sub

Private Sub txtcode_KeyUp(KeyCode As Integer, Shift As Integer)
Dim r As New Recordset
Dim R1 As New Recordset
Dim r2 As New Recordset
Dim TOTDISC As Double
Dim UNITDISC As Double
Dim STOCKDESC As String
Dim TAXCode1 As Double
Dim DISCOUNT As Double
Dim DiscountCode As String
Dim UNITPRICE As Double
If KeyCode = vbKeyF8 Then
    frmSupervisor.Show vbModal
End If
If KeyCode = vbKeyEscape Then
    frmOpenDraw.Show vbModal
End If
If KeyCode = vbKeyF11 Then
    Me.PC_Click
End If


If KeyCode = vbKeyF12 Then
    res = MsgBox("Are you sur you wish to exit?", vbYesNo + vbQuestion, "Exit?")
    If res = vbYes Then
        With r
        .Open "delete * from sale", HC, adOpenDynamic, adLockOptimistic
        End With
        ms1.PortOpen = True

        ms1.Output = Chr(27) + "@"
        ms1.PortOpen = False

        End
    End If
End If
If KeyCode = vbKeyF1 Then
    txtqty.Visible = True
    lqty.Visible = True
    txtqty.SetFocus
End If
If KeyCode = vbKeyF3 Then
    frmItemSearch.Show vbModal
End If
If KeyCode = vbKeyF4 Then
    Trans = "Adjust"
    txtDesc = "Adjustment"
End If
If KeyCode = vbKeyF5 Then
    Trans = "Sale"
    txtDesc = "Next Item"
End If
If KeyCode = vbKeyEnd Then
    T1.Text = ""
    With r
    .Open "delete * from sale", HC, adOpenDynamic, adLockOptimistic
    End With
    Trans = "Sale"
    txtDesc = "Scan / Enter item code"
    txtcode.Text = ""
    lineN = 1
    lTOTAL = "0.00"
    txtqty.Text = "1"
    ms1.PortOpen = True

    ms1.Output = Chr(27) + "@"
    ms1.PortOpen = False

    txtcode.SetFocus
    Exit Sub
End If

If KeyCode = vbKeyF2 Then
    ms1.PortOpen = True
    ms1.Output = Chr(27) + "@"
    ms1.Output = "TOTAL DUE" & vbCrLf
    ms1.Output = "            R" & Format(lTOTAL, "####.00")
    ms1.PortOpen = False
    frmTender.Show vbModal
End If
If KeyCode = vbKeyF7 Then
    txtDisc.Visible = True
    ldisc.Visible = True
    txtDisc.SetFocus
End If
If KeyCode = vbKeyF9 Then
    If lCRIT = "Barcode" Then
        lCRIT = "Serial#"
    
    ElseIf lCRIT = "Serial#" Then
        lCRIT = "Stock Code"
    
    ElseIf lCRIT = "Stock Code" Then
        lCRIT = "Barcode"
    End If
End If



If KeyCode = vbKeyReturn Then
    If Len(txtcode.Text) = 0 Then Exit Sub
    If lCRIT <> "Serial#" Then
        If lCRIT = "Barcode" Then
            With r
            .Open "select * from stock where stockcode='" & txtcode.Text & "'", c, adOpenDynamic, adLockOptimistic
            End With
            If r.EOF = True Then
                txtDesc.Caption = "Item not found!"
                txtcode.Text = ""
                txtqty.Text = "1"
                txtDisc = "0"
                txtcode.SetFocus
                Exit Sub
            End If
            
            With R1
            .Open "select * from serialnumber where stockcode = '" & r!stockcodemain & "'", c, adOpenDynamic, adLockOptimistic
            End With
            If R1.EOF = False Then
                MsgBox "Please scan serial number only!", vbCritical, "Status"
                lCRIT.Caption = "Serial#"
                txtcode.Text = ""
                txtcode.SetFocus
                R1.Close
                r.Close
                Exit Sub
            Else
                R1.Close
                r.Close
            End If
        Else
            With R1
            .Open "select * from serialnumber where stockcode = '" & txtcode.Text & "'", c, adOpenDynamic, adLockOptimistic
            End With
            If R1.EOF = False Then
                MsgBox "Please scan serial number only!", vbCritical, "Status"
                lCRIT.Caption = "Serial#"
                txtcode.Text = ""
                txtcode.SetFocus
                R1.Close
                
                Exit Sub
            Else
                R1.Close
                
            End If
        End If
    End If
If lCRIT = "Serial#" Then
    If CInt(txtqty.Text) > 1 Then
        MsgBox "Quantity cannot exceed 1!", vbExclamation, "Status"
        txtqty.Text = "1"
        txtcode.Text = ""
        txtcode.SetFocus
        Exit Sub
    End If
End If

If Trans <> "Adjust" And PCOn = True Then
    frmPriceView.Show vbModal
    txtcode.SetFocus
End If

    If lCRIT = "Barcode" Then
        With r
        .Open "select * from stock where stockcode='" & txtcode.Text & "'", c, adOpenDynamic, adLockOptimistic
        End With
    End If
    
    
    If lCRIT = "Serial#" Then
        With R1
        .Open "select * from serialnumber where serialnumber='" & txtcode.Text & "'", c, adOpenDynamic, adLockOptimistic
        End With
        If R1.EOF = False Then
            With r
            .Open "select * from stock where stockcodemain='" & R1!stockcode & "'", c, adOpenDynamic, adLockOptimistic
            End With
            R1.Close
            With R1
            .Open "select * from sale where itemcode='" & r!stockcode & "' and itemcodemain='" & r!stockcodemain & "' and serialnumber='" & txtcode.Text & "'", HC, adOpenDynamic, adLockOptimistic
            End With
            If R1.EOF = False And Trans <> "Adjust" Then
                MsgBox "Cannot scan serial number twice!", vbCritical, "Cannot scan serial!"
                txtcode.Text = ""
                txtcode.SetFocus
                Exit Sub
            End If
            R1.Close
        Else
            txtDesc = "Item not found!"
            txtcode.Text = ""
            txtqty.Text = "1"
            txtDisc.Text = "0"

            Exit Sub
        End If
            
    End If
    If lCRIT = "Stock Code" Then
        With r
        .Open "select * from stock where stockcodemain='" & txtcode.Text & "'", c, adOpenDynamic, adLockOptimistic
        End With
    End If

    If r.EOF = False Then
        If lCRIT = "Serial#" Then
            serialn = txtcode
        Else
            serialn = "N/A"
        End If

        With R1
        .Open "select * from sale where itemcode='" & r!stockcode & "' and itemcodemain='" & r!stockcodemain & "' and serialnumber='" & serialn & "'", HC, adOpenDynamic, adLockOptimistic
        End With
        If R1.EOF = False Then
            txtDisc.Text = R1!discadded
        End If
        R1.Close
        serialn = ""
        If Len(r!STOCKDESC) < 21 Then
            If Len(r!STOCKDESC) = 1 Then
                STOCKDESC = r!STOCKDESC & "                    "
            End If
            If Len(r!STOCKDESC) = 2 Then
                STOCKDESC = r!STOCKDESC & "                   "
            End If
            If Len(r!STOCKDESC) = 3 Then
                STOCKDESC = r!STOCKDESC & "                  "
            End If
            If Len(r!STOCKDESC) = 4 Then
                STOCKDESC = r!STOCKDESC & "                 "
            End If
            If Len(r!STOCKDESC) = 5 Then
                STOCKDESC = r!STOCKDESC & "                "
            End If
            If Len(r!STOCKDESC) = 6 Then
                STOCKDESC = r!STOCKDESC & "               "
            End If
            If Len(r!STOCKDESC) = 7 Then
                STOCKDESC = r!STOCKDESC & "              "
            End If
            If Len(r!STOCKDESC) = 8 Then
                STOCKDESC = r!STOCKDESC & "             "
            End If
            If Len(r!STOCKDESC) = 9 Then
                STOCKDESC = r!STOCKDESC & "            "
            End If
            If Len(r!STOCKDESC) = 10 Then
                STOCKDESC = r!STOCKDESC & "           "
            End If
            If Len(r!STOCKDESC) = 11 Then
                STOCKDESC = r!STOCKDESC & "          "
            End If
            If Len(r!STOCKDESC) = 12 Then
                STOCKDESC = r!STOCKDESC & "         "
            End If
            If Len(r!STOCKDESC) = 13 Then
                STOCKDESC = r!STOCKDESC & "        "
            End If
            If Len(r!STOCKDESC) = 14 Then
                STOCKDESC = r!STOCKDESC & "       "
            End If
            If Len(r!STOCKDESC) = 15 Then
                STOCKDESC = r!STOCKDESC & "      "
            End If
            If Len(r!STOCKDESC) = 16 Then
                STOCKDESC = r!STOCKDESC & "     "
            End If
            If Len(r!STOCKDESC) = 17 Then
                STOCKDESC = r!STOCKDESC & "    "
            End If
            If Len(r!STOCKDESC) = 18 Then
                STOCKDESC = r!STOCKDESC & "   "
            End If
            If Len(r!STOCKDESC) = 19 Then
                STOCKDESC = r!STOCKDESC & "  "
            End If

            If Len(r!STOCKDESC) = 20 Then
                STOCKDESC = r!STOCKDESC & " "
            End If
        Else
            STOCKDESC = Mid(r!STOCKDESC, 1, 20) & " "
        End If
        
        If Trans = "Adjust" Then
            If lCRIT = "Barcode" Then
                With r2
                .Open "select * from sale where itemcode='" & txtcode.Text & "' and itemcodemain='" & r!stockcodemain & "'", HC, adOpenDynamic, adLockOptimistic
                End With
            End If
            If lCRIT = "Stock Code" Then
                With r2
                .Open "select * from sale where itemcodemain='" & txtcode.Text & "' and itemcode='" & r!stockcode & "'", HC, adOpenDynamic, adLockOptimistic
                End With
            End If
            If lCRIT = "Serial#" Then
                With r2
                .Open "select * from sale where serialnumber='" & txtcode.Text & "' and itemcodemain='" & r!stockcodemain & "' and itemcode='" & r!stockcode & "'", HC, adOpenDynamic, adLockOptimistic
                End With
            End If

            If r2.EOF = True Then
                txtDesc = "Item cannot adjust!"
                txtcode.Text = ""
                txtqty.Text = "1"
                Trans = "Sale"
                Exit Sub
            End If
            If r2!qty < CInt(txtqty.Text) Then
                txtDesc = "Item cannot adjust!"
                txtcode.Text = ""
                txtqty.Text = "1"
                Trans = "Sale"
                Exit Sub
            End If
            vpi = Format(CDbl(r2!vat) / CDbl(r2!qty), "#####.00")
            
            r2!vat = r2!vat - (CDbl(txtqty.Text) * CDbl(vpi))
            r2!qty = r2!qty - txtqty.Text
            r2!total = Format(r2!qty * CDbl(r2!UNITPRICE), "####.00")
            r2!TOTDISC = CDbl(r2!TOTDISC) - (CDbl(txtqty) * CDbl(r2!UNITDISC))
            

            r2.Update
            If r2!qty <= 0 Then
                r2.Delete
            End If
            r2.Close
        End If
        If lCRIT = "Stock Code" Then
            With r2
            .Open "select * from stockdiscount where stockcodemain='" & txtcode & "' and fromdate<='" & Year(Date) & "-" & Month(Date) & "-" & Day(Date) & " 00:00:00.000 ' and todate>='" & Year(Date) & "-" & Month(Date) & "-" & Day(Date) & " 00:00:00.000 '", c, adOpenDynamic, adLockOptimistic
            End With
        End If
        If lCRIT = "Barcode" Then
            With r2
            .Open "select * from stockdiscount where stockcode='" & txtcode & "' and fromdate<='" & Year(Date) & "-" & Month(Date) & "-" & Day(Date) & " 00:00:00.000 ' and todate>='" & Year(Date) & "-" & Month(Date) & "-" & Day(Date) & " 00:00:00.000 '", c, adOpenDynamic, adLockOptimistic
            End With
        End If
        If lCRIT = "Serial#" Then
            With R1
            .Open "select * from serialnumber where serialnumber='" & txtcode & "'", c, adOpenDynamic, adLockOptimistic
            End With

            With r2
            .Open "select * from stockdiscount where stockcodemain='" & R1!stockcode & "' and fromdate<='" & Year(Date) & "-" & Month(Date) & "-" & Day(Date) & " 00:00:00.000 ' and todate>='" & Year(Date) & "-" & Month(Date) & "-" & Day(Date) & " 00:00:00.000 '", c, adOpenDynamic, adLockOptimistic
            End With
            R1.Close
        End If

        If r2.EOF = False Then
            If lCRIT = "Serial#" Then
                prefixs = "(SN)"
            End If
            If lCRIT = "Barcode" Then
                prefixs = "(BC)"
            End If
            If lCRIT = "Stock Code" Then
                prefixs = "(SC)"
            End If

            T1.Text = T1.Text & txtcode.Text & prefixs & "  " & STOCKDESC & "     " & r!taxcode & "     " & r2!disccode & vbCrLf
            DISCOUNT = CDbl(r2!disc)
            DiscountCode = r2!disccode

        Else
            If lCRIT = "Serial#" Then
                prefixs = "(SN)"
            End If
            If lCRIT = "Barcode" Then
                prefixs = "(BC)"
            End If
            If lCRIT = "Stock Code" Then
                prefixs = "(SC)"
            End If

            T1.Text = T1.Text & txtcode.Text & prefixs & "  " & STOCKDESC & "     " & r!taxcode & "     0" & vbCrLf
            DISCOUNT = 0
            DiscountCode = "0"
        End If
        
        r2.Close
        TOTDISC = 0
        UNITPRICE = r!UNITPRICE - (r!UNITPRICE * DISCOUNT / 100)
         UNITPRICE = UNITPRICE - (UNITPRICE * txtDisc / 100)
        UNITDISC = Format(CDbl(r!UNITPRICE - UNITPRICE), "#####.00")
        TOTDISC = Format(CDbl(txtqty) * CDbl(r!UNITPRICE - UNITPRICE), "#####.00")
        With r2
        .Open "select * from taxcode where taxcode='" & r!taxcode & "'", c, adOpenDynamic, adLockOptimistic
        End With
        TAXCode1 = 100 / (100 + r2!tax)
        r2.Close
        If Trans = "Adjust" Then
            T1.Text = T1.Text & txtqty & "          " & Format(UNITPRICE, "####.00") & "         " & Format(CDbl(Format(UNITPRICE, "####.00")) * TAXCode1, "####.00") & "    " & TOTDISC & "(" & txtDisc & "%)            " & Format((CDbl(UNITPRICE) * CDbl(txtqty)) * -1, "####.00") & vbCrLf
        Else
            T1.Text = T1.Text & txtqty & "          " & Format(UNITPRICE, "####.00") & "         " & Format(CDbl(Format(UNITPRICE, "####.00")) * TAXCode1, "####.00") & "    " & TOTDISC & "(" & txtDisc & "%)            " & Format(CDbl(UNITPRICE) * CDbl(txtqty), "####.00") & vbCrLf
        End If
        '///////////////DISPLAY//////////
        If Trans <> "Adjust" Then
            ms1.PortOpen = True
            ms1.Output = Chr(27) + "@"
            If Len(STOCKDESC) > 14 Then
            ms1.Output = txtqty & "X" & "  " & Mid(STOCKDESC, 1, 14) & vbCrLf
            Else
            ms1.Output = txtqty & "X" & "  " & STOCKDESC & vbCrLf
            End If
            ms1.Output = "            R" & Format(UNITPRICE * txtqty, "#####.00")
            ms1.PortOpen = False
        Else
            ms1.PortOpen = True
            ms1.Output = Chr(27) + "@"
            If Len(STOCKDESC) > 14 Then
            ms1.Output = "ADJ." & "  " & Mid(STOCKDESC, 1, 14) & vbCrLf
            Else
            ms1.Output = "ADJ." & "  " & STOCKDESC & vbCrLf
            End If
            ms1.Output = "            -R" & Format(UNITPRICE * txtqty, "#####.00")
            ms1.PortOpen = False
        End If
        If Trans = "Adjust" Then
            T1.Text = T1.Text & "*" & lineN & "*---ADJUSTMENT----------------------------------------------" & vbCrLf
        Else
            T1.Text = T1.Text & "*" & lineN & "*-----------------------------------------------------------" & vbCrLf
        End If
        
        
        
        T1.Text = T1.Text + vbCrLf
        If Trans = "Adjust" Then
            T1.Find "*" & lineN & "*---ADJUSTMENT----------------------------------------------"
        Else
            T1.Find "*" & lineN & "*-----------------------------------------------------------"
        End If
        
        If PTrim = True Then
            For i = 50 To 71 Step 2
                Shape3(i).FillColor = vbBlack
            Next i
            For i = 51 To 71 Step 2
                Shape3(i).FillColor = vbWhite
            Next i
            PTrim = False
        Else
            For i = 50 To 71 Step 2
                Shape3(i).FillColor = vbWhite
            Next i
            For i = 51 To 71 Step 2
                Shape3(i).FillColor = vbBlack
            Next i
            PTrim = True
        End If
        
                
            
        
        txtDesc = r!STOCKDESC
        lineN = lineN + 1
        If lCRIT = "Serial#" Then
            serialn = txtcode
        Else
            serialn = "N/A"
        End If
        txtcode.Text = ""

        With R1
        .Open "select * from sale where itemcode='" & r!stockcode & "' and itemcodemain='" & r!stockcodemain & "' and serialnumber='" & serialn & "'", HC, adOpenDynamic, adLockOptimistic
        End With
        If Trans <> "Adjust" Then
            If R1.EOF = True Then
                R1.AddNew
                R1!itemcodemain = r!stockcodemain
                If lCRIT.Caption = "Serial#" Then
                    R1!serialnumber = serialn
                Else
                    R1!serialnumber = "N/A"
                End If
                R1!itemcode = r!stockcode
                R1!Description = r!STOCKDESC
                R1!qty = txtqty.Text
                R1!taxcode = r!taxcode
                R1!disccode = DiscountCode
                R1!UNITPRICE = Format(UNITPRICE, "####.00")
                R1!total = Format(R1!qty * CDbl(UNITPRICE), "####.00")
                R1!vat = CDbl(txtqty) * (CDbl(UNITPRICE) - Format(CDbl(Format(UNITPRICE, "####.00")) * TAXCode1, "####.00"))
                R1!discadded = txtDisc
                R1!TOTDISC = TOTDISC
                R1!UNITDISC = UNITDISC
                R1.Update
            Else
                R1!itemcodemain = r!stockcodemain
                If lCRIT.Caption = "Serial#" Then
                    R1!serialnumber = serialn
                Else
                    R1!serialnumber = "N/A"
                End If
                R1!itemcode = r!stockcode
                R1!Description = r!STOCKDESC
                R1!qty = R1!qty + txtqty.Text
                R1!taxcode = r!taxcode
                R1!disccode = DiscountCode
                R1!UNITPRICE = Format(UNITPRICE, "####.00")
                R1!total = Format(R1!qty * CDbl(UNITPRICE), "####.00")
                R1!vat = CDbl(R1!vat) + (CDbl(txtqty) * (CDbl(UNITPRICE) - Format(CDbl(Format(UNITPRICE, "####.00")) * TAXCode1, "####.00")))
                R1!discadded = txtDisc
                R1!TOTDISC = R1!TOTDISC + TOTDISC
                R1!UNITDISC = UNITDISC
                R1.Update
            End If
            
        End If
        If Trans = "Adjust" Then
            lTOTAL = Format(CDbl(lTOTAL) - (CDbl(txtqty.Text) * CDbl(UNITPRICE)), "#####.00")
        Else
            lTOTAL = Format(CDbl(lTOTAL) + (CDbl(txtqty.Text) * CDbl(UNITPRICE)), "#####.00")
        End If
        R1.Close
        r.Close
        txtqty.Text = "1"
        txtDisc.Text = "0"
        Trans = "Sale"
    Else
        txtDesc = "Item not found!"
        txtcode.Text = ""
        txtqty.Text = "1"
        txtDisc.Text = "0"
    End If
End If

End Sub

Private Sub txtDisc_Change()
lD.Caption = "Disc-" & txtDisc.Text & "%"
    

End Sub

Private Sub txtDisc_GotFocus()
txtDisc.SelStart = 1 - 1
txtDisc.SelLength = Len(txtDisc.Text)

End Sub

Private Sub txtDisc_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    A = InStr(1, txtDisc, ".")
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

Private Sub txtDisc_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If Len(txtDisc.Text) = 0 Then
        txtDisc.Text = "0"
    End If
    If CDbl(txtDisc.Text) > 100 Then
        MsgBox "Discount % cannot exceed 100%!", vbExclamation, "Status!"
        txtDisc = "0"
        Exit Sub
    End If
    
    txtcode.SetFocus
    txtDisc.Visible = False
    ldisc.Visible = False
End If

End Sub

Private Sub txtqty_Change()
Lq.Caption = "Qty-" & txtqty.Text

End Sub

Private Sub txtqty_GotFocus()
txtqty.SelStart = 1 - 1
txtqty.SelLength = Len(txtqty.Text)

End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If

End Sub

Private Sub txtqty_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If Len(txtqty.Text) = 0 Then
        txtqty.Text = "1"
    End If
    txtcode.SetFocus
    txtqty.Visible = False
    lqty.Visible = False
End If

End Sub
