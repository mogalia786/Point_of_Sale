VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmChange 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change (Spacebar to Exit)"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   3825
   StartUpPosition =   2  'CenterScreen
   Begin MSCommLib.MSComm ms1 
      Left            =   1680
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label lchange 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1000.00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim r As New Recordset
If KeyCode = vbKeySpace Then
        With r
        .Open "select * from poledisplay", c, adOpenDynamic, adLockOptimistic
        End With
        ms1.PortOpen = True
        ms1.Output = Chr(27) + "@"
        ms1.Output = CStr(r!Line1) & vbCrLf
        ms1.Output = CStr(r!Line2)
        ms1.PortOpen = False
        r.Close
        Unload Me
End If

End Sub

Private Sub Form_Load()
lchange.Caption = "R" & MyChange
ms1.PortOpen = True
        ms1.Output = Chr(27) + "@"
ms1.Output = "Change" & vbCrLf
ms1.Output = "            R" & MyChange
    ms1.PortOpen = False
'OpenTill


End Sub

