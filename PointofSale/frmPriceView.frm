VERSION 5.00
Begin VB.Form frmPriceView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Confirm Qty and Price by pressing enter"
   ClientHeight    =   600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtqty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3960
      TabIndex        =   1
      Text            =   "1"
      Top             =   75
      Width           =   855
   End
   Begin VB.TextBox txtPrice 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6840
      TabIndex        =   2
      Top             =   75
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Confirm Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   3240
      X2              =   3240
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Label lUnitP 
      Alignment       =   2  'Center
      Caption         =   "Unit Price - R100.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmPriceView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public UnitPrice2 As Double
Public LOADME As Boolean
Private Sub Form_Load()
Dim r As New Recordset
Dim R1 As New Recordset
LOADME = True
    If frmMain.lCRIT = "Barcode" Then
        With r
        .Open "select * from stock where stockcode='" & frmMain.txtcode.Text & "'", c, adOpenDynamic, adLockOptimistic
        End With
        If r.EOF = False Then
            UnitPrice2 = Format(r!UNITPRICE, "#####.00")
        Else
            frmMain.txtDesc = "Item not found!"
            LOADME = False
        End If
        
        
    End If
    
    
    If frmMain.lCRIT = "Serial#" Then
        With R1
        .Open "select * from serialnumber where serialnumber='" & frmMain.txtcode.Text & "'", c, adOpenDynamic, adLockOptimistic
        End With
        If R1.EOF = False Then
            With r
            .Open "select * from stock where stockcodemain='" & R1!stockcode & "'", c, adOpenDynamic, adLockOptimistic
            End With
            If r.EOF = False Then
                UnitPrice2 = Format(r!UNITPRICE, "#####.00")
            Else
                frmMain.txtDesc = "Item not found!"
                LOADME = False
            End If

            R1.Close
        Else
            frmMain.txtDesc = "Item not found!"
            LOADME = False
        End If
            
    End If
    
    If frmMain.lCRIT = "Stock Code" Then
        With r
        .Open "select * from stock where stockcodemain='" & frmMain.txtcode.Text & "'", c, adOpenDynamic, adLockOptimistic
        End With
        If r.EOF = False Then
            UnitPrice2 = Format(r!UNITPRICE, "#####.00")
        Else
            frmMain.txtDesc = "Item not found!"
            LOADME = False
        End If

    End If
lUnitP = "Unit Price - R" & UnitPrice2
txtPrice.Text = UnitPrice2
txtPrice.SelStart = 0
txtPrice.SelLength = Len(txtPrice.Text)
txtqty.SelStart = 0
txtqty.SelLength = Len(txtqty.Text)

End Sub

Private Sub txtPrice_GotFocus()
If LOADME = False Then
    frmMain.txtcode = ""
    frmMain.txtDisc = "00.00"
    frmMain.txtqty = "1"
    Unload Me
End If

End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then
    A = InStr(1, txtPrice, ".")
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

Private Sub txtPrice_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If Len(txtqty.Text) = 0 Then
        MsgBox "Please enter qty!", vbCritical, "Status"
        txtqty.SetFocus
        Exit Sub
    End If

    If Len(txtPrice.Text) = 0 Then
        MsgBox "Please confirm price!", vbCritical, "Status"
        txtPrice.SetFocus
        Exit Sub
    End If
        frmMain.txtDisc = Round((UnitPrice2 - CDbl(txtPrice)) / UnitPrice2 * 100, 2)
        frmMain.txtqty = txtqty.Text
        Unload Me
    
End If

        
End Sub

Private Sub txtqty_GotFocus()
If LOADME = False Then
    frmMain.txtcode = ""
    frmMain.txtDisc = "00.00"
    frmMain.txtqty = "1"
    Unload Me
End If

End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 8 Or KeyAscii = 46 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
End If

End Sub

Private Sub txtqty_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If Len(txtqty) = 0 Then
        txtqty.SetFocus
    Else
        txtPrice.SetFocus
    End If
End If

End Sub

Private Sub txtqty_LostFocus()
If frmMain.lCRIT = "Serial#" And Len(txtqty.Text) > 0 Then
    If CInt(txtqty.Text) > 1 Then
        MsgBox "Quantity cannot exceed 1!", vbExclamation, "Status"
        txtqty.Text = "1"
        txtqty.SelStart = 0
        txtqty.SelLength = 1
        txtqty.SetFocus
    End If
End If

End Sub
