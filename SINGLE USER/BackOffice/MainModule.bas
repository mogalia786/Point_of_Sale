Attribute VB_Name = "MainModule"
Public CompNum(1 To 100) As String
Public numComp As Integer
Public TillId As String
Public CardNo As String
Public Trans As String
Public PrinterConn As String
Public PortNo As String
Public ShiftStartTime As Date
Public ShiftEndTime As Date
Public NumofCards As Integer
Public ReaderPort As Integer
Public SP As Double
Public CompName As String





Public Sub Main()
Dim r As New Recordset
Dim cs As New Connection
'With cs
'.ConnectionString = App.Path & "/Temp.mdb"
'.Provider = "microsoft.jet.oledb.4.0"
'.Open
'End With
'With r
'.Open "setup", cs, adOpenDynamic, adLockOptimistic
'End With
'If r!initialized = "No" Then
'frmSetup2.Show
'Else
frmsetup1.Show
'End If

End Sub
