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
Public ReaderPort As Integer



Public Sub Main()
frmsetup1.Show
End Sub

Public Sub OpenTill()
Printer.Font.Name = "Control"
Printer.Font.Size = "10"
Printer.Print "A"
Printer.EndDoc
End Sub
