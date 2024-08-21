Attribute VB_Name = "ConnectionModule"
Public c As New Connection
Public HC As New Connection
Public servername As String
Public sharename As String
Public SqlServerName As String
Public SQLUserid As String
Public SQLUserpwd As String
Public CurrentUser As String


Public Sub ConnectMe(SqlServerName As String, SqlUID As String, SQLpwd As String)
With c
'.ConnectionString = "\\" & server & "\" & share & "\zk.ngd"
'.Provider = "microsoft.jet.oledb.4.0"
'.Open
.ConnectionString = "server=" & SqlServerName & ";driver={sql server};uid=" & SqlUID & ";pwd=" & SQLpwd
.Open
.DefaultDatabase = "POS"

End With
End Sub

Public Sub ConnectMe2(SqlServerName As String, SqlUID As String, SQLpwd As String)
With HC
'.ConnectionString = "\\" & server & "\" & share & "\zk.ngd"
'.Provider = "microsoft.jet.oledb.4.0"
'.Open
.ConnectionString = "server=" & SqlServerName & ";driver={sql server};uid=" & SqlUID & ";pwd=" & SQLpwd
.Open
.DefaultDatabase = "HeadCount"

End With
End Sub

