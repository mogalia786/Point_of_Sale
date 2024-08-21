Attribute VB_Name = "Transactionlog"

Public Sub RecordAction(Uname As String, DateofAction As Date, TimeofAction As String, ResultofAction As String, Action As String)
Dim rAction As New Recordset
With rAction
.Open "select * from transactionlogsheet", c, adOpenDynamic, adLockOptimistic
.AddNew
!username = Uname
!dateofevent = DateofAction
!timeofevent = TimeofAction
!event = Action
!Result = ResultofAction
.Update
.Close
End With


End Sub
