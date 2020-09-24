Attribute VB_Name = "CommonFn"
Dim db As Database
Dim rs As Recordset
Dim trc, trp, trch As Long

Public Sub StatRefresh()
Set db = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\db2.mdb")
Set rs = db.OpenRecordset("Products")
trp = rs.RecordCount
frmMain.txtprod.Text = trp
Set rs = db.OpenRecordset("Contacts")
trc = rs.RecordCount
frmMain.txtcust.Text = trc
Set rs = db.OpenRecordset("Cashrec")
rs.MoveFirst
trch = rs!Cash
frmMain.txtcash.Text = trch
End Sub
