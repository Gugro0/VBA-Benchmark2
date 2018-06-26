# VBA-Benchmark

This little Access Add-In allows you to Benchmark your VBA Code.

You ever wanted to know if DLookup or opening a Recordset is faster, just try it !!

Create a Class e.g. named "clsLookup_Bench":

```
Public Sub Bench_DLookup()
 Debug.Assert DLookup("Feld1", "tblTest", "ID=1") = "Peter Müller"
End Sub

Public Sub Bench_DAO_Recordset()
 With CurrentDb.OpenRecordset("SELECT TOP 1 Feld1 FROM tblTest WHERE ID=1")
   Debug.Assert .Fields("Feld1") = "Peter Müller"
 End With
End Sub

Public Sub Bench_ADO_Recordset()
 With CurrentProject.Connection.Execute("SELECT TOP 1 Feld1 FROM tblTest WHERE ID=1")
   Debug.Assert .Fields("Feld1") = "Peter Müller"
 End With
End Sub
```

and start the VBA Benchmark AddIn:
