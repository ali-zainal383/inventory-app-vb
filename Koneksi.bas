Attribute VB_Name = "Module1"
Public CONN As ADODB.Connection
Public RSBarang As ADODB.Recordset
Public RSKaryawan As ADODB.Recordset
Public RSMaintenance As ADODB.Recordset
Public RSAdmin As ADODB.Recordset
Public LokasiData As String

Public Sub Koneksi()
Set CONN = New ADODB.Connection
Set RSBarang = New ADODB.Recordset
Set RSKaryawan = New ADODB.Recordset
Set RSMaintenance = New ADODB.Recordset
Set RSAdmin = New ADODB.Recordset

LokasiData = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\DBJne.mdb"
CONN.Open LokasiData
End Sub

