Attribute VB_Name = "Module1"

Public Conn As New ADODB.Connection
Public RSBarang As ADODB.Recordset
Public RSSupplier As ADODB.Recordset
Public RSCustomer As ADODB.Recordset
Public RSDetailKeluar As ADODB.Recordset
Public RSDetailMinta As ADODB.Recordset
Public RSPemakai As ADODB.Recordset
Public RSDetailTerima As ADODB.Recordset
Public RSPenerimaan As ADODB.Recordset
Public RSPengeluaran As ADODB.Recordset
Public RSMintaBeli As ADODB.Recordset
Public RSPermintaanUser As ADODB.Recordset
Public PathData As String


Public Sub Koneksi()
Set Conn = New ADODB.Connection
Set RSBarang = New ADODB.Recordset
Set RSSupplier = New ADODB.Recordset
Set RSCustomer = New ADODB.Recordset
Set RSDetailKeluar = New ADODB.Recordset
Set RSDetailMinta = New ADODB.Recordset
Set RSPemakai = New ADODB.Recordset
Set RSDetailTerima = New ADODB.Recordset
Set RSPenerimaan = New ADODB.Recordset
Set RSPengeluaran = New ADODB.Recordset
Set RSMintaBeli = New ADODB.Recordset
Set RSPermintaanUser = New ADODB.Recordset
PathData = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBPersediaan.mdb"
Conn.Open PathData
End Sub





