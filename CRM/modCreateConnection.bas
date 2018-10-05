Attribute VB_Name = "modCreateConnection"
Option Explicit

Public conn As ADODB.Connection
Public rs As ADODB.Recordset

'str = "Microsoft.ACE.OLEDB.12.0;Data Source = " & App.Path & "\data\data.accdb;"


Public Sub createConnection()
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    conn.CursorLocation = adUseClient
    conn.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source = " & App.Path & "\data\data.accdb; Jet OLEDB"
    conn.Open


End Sub

