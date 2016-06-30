dbMemo "SQL" ="SELECT tsys_App_Releases.ID, 'Version ' & [VersionNumber] & ' (' & [ReleaseDate]"
    " & ')' AS Version\015\012FROM tsys_App_Releases;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xa4b932da067afc479bc040b6992e7cba
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="AppUser"
        dbInteger "ColumnWidth" ="3195"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xb4c579182da7de4189d79f53a69d9ef9
        End
    End
    Begin
        dbText "Name" ="c.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AccessLevel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="a.AccessLevel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_App_Releases.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Version"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xa897549d74de864d83714590c72ca5c7
        End
    End
End
