dbMemo "SQL" ="SELECT DISTINCT tsys_Link_Dbs.Is_ODBC, IIf([Type]=4,ParseConnectionStr([Connect]"
    "),ParseFileName([Database])) AS CurrDb, tsys_Link_Tables.Link_db, tsys_Link_Dbs."
    "Server, MSysObjects.Database AS CurrPath, tsys_Link_Dbs.File_path\015\012FROM ts"
    "ys_Link_Dbs INNER JOIN (MSysObjects INNER JOIN tsys_Link_Tables ON MSysObjects.N"
    "ame = tsys_Link_Tables.Link_table) ON tsys_Link_Dbs.Link_db = tsys_Link_Tables.L"
    "ink_db\015\012WHERE (((MSysObjects.Type) In (4,6)) AND ((IIf([Type]=4,ParseConne"
    "ctionStr([Connect]),ParseFileName([Database])))<>[tsys_Link_Tables].[Link_db])) "
    "OR (((MSysObjects.Type) In (4,6)) AND ((IIf([Type]=4,ParseConnectionStr([Connect"
    "],'SERVER=')))<>[Server])) OR (((MSysObjects.Type) In (4,6)) AND ((MSysObjects.D"
    "atabase)<>[File_path])) OR (((MSysObjects.Type)=4) AND ((tsys_Link_Dbs.Is_ODBC)="
    "False)) OR (((MSysObjects.Type)=6) AND ((tsys_Link_Dbs.Is_ODBC)=True)) OR (((IIf"
    "([Type]=4,ParseConnectionStr([Connect],'SERVER='))) Is Null) AND ((tsys_Link_Dbs"
    ".Server) Is Not Null)) OR (((IIf([Type]=4,ParseConnectionStr([Connect],'SERVER='"
    "))) Is Not Null) AND ((tsys_Link_Dbs.Server) Is Null)) OR (((MSysObjects.Databas"
    "e) Is Null) AND ((tsys_Link_Dbs.File_path) Is Not Null)) OR (((MSysObjects.Datab"
    "ase) Is Not Null) AND ((tsys_Link_Dbs.File_path) Is Null));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xe1f76eb6cb92684fb73761ab9b1a4c67
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tsys_Link_Dbs.Is_ODBC"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x933515ffbfb92e498658c2d175968dad
        End
    End
    Begin
        dbText "Name" ="CurrDb"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x0a751d867b87964f80ff51de9783d6db
        End
    End
    Begin
        dbText "Name" ="tsys_Link_Tables.Link_db"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x67ca000795558c419bc26666bcb38ca2
        End
    End
    Begin
        dbText "Name" ="tsys_Link_Dbs.Server"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x093fdb5dc079cc49b5949efef5b73b13
        End
    End
    Begin
        dbText "Name" ="CurrPath"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0x3f33a8983034f346838ec9d9e2277031
        End
    End
    Begin
        dbText "Name" ="tsys_Link_Dbs.File_path"
        dbLong "AggregateType" ="-1"
        dbBinary "GUID" = Begin
            0xd2fa7d6a14975545bd8d8d66451bb65a
        End
    End
End
