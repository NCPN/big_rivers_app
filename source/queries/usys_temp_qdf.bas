dbMemo "SQL" ="PARAMETERS RefTable Text ( 25 ), RefID Long, ID Long, Activity Text ( 2 ), Actio"
    "nDate DateTime;\015\012INSERT INTO RecordAction ( ReferenceType, Reference_ID, C"
    "ontact_ID, Activity, ActionDate )\015\012VALUES ([RefTable], [RefID], [ID], [Act"
    "ivity], [ActionDate]);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0xd0adfdbe4d7a184da5399778863d2cc8
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End
