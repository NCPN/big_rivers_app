dbMemo "SQL" ="PARAMETERS uname Text ( 50 ), activity Text ( 50 ), accesslvl Text ( 25 );\015\012"
    "INSERT INTO tsys_Logins ( UserName, ActionTaken, ReleaseNumber, AccessLevel )\015"
    "\012VALUES ([uname], [activity], [version], [accesslvl]);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End
