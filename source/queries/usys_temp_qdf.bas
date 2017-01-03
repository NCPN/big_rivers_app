﻿dbMemo "SQL" ="PARAMETERS pkcode Text ( 4 );\015\012SELECT DISTINCT t.Master_PLANT_Code, t.LU_C"
    "ode, p.Utah_species, t.LU_Code + \" (\" + p.Utah_species + \")\" AS ddSpecies\015"
    "\012FROM tsys_species_presence_by_park AS t INNER JOIN tlu_NCPN_Plants AS p ON p"
    ".Master_PLANT_Code = t.Master_PLANT_Code\015\012WHERE t.LU_Code IS NOT NULL\015\012"
    "AND t.presence <>'NP'\015\012AND t.ParkCode = [pkcode]\015\012ORDER BY t.LU_Code"
    ";\015\012"
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
