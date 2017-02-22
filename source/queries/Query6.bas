dbMemo "SQL" ="PARAMETERS pcode Text ( 4 ), scode Text ( 2 );\015\012SELECT l.ID, CollectionSou"
    "rceName, LocationType, LocationName, HeadToOrientDistance_m, HeadToOrientBearing"
    ", LocationNotes, SWITCH (\015\012LocationType ='P', \015\012(\015\012SELECT DIST"
    "INCT vp.ID \015\012FROM ((VegPlot vp\015\012INNER JOIN Site s ON s.ID = vp.Site_"
    "ID)\015\012INNER JOIN Park p ON p.ID = s.Park_ID)\015\012WHERE \015\012CStr(Plot"
    "Number) = CStr(CollectionSourceName)\015\012AND p.ParkCode = [pcode]\015\012AND "
    "s.SiteCode = [scode]\015\012),\015\012LocationType ='T',  \015\012(\015\012SELEC"
    "T DISTINCT vt.ID \015\012FROM (((VegTransect vt\015\012INNER JOIN Site_VegTranse"
    "ct svt ON svt.VegTransect_ID = vt.ID)\015\012INNER JOIN Site s ON s.ID = svt.Sit"
    "e_ID)\015\012INNER JOIN Park p ON p.ID = s.Park_ID)\015\012WHERE \015\012CStr(Tr"
    "ansectNumber) = CStr(CollectionSourceName)\015\012AND p.ParkCode = [pcode]\015\012"
    "AND s.SiteCode = [scode]\015\012),\015\012LocationType ='F',  \015\012(\015\012P"
    "ARAMETERS pcode TEXT(4);\015\012SELECT DISTINCT f.ID \015\012FROM (((Feature f\015"
    "\012INNER JOIN Site_Feature sf ON sf.Feature_ID = f.ID)\015\012INNER JOIN Site s"
    " ON s.ID = sf.Site_ID)\015\012INNER JOIN Park p ON p.ID = s.Park_ID)\015\012WHER"
    "E \015\012CStr(Feature) = CStr(CollectionSourceName)\015\012AND p.ParkCode = [pc"
    "ode]\015\012AND s.SiteCode = [scode];\015\012)\015\012) AS LocTypeID, (SELECT CO"
    "UNT(sl.Location_ID) FROM SensitiveLocations sl WHERE sl.Location_ID = l.ID\015\012"
    ") AS IsSensitive\015\012FROM Location AS l;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x1bdc6617c78ea34f8e4a186e6e989936
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="LocationType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="HeadToOrientDistance_m"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="HeadToOrientBearing"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocationNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CollectionSourceName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocationName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocTypeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vt.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vp.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.Version"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.IsSupported"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.Context"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.Syntax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.TemplateName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.Params"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.Template"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.Remarks"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.EffectiveDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.RetireDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.CreateDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.CreatedBy_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.LastModified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Db_Templates.LastModifiedBy_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vt.TransectNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.SiteCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[csn]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[ltype]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[lname]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[dist]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[brg]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[lnotes]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1006"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[CID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1008"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[LMID]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1007"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IsSensitive"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.ParkCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="f.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="f.Feature"
        dbLong "AggregateType" ="-1"
    End
End
