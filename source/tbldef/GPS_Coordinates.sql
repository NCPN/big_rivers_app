CREATE TABLE [GPS_Coordinates] (
  [GIS_Coordinate_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [GIS_Location_ID] VARCHAR (50),
  [Coordinate_System] VARCHAR (50),
  [UTM_Zone] SHORT ,
  [Datum] VARCHAR (5),
  [GPS_Filename] VARCHAR (50),
  [GPS_Unit] VARCHAR (50),
  [Max_PDOP] DOUBLE ,
  [Max_HDOP] DOUBLE ,
  [Coordinate_Type] VARCHAR (4),
  [X_Coord] DOUBLE ,
  [Y_Coord] DOUBLE ,
  [Z_Coord] DOUBLE ,
  [Coord_Type] VARCHAR (1)
)
