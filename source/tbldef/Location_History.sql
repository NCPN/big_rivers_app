CREATE TABLE [Location_History] (
  [Location_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Loc_ID] VARCHAR (255),
  [Loc_Type] VARCHAR (25),
  [Loc_Name] VARCHAR (100),
  [GIS_Location_ID] VARCHAR (50),
  [Meta_MID] VARCHAR (50),
  [Head_to_Orient_Distance] SHORT ,
  [Head_to_Orient_Bearing] SHORT ,
  [Updated_Date] VARCHAR (50),
  [Loc_Notes] LONGTEXT ,
  [Last_Update] DATETIME ,
  [Updated_By] VARCHAR (50)
)
