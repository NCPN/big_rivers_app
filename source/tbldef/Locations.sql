CREATE TABLE [Locations] (
  [Location_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Loc_ID] VARCHAR (255),
  [Loc_Type] VARCHAR (25),
  [Loc_Name] VARCHAR (100),
  [Head_to_Orient_Distance] SHORT ,
  [Head_to_Orient_Bearing] SHORT ,
  [Updated_Date] VARCHAR (50),
  [Loc_Notes] LONGTEXT 
)
