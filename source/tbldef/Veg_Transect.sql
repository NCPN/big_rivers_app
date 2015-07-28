CREATE TABLE [Veg_Transect] (
  [Transect_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Loc_ID_FK] VARCHAR (50),
  [Event_ID_FK] VARCHAR (50),
  [Transect_Number] SHORT ,
  [Transect_Type] VARCHAR (1),
  [Sample_Date] DATETIME ,
  [Observer] VARCHAR (50),
  [Recorder] VARCHAR (50),
  [AF_Size] BYTE ,
  [AF_Type] VARCHAR (255)
)
