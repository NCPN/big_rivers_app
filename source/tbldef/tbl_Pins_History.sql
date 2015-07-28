CREATE TABLE [tbl_Pins_History] (
  [Pin_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Location_ID_FK] VARCHAR (50),
  [Coordinate_ID_FK] VARCHAR (50),
  [Pin_Type] VARCHAR (2),
  [Last_Update] DATETIME ,
  [Updated_By] VARCHAR (50)
)
