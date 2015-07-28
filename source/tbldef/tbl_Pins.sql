CREATE TABLE [tbl_Pins] (
  [Pin_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Location_ID_FK] VARCHAR (50),
  [Coordinate_ID_FK] SHORT ,
  [Pin_Type] VARCHAR (2),
  [Pin_Name] VARCHAR (255)
)
