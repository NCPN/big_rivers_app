CREATE TABLE [Photo_Points] (
  [Photo_Point_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Location_ID_FK] VARCHAR (50),
  [Coordinate_ID_FK] VARCHAR (50),
  [Photos_Per_Point] LONG 
)
