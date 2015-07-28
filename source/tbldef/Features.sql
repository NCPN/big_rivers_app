CREATE TABLE [Features] (
  [Feature_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Loc_ID_FK] LONG ,
  [Feature] VARCHAR (1),
  [Feature_description] VARCHAR (255),
  [Feature_directions] LONGTEXT 
)
