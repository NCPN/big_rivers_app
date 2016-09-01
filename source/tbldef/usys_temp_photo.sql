CREATE TABLE [usys_temp_photo] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Event_ID] LONG ,
  [Photographer_ID] LONG ,
  [PhotoPath] VARCHAR (255),
  [PhotoFilename] VARCHAR (255),
  [PhotoDate] DATETIME ,
  [PhotoType] VARCHAR (1)
)
