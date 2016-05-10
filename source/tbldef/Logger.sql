CREATE TABLE [Logger] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Site_ID] LONG ,
  [SensorType] VARCHAR (5),
  [SensorNumber] VARCHAR (255),
  [Sequence] SHORT 
)
