CREATE TABLE [SensitiveLocations] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [ParkCode] VARCHAR (4),
  [Location_ID] LONG ,
  [CreateDate] DATETIME ,
  [CreatedBy_ID] SHORT ,
  [LastModified] DATETIME ,
  [ModifiedBy_ID] SHORT 
)
