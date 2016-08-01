CREATE TABLE [SensitiveSpecies] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [ParkCode] VARCHAR (4),
  [LUcode] VARCHAR (25),
  [CreateDate] DATETIME ,
  [CreatedBy_ID] SHORT ,
  [LastModified] DATETIME ,
  [ModifiedBy_ID] SHORT 
)
