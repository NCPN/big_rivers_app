CREATE TABLE [ImportedData] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [ImportDate] DATETIME ,
  [SourceFile] VARCHAR (50),
  [DestinationTable] VARCHAR (25),
  [NumberOfRecordsImported] SHORT ,
  [StartRecord_ID] SHORT ,
  [EndRecord_ID] SHORT 
)
