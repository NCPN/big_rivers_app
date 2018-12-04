CREATE TABLE [DbCleanerTables] (
  [TableName] VARCHAR (255) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [IsExcluded] BYTE 
)
