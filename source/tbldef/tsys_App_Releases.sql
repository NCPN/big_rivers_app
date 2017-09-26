CREATE TABLE [tsys_App_Releases] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [ReleaseDate] DATETIME ,
  [DatabaseTitle] VARCHAR (100),
  [VersionNumber] VARCHAR (20),
  [FileName] VARCHAR (50),
  [ReleaseBy_ID] LONG ,
  [ReleaseBy] VARCHAR (50),
  [ReleaseNotes] LONGTEXT ,
  [IsSupported] BYTE ,
   CONSTRAINT 
)
