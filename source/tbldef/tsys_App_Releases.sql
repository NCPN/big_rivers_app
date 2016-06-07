CREATE TABLE [tsys_App_Releases] (
  [ID] AUTOINCREMENT,
  [ReleaseDate] DATETIME ,
  [DatabaseTitle] VARCHAR (100),
  [VersionNumber] VARCHAR (20),
  [FileName] VARCHAR (50),
  [ReleaseBy_ID] SHORT ,
  [ReleaseBy] VARCHAR (50),
  [ReleaseNotes] LONGTEXT ,
  [IsSupported] BYTE ,
   CONSTRAINT 
)
