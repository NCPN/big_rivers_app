CREATE TABLE [tsys_Link_Dbs] (
  [LinkType] VARCHAR (255),
  [LinkDb] VARCHAR (100),
  [DbDesc] VARCHAR (50),
  [Backups] BYTE ,
  [IsODBC] BYTE ,
  [IsNetworkDb] BYTE ,
  [FilePath] VARCHAR (255),
  [Server] VARCHAR (100),
  [NewDb] VARCHAR (100),
  [NewPath] VARCHAR (255),
  [NewServer] VARCHAR (100),
  [SortOrder] SHORT ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([LinkType], [LinkDb])
)
