CREATE TABLE [tsys_Link_Dbs] (
  [Link_db] VARCHAR (100) CONSTRAINT [pk_tsys_Link_Files] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Db_desc] VARCHAR (50),
  [Backups] VARCHAR ,
  [Is_ODBC] VARCHAR ,
  [Is_Network_db] BYTE ,
  [File_path] VARCHAR (255),
  [Server] VARCHAR (100),
  [New_db] VARCHAR (100),
  [New_path] VARCHAR (255),
  [New_server] VARCHAR (100),
  [Sort_order] BYTE 
)
