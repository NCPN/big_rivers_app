CREATE TABLE [tsys_BE_Updates] (
  [Update_ID] VARCHAR (50) CONSTRAINT [pk_tsys_Db_Updates] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Is_done] VARCHAR ,
  [Run_date] DATETIME ,
  [SQL_statement] LONGTEXT ,
  [Update_desc] VARCHAR (100)
)
