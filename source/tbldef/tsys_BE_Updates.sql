CREATE TABLE [tsys_BE_Updates] (
  [ID] VARCHAR (50) CONSTRAINT [pk_tsys_Db_Updates] PRIMARY KEY  UNIQUE  NOT NULL ,
  [IsDone] BIT ,
  [RunDate] DATETIME ,
  [SQLStatement] LONGTEXT ,
  [UpdateDesc] VARCHAR (100)
)
