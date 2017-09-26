CREATE TABLE [tsys_App_Defaults] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Project] VARCHAR (50),
  [Release_ID] LONG ,
  [DataTimeframe] VARCHAR (30),
  [UserName] VARCHAR (50),
  [Park] VARCHAR (50),
  [BackupPromptOnStartup] BYTE ,
  [BackupPromptOnExit] BYTE ,
  [CompactBEOnExit] BYTE ,
  [VerifyLinksOnStartup] BYTE ,
  [WebURL] VARCHAR (200),
  [AppContactName] VARCHAR (50),
  [AppContactOrg] VARCHAR (50),
  [AppContactPhone] VARCHAR (50),
  [AppContactEmail] VARCHAR (50)
)
