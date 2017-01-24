CREATE TABLE [xAction] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [ContactID] LONG ,
  [CategoryID] LONG ,
  [Category] VARCHAR (25),
  [ActionType] VARCHAR (25),
  [ActionDate] DATETIME 
)
