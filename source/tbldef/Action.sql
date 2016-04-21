CREATE TABLE [Action] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [ContactID] SHORT ,
  [CategoryID] SHORT ,
  [Category] VARCHAR (25),
  [ActionType] VARCHAR (25),
  [ActionDate] DATETIME 
)
