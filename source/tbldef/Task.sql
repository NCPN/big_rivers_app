CREATE TABLE [Task] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [TaskType] VARCHAR (255),
  [TypeID] LONG ,
  [Task] VARCHAR (255),
  [Status] LONG ,
  [Priority] LONG ,
  [RequestedBy] LONG ,
  [RequestDate] DATETIME ,
  [CompletedBy] LONG ,
  [CompleteDate] DATETIME ,
  [LastUpdateBy] LONG ,
  [LastUpdate] DATETIME 
)
