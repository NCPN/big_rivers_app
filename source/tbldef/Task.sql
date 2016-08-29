CREATE TABLE [Task] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [TaskType] VARCHAR (255),
  [TaskType_ID] LONG ,
  [Task] VARCHAR (255),
  [Status_ID] LONG ,
  [Priority_ID] LONG ,
  [RequestedBy_ID] LONG ,
  [RequestDate] DATETIME ,
  [CompletedBy_ID] LONG ,
  [CompleteDate] DATETIME ,
  [LastModifiedBy_ID] LONG ,
  [LastModified] DATETIME 
)
