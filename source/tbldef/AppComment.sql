CREATE TABLE [AppComment] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [CommentType] VARCHAR (255),
  [TypeID] LONG ,
  [Comment] VARCHAR (255),
  [CreateDate] DATETIME ,
  [CreatedBy] SHORT 
)
