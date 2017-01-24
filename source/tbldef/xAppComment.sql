CREATE TABLE [xAppComment] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [CommentType] VARCHAR (255),
  [CommentType_ID] LONG ,
  [Comment] VARCHAR (255),
  [CreateDate] DATETIME ,
  [CreatedBy_ID] LONG ,
  [LastModified] DATETIME ,
  [LastModifiedBy_ID] LONG 
)
