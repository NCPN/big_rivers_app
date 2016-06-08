CREATE TABLE [xCopy Of tsys_Db_Templates] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Version] DOUBLE ,
  [IsSupported] SHORT ,
  [Context] VARCHAR (255),
  [Syntax] VARCHAR (10),
  [TemplateName] VARCHAR (255),
  [Params] VARCHAR (255),
  [Template] LONGTEXT ,
  [Remarks] VARCHAR (255),
  [EffectiveDate] DATETIME ,
  [RetireDate] DATETIME ,
  [CreateDate] DATETIME ,
  [CreatedBy_ID] SHORT ,
  [LastModified] DATETIME ,
  [LastModifiedBy_ID] SHORT 
)
