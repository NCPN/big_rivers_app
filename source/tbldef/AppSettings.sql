CREATE TABLE [AppSettings] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [DisplayName] VARCHAR (25),
  [FormName] VARCHAR (25),
  [FormatIcon] VARCHAR (50),
  [OArgs] VARCHAR (10),
  [Sequence] SHORT 
)
