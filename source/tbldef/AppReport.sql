CREATE TABLE [AppReport] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [DisplayName] VARCHAR (25),
  [ReportName] VARCHAR (25),
  [ReportTemplate] VARCHAR (25),
  [FormatIcon] VARCHAR (50),
  [Sequence] SHORT 
)
