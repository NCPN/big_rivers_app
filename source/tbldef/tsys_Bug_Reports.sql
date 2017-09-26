CREATE TABLE [tsys_Bug_Reports] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Release_ID] VARCHAR (50),
  [ReportDate] DATETIME ,
  [FoundBy_ID] LONG ,
  [ReportedBy_ID] LONG ,
  [ReportDetails] LONGTEXT ,
  [FixDate] DATETIME ,
  [FixedBy_ID] LONG ,
  [FixDetails] LONGTEXT 
)
