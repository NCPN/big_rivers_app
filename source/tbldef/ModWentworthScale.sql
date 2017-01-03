CREATE TABLE [ModWentworthScale] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Code] VARCHAR (3),
  [Label] VARCHAR (25),
  [DiameterRange_mm] VARCHAR (255),
  [CategoryOrder] SHORT ,
  [KeyOrder] SHORT ,
  [ActiveYear] SHORT ,
  [RetireYear] SHORT 
)
