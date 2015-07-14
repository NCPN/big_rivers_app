CREATE TABLE [tsys_Bug_Reports] (
  [Bug_ID] VARCHAR (50),
  [Release_ID] VARCHAR (50),
  [Report_date] DATETIME ,
  [Found_by] VARCHAR (50),
  [Reported_by] VARCHAR (50),
  [Report_details] LONGTEXT ,
  [Fix_date] DATETIME ,
  [Fixed_by] VARCHAR (255),
  [Fix_details] LONGTEXT 
)
