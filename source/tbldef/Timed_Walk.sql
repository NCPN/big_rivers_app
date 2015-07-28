CREATE TABLE [Timed_Walk] (
  [Timed_Walk_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Event_ID_FK] VARCHAR (50),
  [Collection_Place_FK] VARCHAR (50),
  [Collection_Type] VARCHAR (50),
  [Walk_Start_Date] DATETIME ,
  [Start_Time] DATETIME ,
  [Walk_End_Date] DATETIME ,
  [End_Time] DATETIME ,
  [Master_PLANT_Code_FK] VARCHAR (20)
)
