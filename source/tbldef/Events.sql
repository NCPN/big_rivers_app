CREATE TABLE [Events] (
  [Event_ID] VARCHAR (50),
  [Site_ID_FK] VARCHAR (50),
  [Location_FK] VARCHAR (50),
  [Start_Date] DATETIME ,
  [Location_ID] VARCHAR (1),
  [Protocol_Name] VARCHAR (255),
  [Protocol_Version_Key] VARCHAR (50),
  [Observer] VARCHAR (50),
  [Recorder] VARCHAR (50),
  [Visit_Comments] LONGTEXT 
)
