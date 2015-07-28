CREATE TABLE [Understory] (
  [Understory_ID] VARCHAR (50) CONSTRAINT [Understory_ID] UNIQUE ,
  [Event_ID_FK] VARCHAR (50),
  [Plot_ID] VARCHAR (50),
  [Master_PLANT_Code_FK] VARCHAR (20),
  [IsSeedling] BYTE 
)
