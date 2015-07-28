CREATE TABLE [Woody_Canopy] (
  [Woody_Canopy_ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Event_ID_FK] VARCHAR (50),
  [Plot_ID] VARCHAR (50),
  [Master_PLANT_Code_FK] VARCHAR (20),
  [Woody_Canopy_Pct_Cover] SHORT 
)
