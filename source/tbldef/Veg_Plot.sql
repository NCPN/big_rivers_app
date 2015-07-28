CREATE TABLE [Veg_Plot] (
  [Plot_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Plot_Number] SHORT ,
  [Transect_ID_FK] VARCHAR (50),
  [Modal_Sediment_Size] VARCHAR (3),
  [Percent_Fine] SHORT ,
  [Understory_Rooted_Pct_Cover] SHORT ,
  [Woody_Canopy_Pct_Cover] SHORT ,
  [Plot_Density] SHORT ,
  [No_Rooted_Veg] BYTE 
)
