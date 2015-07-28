CREATE TABLE [Overlays] (
  [Overlay_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Overlain_Item_ID] VARCHAR (50),
  [Overlay] VARCHAR (50),
  [Overlay_X] DOUBLE ,
  [Overlay_Y] DOUBLE ,
  [Create_Date] DATETIME ,
  [Created_By] VARCHAR (50)
)
