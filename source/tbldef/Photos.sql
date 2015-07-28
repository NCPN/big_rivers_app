CREATE TABLE [Photos] (
  [Photo_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Photo_Date] DATETIME ,
  [Photo_Type] VARCHAR (1),
  [Photographer] VARCHAR (50),
  [Digital_Filename] VARCHAR (50),
  [Comments] LONGTEXT ,
  [NCPN_Image_ID] VARCHAR (50),
  [Photog_Facing] VARCHAR (4),
  [Photog_Point_Location] VARCHAR (10),
  [Photog_Location_Desc] VARCHAR (255),
  [Photog_Orientation] VARCHAR (255),
  [Photo_Coord_FK] VARCHAR (50),
  [Subject_Pt_Location] VARCHAR (10),
  [Subject] VARCHAR (3),
  [IsCloseup] BYTE ,
  [InActive] BYTE ,
  [Last_Update] DATETIME 
)
