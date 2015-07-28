CREATE TABLE [tbl_Tags] (
  [Tag_ID] AUTOINCREMENT,
  [Transect_ID] LONG ,
  [Tag] VARCHAR (10),
  [Feature] VARCHAR (2),
  [Transect] SHORT ,
  [Type] VARCHAR (3),
  [Headpin_Distance] LONG ,
  [Orientation_Pin] VARCHAR (1),
  [Label_Number] LONG ,
  [Tag_Comments] VARCHAR (255),
  [IsActive] LONG 
)
