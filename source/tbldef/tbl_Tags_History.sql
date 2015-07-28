CREATE TABLE [tbl_Tags_History] (
  [Tag_ID] VARCHAR (10) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Feature] VARCHAR (2),
  [Transect] SHORT ,
  [Type] VARCHAR (3),
  [Headpin_Distance] LONG ,
  [Orientation_Pin] VARCHAR (1),
  [Label_Number] LONG ,
  [Tag_Comments] VARCHAR (255),
  [IsReplaced] BYTE ,
  [Last_Update] DATETIME ,
  [Updated_By] VARCHAR (50)
)
