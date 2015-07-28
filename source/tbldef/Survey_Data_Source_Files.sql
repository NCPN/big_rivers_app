CREATE TABLE [Survey_Data_Source_Files] (
  [Source_ID] VARCHAR (50),
  [Source_File_Name] VARCHAR (255),
  [Source_File_Path] VARCHAR (255),
  [Survey_Type] VARCHAR (1),
  [Survey_Source] VARCHAR (4),
  [Translation_Point_FK] LONG ,
  [Rotation_Point_FK] LONG ,
  [Translation_Error_FK] VARCHAR (50),
  [Rotation_Error_FK] VARCHAR (50),
  [Base_Error_FK] VARCHAR (50),
  [Survey_Error_FK] VARCHAR (50),
  [Survey_Comments] VARCHAR (255)
)
