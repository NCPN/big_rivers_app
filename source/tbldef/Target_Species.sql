CREATE TABLE [Target_Species] (
  [Master_Plant_Code_FK] VARCHAR (20),
  [List_Type] VARCHAR (2),
  [Target_Year] SHORT ,
  [Species_Name] VARCHAR (255),
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([Master_Plant_Code_FK], [List_Type], [Target_Year])
)
