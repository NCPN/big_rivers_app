CREATE TABLE [tbl_Taglines] (
  [Line_Distance_ID] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Line_Distance_Source] VARCHAR (255),
  [Line_Distance_Source_ID_FK] VARCHAR (50),
  [Line_Distance_Type] VARCHAR (2),
  [Line_Distance] VARCHAR ,
  [Height_Type] VARCHAR (2),
  [Height] VARCHAR 
)
