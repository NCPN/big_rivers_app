CREATE TABLE [tsys_Link_Files] (
  [Link_type] VARCHAR (50) CONSTRAINT [PrimaryKey] PRIMARY KEY  UNIQUE  NOT NULL ,
  [Link_file_name] VARCHAR (100),
  [Link_file_path] VARCHAR (255),
  [Link_description] VARCHAR (255),
  [New_file_name] VARCHAR (100),
  [New_file_path] VARCHAR (255)
)
