CREATE TABLE [tsys_Link_Tables] (
  [Link_type] VARCHAR (50),
  [Link_table] VARCHAR (100),
  [Link_db] VARCHAR (100),
  [Table_type] VARCHAR (50),
  [Description_text] VARCHAR (255),
  [Is_hidden] VARCHAR ,
  [Allow_edits_lookup] VARCHAR ,
  [Browser_view] VARCHAR ,
  [Sort_order] BYTE ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([Link_type], [Link_table])
)
