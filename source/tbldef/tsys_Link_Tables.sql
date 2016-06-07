CREATE TABLE [tsys_Link_Tables] (
  [LinkType] VARCHAR (50),
  [LinkTable] VARCHAR (100),
  [LinkDb] VARCHAR (100),
  [TableType] VARCHAR (50),
  [DescriptionText] VARCHAR (255),
  [IsHidden] BYTE ,
  [AllowEditsLookup] BYTE ,
  [BrowserView] BYTE ,
  [Is_hidden] VARCHAR ,
  [Allow_edits_lookup] VARCHAR ,
  [Browser_view] VARCHAR ,
  [SortOrder] BYTE ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([LinkType], [LinkTable])
)
