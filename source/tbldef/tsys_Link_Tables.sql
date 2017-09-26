CREATE TABLE [tsys_Link_Tables] (
  [LinkType] VARCHAR (50),
  [LinkTable] VARCHAR (100),
  [LinkDb] VARCHAR (100),
  [TableType] VARCHAR (50),
  [DescriptionText] VARCHAR (255),
  [IsHidden] BYTE ,
  [AllowEditsLookup] BYTE ,
  [BrowserView] BYTE ,
  [Is_hidden] BIT ,
  [Allow_edits_lookup] BIT ,
  [Browser_view] BIT ,
  [SortOrder] BYTE ,
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([LinkType], [LinkTable])
)
