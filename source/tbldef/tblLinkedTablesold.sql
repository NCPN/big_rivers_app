CREATE TABLE [tblLinkedTablesold] (
  [LinkCategory] VARCHAR (50),
  [LinkTableName] VARCHAR (50),
  [LinkTableVersion] VARCHAR (5),
   CONSTRAINT [PrimaryKey] PRIMARY KEY ([LinkCategory], [LinkTableName])
)
