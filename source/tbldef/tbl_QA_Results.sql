CREATE TABLE [tbl_QA_Results] (
  [Query_name] VARCHAR (100),
  [Data_scope] BYTE ,
  [Time_frame] VARCHAR (30),
  [Query_type] VARCHAR (20),
  [Query_result] VARCHAR (50),
  [Query_run_time] DATETIME ,
  [Query_description] LONGTEXT ,
  [Query_expression] LONGTEXT ,
  [Remedy_desc] LONGTEXT ,
  [Remedy_date] DATETIME ,
  [QA_user] VARCHAR (50),
  [Is_done] VARCHAR ,
   CONSTRAINT [pk_tbl_QA_Results] PRIMARY KEY ([Query_name], [Time_frame], [Data_scope])
)
