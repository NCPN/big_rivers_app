CREATE TABLE [Task_List] (
  [Location_ID] VARCHAR (50),
  [Request_date] DATETIME ,
  [Task_desc] VARCHAR (100),
  [Requested_by] VARCHAR (50),
  [Task_status] VARCHAR (50),
  [Date_completed] DATETIME ,
  [Followup_by] VARCHAR (50),
  [Task_notes] LONGTEXT ,
  [Followup_notes] LONGTEXT ,
   CONSTRAINT [pk_tbl_Task_List] PRIMARY KEY ([Location_ID], [Request_date], [Task_desc])
)
