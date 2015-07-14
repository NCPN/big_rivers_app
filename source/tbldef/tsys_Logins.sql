CREATE TABLE [tsys_Logins] (
  [Time_stamp] DATETIME ,
  [User_name] VARCHAR (50),
  [Action_taken] VARCHAR (50),
   CONSTRAINT [pk_tsys_Logins] PRIMARY KEY ([User_name], [Time_stamp])
)
