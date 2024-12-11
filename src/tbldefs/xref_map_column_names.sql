CREATE TABLE [xref_map_column_names] (
  [source_col_name] VARCHAR (255) CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [target_col_name] VARCHAR (255),
  [parameter_name] VARCHAR (255),
  [parameter_unit] VARCHAR (255)
)
