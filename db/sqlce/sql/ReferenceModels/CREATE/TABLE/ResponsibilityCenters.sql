CREATE TABLE ResponsibilityCenters
(
	ResponsibilityCenterId INTEGER NOT NULL UNIQUE IDENTITY(1,1),
	Code NVARCHAR(255) NULL,
	Name NVARCHAR(255) NULL,
	Title NVARCHAR(255) NULL
);

