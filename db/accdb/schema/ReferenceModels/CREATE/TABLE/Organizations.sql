CREATE TABLE Organizations
(
	OrganizationId INTEGER NOT NULL UNIQUE CONSTRAINT PK_Organizations PRIMARY KEY AUTOINCREMENT,
	Code TEXT(255) NOT NULL,
	Name TEXT(255) NULL
);
