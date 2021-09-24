CREATE TABLE "Transfers" (
	"TransferId"	INTEGER UNIQUE,
	"BudgetLevel"	TEXT DEFAULT 'NS',
	"DocPrefix"	TEXT DEFAULT 'NS',
	"DocType"	TEXT DEFAULT 'NS',
	"BFY"	TEXT DEFAULT 'NS',
	"RpioCode"	TEXT DEFAULT 'NS',
	"RpioName"	TEXT DEFAULT 'NOT SPECIFIED',
	"FundCode"	TEXT DEFAULT 'NS',
	"FundName"	TEXT DEFAULT 'NOT SPECIFIED',
	"ReprogrammingNumber"	TEXT DEFAULT 'NS',
	"ControlNumber"	TEXT DEFAULT 'NS',
	"ProcessedDate"	TEXT DEFAULT 'NS',
	"Quarter"	TEXT DEFAULT 'NS',
	"Subline"	TEXT DEFAULT 'NS',
	"AhCode"	TEXT DEFAULT 'NS',
	"AhName"	TEXT DEFAULT 'NOT SPECIFIED',
	"OrgCode"	TEXT DEFAULT 'NS',
	"OrgName"	TEXT DEFAULT 'NOT SPECIFIED',
	"RcCode"	TEXT DEFAULT 'NS',
	"RcName"	TEXT DEFAULT 'NOT SPECIFIED',
	"AccountCode"	TEXT DEFAULT 'NS',
	"ProgramAreaCode"	TEXT DEFAULT 'NS',
	"ProgramAreaName"	TEXT DEFAULT 'NOT SPECIFIED',
	"ProgramProjectName"	TEXT DEFAULT 'NOT SPECIFIED',
	"ProgramProjectCode"	TEXT DEFAULT 'NS',
	"FromTo"	TEXT DEFAULT 'NS',
	"BocCode"	TEXT DEFAULT 'NS',
	"BocName"	TEXT DEFAULT 'NOT SPECIFIED',
	"NpmCode"	TEXT DEFAULT 'NS',
	"Amount"	REAL DEFAULT 0,
	"Purpose"	TEXT DEFAULT 'NS',
	"ExtendedPurpose"	TEXT DEFAULT 'NS',
	"ResourceType"	TEXT DEFAULT 'NS',
	PRIMARY KEY("TransferId" AUTOINCREMENT)
);