CREATE VIEW "Balances" AS
SELECT Allocations.BFY AS BFY, 
    Allocations.RPIO AS RPIO,
    Allocations.AhCode AS AhCode,
	Allocations.FundName As FundName,
    Allocations.FundCode AS FundCode,
	Allocations.ProgramProjectName AS ProgramProjectName,
    Allocations.ProgramProjectCode AS ProgramProjectCode,
    Allocations.OrgCode AS OrgCode,
	Allocations.BocName AS BocName,
    Allocations.BocCode AS BocCode,
	Allocations.DivisionName As DivisionName,
    Allocations.RcCode AS RcCode,
    Allocations.Amount AS Authority,
	CASE
	WHEN
    (SELECT ROUND(Outlays.OpenCommitments + Outlays.Obligations)
		FROM Outlays
        WHERE Outlays.BFY = Allocations.BFY AND
            Outlays.AhCode = Allocations.AhCode AND
            Outlays.FundCode = Allocations.FundCode AND
            Outlays.ProgramProjectCode = Allocations.ProgramProjectCode AND
            Outlays.BocCode = Allocations.BocCode AND
            Outlays.RcCode = Allocations.RcCode ) IS NULL THEN 0.0
	WHEN NOT 
		(SELECT ROUND(Outlays.OpenCommitments + Outlays.Obligations)
			FROM Outlays
			WHERE Outlays.BFY = Allocations.BFY AND
				Outlays.AhCode = Allocations.AhCode AND
				Outlays.FundCode = Allocations.FundCode AND
				Outlays.ProgramProjectCode = Allocations.ProgramProjectCode AND
				Outlays.BocCode = Allocations.BocCode AND
				Outlays.RcCode = Allocations.RcCode ) IS NULL THEN 
			(SELECT ROUND(Outlays.OpenCommitments + Outlays.Obligations)
				FROM Outlays
				WHERE Outlays.BFY = Allocations.BFY AND
					Outlays.AhCode = Allocations.AhCode AND
					Outlays.FundCode = Allocations.FundCode AND
					Outlays.ProgramProjectCode = Allocations.ProgramProjectCode AND
					Outlays.BocCode = Allocations.BocCode AND
					Outlays.RcCode = Allocations.RcCode ) 
	END Used,
	CASE
		WHEN (SELECT ROUND(Outlays.OpenCommitments + Outlays.Obligations)
			FROM Outlays
			WHERE Outlays.BFY = Allocations.BFY AND
				Outlays.AhCode = Allocations.AhCode AND
				Outlays.FundCode = Allocations.FundCode AND
				Outlays.ProgramProjectCode = Allocations.ProgramProjectCode AND
				Outlays.BocCode = Allocations.BocCode AND
				Outlays.RcCode = Allocations.RcCode ) IS NULL THEN Allocations.Amount
		WHEN NOT (SELECT ROUND(Outlays.OpenCommitments + Outlays.Obligations)
			FROM Outlays
			WHERE Outlays.BFY = Allocations.BFY AND
				Outlays.AhCode = Allocations.AhCode AND
				Outlays.FundCode = Allocations.FundCode AND
				Outlays.ProgramProjectCode = Allocations.ProgramProjectCode AND
				Outlays.BocCode = Allocations.BocCode AND
				Outlays.RcCode = Allocations.RcCode ) IS NULL THEN
			(SELECT ROUND(Allocations.Amount - (Outlays.OpenCommitments + ABS(Outlays.Obligations)))
			FROM Outlays
			WHERE Outlays.BFY = Allocations.BFY AND
				Outlays.AhCode = Allocations.AhCode AND
				Outlays.FundCode = Allocations.FundCode AND
				Outlays.ProgramProjectCode = Allocations.ProgramProjectCode AND
				Outlays.BocCode = Allocations.BocCode AND
				Outlays.RcCode = Allocations.RcCode )  
	END Available
FROM Allocations
WHERE 
	Allocations.BudgetLevel = '8' AND
	BocCode IN ('21', '28', '36', '37', '38', '41') AND
	Allocations.Amount > 0
ORDER BY BFY DESC