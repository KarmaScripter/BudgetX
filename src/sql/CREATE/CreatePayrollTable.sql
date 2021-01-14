SELECT DISTINCTROW DivisionAuthority.PrcId, DivisionAuthority.RPIO, DivisionAuthority.BFY, DivisionAuthority.AhCode,
    DivisionAuthority.FundCode, DivisionAuthority.FundName, DivisionAuthority.OrgCode, DivisionAuthority.AccountCode, DivisionAuthority.ProgramProjectCode, DivisionAuthority.ProgramProjectName, DivisionAuthority.BocCode, DivisionAuthority.BocName, DivisionAuthority.RcCode, DivisionAuthority.DivisionName, DivisionAuthority.Amount AS
    Authority, Sum(Purchases.OpenCommitments) AS [Open Commitments], Sum(Purchases.Obligations) AS Obligations, 
    Sum([Purchases].[OpenCommitments]+[Purchases].[Obligations]) AS Used, Sum((Authority-Used)) AS Available
INTO [Division Available Balances]
FROM DivisionAuthority 
INNER JOIN Purchases 
ON (DivisionAuthority.RcCode = Purchases.RcCode) AND 
    (DivisionAuthority.AccountCode = Purchases.AccountCode) AND 
    (DivisionAuthority.OrgCode = Purchases.OrgCode) AND 
    (DivisionAuthority.BocCode = Purchases.BocCode) AND 
    (DivisionAuthority.FundCode = Purchases.FundCode) AND 
    (DivisionAuthority.BFY = Purchases.BFY)
GROUP BY DivisionAuthority.PrcId, DivisionAuthority.RPIO, DivisionAuthority.BFY, DivisionAuthority.AhCode, 
    DivisionAuthority.FundCode, DivisionAuthority.FundName, DivisionAuthority.OrgCode, DivisionAuthority.AccountCode, DivisionAuthority.ProgramProjectCode, DivisionAuthority.ProgramProjectName, DivisionAuthority.BocCode, DivisionAuthority.BocName, DivisionAuthority.RcCode, DivisionAuthority.DivisionName, DivisionAuthority.Amount
HAVING (((DivisionAuthority.BocCode) IN ("21","28","36","37","38","41")))
ORDER BY DivisionAuthority.BFY DESC;