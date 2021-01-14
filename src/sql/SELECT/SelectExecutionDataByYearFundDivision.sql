SELECT DISTINCTROW DivisionAuthority.RPIO, DivisionAuthority.BFY, DivisionAuthority.AhCode, 
    DivisionAuthority.FundCode, DivisionAuthority.FundName, DivisionAuthority.OrgCode, DivisionAuthority.AccountCode, DivisionAuthority.ProgramProjectName AS ProgramProjectName, DivisionAuthority.BocCode, DivisionAuthority.BocName, DivisionAuthority.RcCode, DivisionAuthority.DivisionName, Sum(DivisionAuthority.Amount) AS Authority, Sum(Purchases.OpenCommitments) AS OpenCommitments, Sum(Purchases.Obligations) AS Obligations, Sum(([Purchases].[OpenCommitments]+[Purchases].[Obligations])) AS Used, Sum((DivisionAuthority.Amount-Purchases.[OpenCommitments]-Purchases.[Obligations])) AS Available
FROM DivisionAuthority 
INNER JOIN Purchases 
ON (DivisionAuthority.BFY = Purchases.BFY) AND 
    (DivisionAuthority.FundCode = Purchases.FundCode) AND 
    (DivisionAuthority.AccountCode = Purchases.AccountCode) AND 
    (DivisionAuthority.BocCode = Purchases.BocCode) AND 
    (DivisionAuthority.RcCode = Purchases.RcCode)
GROUP BY DivisionAuthority.RPIO, DivisionAuthority.BFY, DivisionAuthority.AhCode, DivisionAuthority.FundCode, 
    DivisionAuthority.FundName, DivisionAuthority.OrgCode, DivisionAuthority.AccountCode, DivisionAuthority.ProgramProjectName, DivisionAuthority.BocCode, DivisionAuthority.BocName, DivisionAuthority.RcCode, DivisionAuthority.DivisionName, DivisionAuthority.Amount
HAVING (((DivisionAuthority.BFY)=[Enter a Fiscal Year]) AND 
    ((DivisionAuthority.FundCode)=[Enter a Fund Code]) AND 
    ((DivisionAuthority.RcCode)=[Enter a RC Code]));
