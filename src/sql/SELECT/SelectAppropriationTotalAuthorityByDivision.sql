SELECT DISTINCT DivisionAuthority.BFY, DivisionAuthority.DivisionName, DivisionAuthority.FundName, 
    SUM(DivisionAuthority.Amount) AS Authority
FROM DivisionAuthority
WHERE DivisionAuthority.RcCode = [Enter a RC Code]
GROUP BY DivisionAuthority.BFY, DivisionAuthority.DivisionName, DivisionAuthority.FundName;