SELECT DISTINCT DivisionAuthority.BFY, DivisionAuthority.DivisionName, DivisionAuthority.FundName, Sum(DivisionAuthority.Amount) AS Authority
FROM DivisionAuthority
WHERE DivisionAuthority.Amount>0
GROUP BY DivisionAuthority.BFY, DivisionAuthority.DivisionName, DivisionAuthority.FundName
HAVING (((Sum(DivisionAuthority.Amount))>0));