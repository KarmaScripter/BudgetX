SELECT DISTINCT DivisionAuthority.BFY, DivisionAuthority.DivisionName, DivisionAuthority.ProgramProjectName, Sum(DivisionAuthority.Amount) AS Authority
FROM DivisionAuthority
WHERE DivisionAuthority.Amount>0
GROUP BY DivisionAuthority.DivisionName, DivisionAuthority.BFY, DivisionAuthority.ProgramProjectName
HAVING (((Sum(DivisionAuthority.Amount))>0));
