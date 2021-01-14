SELECT RegionalAuthority.* 
INTO Expenses
FROM RegionalAuthority
WHERE RegionalAuthority.BocCode = "36"
ORDER BY RegionalAuthority.BFY DESC;
