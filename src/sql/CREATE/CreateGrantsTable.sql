SELECT RegionalAuthority.* 
INTO Payroll
FROM RegionalAuthority
WHERE RegionalAuthority.BocCode = "10"
ORDER BY RegionalAuthority.BFY DESC;
