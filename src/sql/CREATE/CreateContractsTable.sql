SELECT RegionalAuthority.* 
INTO Contracts
FROM RegionalAuthority
WHERE RegionalAuthority.BocCode = "37"
ORDER BY RegionalAuthority.BFY DESC;
