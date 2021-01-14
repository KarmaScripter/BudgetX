SELECT RegionalAuthority.* 
INTO Grants
FROM RegionalAuthority
WHERE RegionalAuthority.BocCode = "41"
ORDER BY RegionalAuthority.BFY DESC;
