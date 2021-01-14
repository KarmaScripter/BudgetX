SELECT RegionalAuthority.* 
INTO FTE
FROM RegionalAuthority
WHERE RegionalAuthority.BocCode = "17"
ORDER BY RegionalAuthority.BFY DESC;
