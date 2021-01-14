SELECT RegionalAuthority.* 
INTO Travel
FROM RegionalAuthority
WHERE RegionalAuthority.BocCode = "21"
ORDER BY RegionalAuthority.BFY DESC;
