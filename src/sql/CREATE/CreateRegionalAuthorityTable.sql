SELECT RegionalAuthority.* 
INTO WorkingCapitalFund
FROM RegionalAuthority
WHERE RegionalAuthority.BocCode = "38"
ORDER BY RegionalAuthority.BFY DESC;
