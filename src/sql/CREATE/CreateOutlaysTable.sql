SELECT RegionalAuthority.* INTO SiteTravel
FROM RegionalAuthority
WHERE RegionalAuthority.BocCode = "28"
ORDER BY RegionalAuthority.BFY DESC;