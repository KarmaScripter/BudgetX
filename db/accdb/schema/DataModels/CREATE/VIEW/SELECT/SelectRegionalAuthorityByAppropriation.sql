SELECT DISTINCT RegionalAuthority.BFY AS BFY, RegionalAuthority.RpioCode, RegionalAuthority.RpioName, 
RegionalAuthority.FundCode, RegionalAuthority.FundName, Sum(CCur([Amount])) AS Amount
FROM RegionalAuthority
GROUP BY RegionalAuthority.BFY, RegionalAuthority.RpioCode, RegionalAuthority.RpioName, 
RegionalAuthority.FundCode, RegionalAuthority.FundName;