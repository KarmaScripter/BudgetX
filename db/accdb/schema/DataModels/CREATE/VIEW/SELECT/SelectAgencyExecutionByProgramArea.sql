SELECT StatusOfFunds.BFY, StatusOfFunds.RpioCode, StatusOfFunds.RpioName, StatusOfFunds.FundName, StatusOfFunds.FundCode, StatusOfFunds.ProgramAreaCode, StatusOfFunds.ProgramAreaName, Sum(StatusOfFunds.Amount) AS Authority, Sum(StatusOfFunds.OpenCommitments) AS OpenCommitments, Sum(StatusOfFunds.Obligations) AS Obligations, Sum(StatusOfFunds.Used) AS Used, Sum(StatusOfFunds.Available) AS Available
FROM StatusOfFunds
WHERE StatusOfFunds.BudgetLevel = '7'
GROUP BY StatusOfFunds.BFY, StatusOfFunds.RpioCode, StatusOfFunds.RpioName, 
StatusOfFunds.FundName, StatusOfFunds.FundCode, StatusOfFunds.ProgramAreaCode, StatusOfFunds.ProgramAreaName
ORDER BY StatusOfFunds.BFY DESC;