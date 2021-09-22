SELECT StatusOfFunds.BFY, StatusOfFunds.RpioCode, StatusOfFunds.RpioName, StatusOfFunds.FundCode, StatusOfFunds.FundName, CCur(Sum(StatusOfFunds.Amount)) AS Amount, CCur(Sum(StatusOfFunds.OpenCommitments)) AS OpenCommitments, CCur(Sum(StatusOfFunds.Obligations)) AS Obligations, CCur(Sum(StatusOfFunds.Used)) AS Used, CCur(Sum(StatusOfFunds.Available)) AS Available
FROM StatusOfFunds
WHERE StatusOfFunds.BudgetLevel = '7'
AND StatusOfFunds.RpioCode NOT IN ('01', '02', '03', '04', '05', '06', '07', '08', '09', '10')
GROUP BY StatusOfFunds.BFY, StatusOfFunds.RpioCode, StatusOfFunds.RpioName, StatusOfFunds.FundCode, 
StatusOfFunds.FundName
ORDER BY StatusOfFunds.BFY DESC , StatusOfFunds.FundCode;
