PARAMETERS RpioCodeArgs LongText;
SELECT StatusOfFunds.BFY, StatusOfFunds.RpioCode, StatusOfFunds.RpioName, StatusOfFunds.ProgramProjectName, StatusOfFunds.AccountCode, Sum(StatusOfFunds.Amount) AS Authority, Sum(StatusOfFunds.OpenCommitments) AS [Open Commitments], Sum(StatusOfFunds.Obligations) AS Obligations, Sum(StatusOfFunds.Used) AS Used, StatusOfFunds.Available
FROM StatusOfFunds
WHERE (((StatusOfFunds.Amount)>0)) 
AND StatusOfFunds.RpioCode = [RpioCodeArgs]
GROUP BY StatusOfFunds.BFY, StatusOfFunds.RpioCode, StatusOfFunds.RpioName, StatusOfFunds.ProgramProjectName, StatusOfFunds.AccountCode, StatusOfFunds.Available
ORDER BY StatusOfFunds.BFY DESC;