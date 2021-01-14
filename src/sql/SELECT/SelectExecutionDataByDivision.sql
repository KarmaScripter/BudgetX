SELECT DivisionExecution.FundName, DivisionExecution.BFY, DivisionExecution.ProgramProjectName, DivisionExecution.BocName, Sum(DivisionExecution.Authority) AS Authority, Sum(DivisionExecution.[Open Commitments]) AS [SumOfOpen Commitments], Sum(DivisionExecution.Obligations) AS Obligations, Sum(DivisionExecution.Used) AS Used, Sum(DivisionExecution.Available) AS Available, DivisionExecution.DivisionName, DivisionExecution.FundCode, DivisionExecution.BocCode
FROM DivisionExecution
WHERE (((DivisionExecution.Authority)>0))
GROUP BY DivisionExecution.FundName, DivisionExecution.BFY, DivisionExecution.ProgramProjectName, DivisionExecution.BocName, DivisionExecution.DivisionName, DivisionExecution.FundCode, DivisionExecution.BocCode, DivisionExecution.RcCode
HAVING (((DivisionExecution.RcCode)=[Enter a Division RC Code]))
ORDER BY DivisionExecution.BFY DESC;
