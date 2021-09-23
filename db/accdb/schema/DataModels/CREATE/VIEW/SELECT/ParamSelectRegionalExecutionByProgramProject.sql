UPDATE FullTimeEquivalents
INNER JOIN Allocations
ON FullTimeEquivalents.AhCode = Allocations.AhCode
SET FullTimeEquivalents.AhName = Allocations.AhName
WHERE FullTimeEquivalents.AhName <> Allocations.AhName;