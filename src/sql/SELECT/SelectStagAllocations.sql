SELECT Allocations.*
FROM Allocations
WHERE Allocations.FundCode LIKE "T%" AND 
    Allocations.AhCode <> "06";
