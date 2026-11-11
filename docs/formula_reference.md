# Advanced Excel Formula Architecture

### 1. Dynamic Metric Retrieval (`XLOOKUP`)
Retrieves country metadata based on the validated Country Code selector:
```excel
=XLOOKUP($C$4, Raw_WorldBank_Data!$A:$A, Raw_WorldBank_Data!$B:$B, "Country Not Found", 0)
```

### 2. Multi-Condition Aggregation (`SUMIFS`)
Extracts GDP and exchange rates filtered simultaneously by `Country Code` and selected `Decade/Year`:
```excel
=SUMIFS(Raw_WorldBank_Data!$F:$F, Raw_WorldBank_Data!$A:$A, $C$4, Raw_WorldBank_Data!$E:$E, $C$5)
```

### 3. Two-Way Coordinate Lookup (`INDEX / MATCH`)
Retrieves cross-table regional benchmark statistics:
```excel
=INDEX(Regional_Matrix!$B$2:$M$15, MATCH($C$4, Regional_Matrix!$A$2:$A$15, 0), MATCH($C$5, Regional_Matrix!$B$1:$M$1, 0))
```

### 4. Data Validation Parameters
- **Country Code Validation:** Dropdown list referencing `=Raw_WorldBank_Data!$A$2:$A$195`.
- **Year Bucket Validation:** Constrained integers between `1960` and `2020` in 5-year step increments.
