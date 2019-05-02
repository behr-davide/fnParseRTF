# fnParseRTF

## Usage
```sql
select 
  id 
  ,rtfData as raw
  ,dbo.fnParseRTF(rtfData) as clean
from dbo.myRTFTable
```
