# Google Sheets Formulas

Paste these into the specified columns. Assumes the column layout from the README. Replace `ROW` with the current row number if your editor does not auto-fill.

> Tasks rollups use `SFID` on Project Pipeline and `Project SFID` on Tasks.

## Project Pipeline

### H: Slack Channel URL
```
=IF(OR(ISBLANK(G2),ISBLANK(C2)),"","https://app.slack.com/client/"&G2&"/"&C2)
```

### AY: Gross Margin %
```
=IF(ISBLANK(AS2),"",ROUND((AS2-AT2)/AS2*100))
```

### AZ: Week of (Monday)
```
=IF(ISBLANK(R2),"",R2-WEEKDAY(R2,2)+1)
```

### BA: open_tasks_count
```
=IFERROR(COUNTIFS(Tasks!B:B,$B2,Tasks!D:D,"<>Done"),0)
```

### BB: overdue_tasks_count
```
=IFERROR(COUNTIFS(Tasks!B:B,$B2,Tasks!D:D,"<>Done",Tasks!G:G,"<"&TODAY()),0)
```

### BC: completed_tasks_count
```
=IFERROR(COUNTIFS(Tasks!B:B,$B2,Tasks!D:D,"Done"),0)
```

### BD: total_blocking_tasks
```
=IFERROR(SUMIFS(Tasks!K:K,Tasks!B:B,$B2),0)
```

### BE: completed_blocking_tasks
```
=IFERROR(SUMIFS(Tasks!K:K,Tasks!B:B,$B2,Tasks!D:D,"Done"),0)
```

### BF: task_progress_%  (blocking tasks only)
```
=IF(BD2=0,0,ROUND(100*BE2/BD2))
```

### BG: can_advance_globally
```
=AND(NOT(ISBLANK(R2)),NOT(ISBLANK(C2)),BB2=0,NOT(AO2))
```

### BH: can_advance_to_Permitting
```
=AND(Y2,AI2,BF2=100)
```

### BI: can_advance_to_Scheduled
```
=AND(J2="Approved",AI2,NOT(ISBLANK(AS2)),NOT(ISBLANK(L2)),BF2=100)
```

### BJ: can_advance_to_Inspections
```
=AND(AK2,AL2,AM2,BF2=100)
```

### BK: can_advance_to_Done
```
=AND(AN2,Z2,AI2,BF2=100)
```

### BL: Advance block reason (custom function)
```
=GET_ADVANCE_BLOCK_REASON(I2,BG2,BH2,BI2,BJ2,BK2)
```

### BN: escalate_ready (24h)
```
=AND(NOT(ISBLANK(AC2)),(NOW()-AC2)>=1)
```

### BO: Month (Deadline)
```
=IF(ISBLANK(R2),"",TEXT(R2,"yyyy-mm"))
```

### BP: Created Month
```
=TEXT(NOW(),"yyyy-mm")
```

### BQ: days_in_permitting
```
=IF(OR(ISBLANK(T2),ISBLANK(S2)),"",S2-T2)
```

### BR: days_to_schedule
```
=IF(OR(ISBLANK(S2),ISBLANK(AE2)),"",AE2-S2)
```

### BS: lead_time_days (Closed-Won → Done)
```
=IF(OR(ISBLANK(AG2),ISBLANK(AF2)),"",AF2-AG2)
```

### BT: Revenue Weighted
```
=IF(OR(ISBLANK(AS2),ISBLANK(L2)),"",AS2*VALUE(L2)/100)
```

### BU: docs_required_but_missing
```
=AND(J2="Approved",NOT(AI2))
```

### BV: aging_days_since_edit
```
=IFERROR(INT((NOW()-X2)),0)
```

### BW: is_active_backlog
```
=OR(I2="Scheduled",I2="Permitting",I2="Inspections")
```

### BX: blocked_hours (optional)
```
=IF(ISBLANK(AC2),"",24*(NOW()-AC2))
```

### BY: staleness_flag (optional)
```
=BV2>=7
```

---

## Tasks

### K: Counts toward completion
```
=IF(OR(D2="Waiting External",D2="Skipped"),0,1)
```

### L: Completed %
```
=IF(D2="Done",100,0)
```
