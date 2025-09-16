# Dashboard and Read-Only Views

## 1) Upcoming tab - A1
Projects ready for Upcoming when permits are approved.
```
=QUERY(
  'Project Pipeline'!A:BY,
  "select A,B,R,I,AA,AR,J,AQ " &
  "where J = 'Approved' " &
  "order by R asc",
  1
)
```
Columns returned:
- Name, SFID, Deadline, Project Status, Final payment date, Architect, Permits, Equipment

## 2) Framing tab - A1
Projects entering Permitting.
```
=QUERY(
  'Project Pipeline'!A:BY,
  "select A,B,R,AR,AQ " &
  "where I = 'Permitting' " &
  "order by R asc",
  1
)
```
Columns returned:
- Name, SFID, Deadline, Architect, Equipment

---

## 3) Dashboard tab - KPI strip

Create a new tab: `Dashboard`. Use these cells for instant KPIs. Adjust cell references to fit your layout.

### Active Backlog (count)
```
=COUNTIF('Project Pipeline'!BW:BW,TRUE)
```

### Blocked Projects (count)
```
=COUNTIF('Project Pipeline'!AC:AC,">0")
```

### Oldest Blocked Hours (max)
```
=IFERROR(MAX('Project Pipeline'!BX:BX),0)
```

### Upcoming in next 7 days (count)
```
=COUNTIFS('Project Pipeline'!J:J,"Approved",'Project Pipeline'!R:R,">="&TODAY(),'Project Pipeline'!R:R,"<="&TODAY()+7)
```

### Upcoming in next 7 days - Weighted Revenue
```
=SUMIFS('Project Pipeline'!BT:BT,'Project Pipeline'!J:J,"Approved",'Project Pipeline'!R:R,">="&TODAY(),'Project Pipeline'!R:R,"<="&TODAY()+7)
```

### Framing due in next 7 days (count)
```
=COUNTIFS('Project Pipeline'!I:I,"Permitting",'Project Pipeline'!R:R,">="&TODAY(),'Project Pipeline'!R:R,"<="&TODAY()+7)
```

### Overdue Tasks (count)
```
=COUNTIFS(Tasks!D:D,"<>Done",Tasks!G:G,"<"&TODAY())
```

### Docs Required but Missing (count)
```
=COUNTIFS('Project Pipeline'!J:J,"Approved",'Project Pipeline'!AI:AI,FALSE)
```

### Duplicate SFID (count)
```
=COUNTIF('Project Pipeline'!AO:AO,TRUE)
```

### Wins This Week (count) - using Created Month proxy
```
=COUNTIFS('Project Pipeline'!P:P,"Salesforce",'Project Pipeline'!BP:BP,TEXT(TODAY(),"yyyy-mm"))
```

### Revenue Forecast this month (weighted)
```
=SUMIFS('Project Pipeline'!BT:BT,'Project Pipeline'!BO:BO,TEXT(TODAY(),"yyyy-mm"))
```

### Revenue Forecast next month (weighted)
```
=SUMIFS('Project Pipeline'!BT:BT,'Project Pipeline'!BO:BO,TEXT(EDATE(TODAY(),1),"yyyy-mm"))
```

---

## 4) Forecast by Month table
```
=QUERY(
  'Project Pipeline'!A:BY,
  "select BO, sum(AS), sum(BT) where BO is not null group by BO order by BO",
  1
)
```
Add a chart with Month on X, sums on Y.

---

## 5) Permit Pipeline - counts by permit status
```
=QUERY(
  'Project Pipeline'!A:BY,
  "select J, count(A) where J is not null group by J label count(A) 'Count'",
  1
)
```

## 6) Permit Pipeline - average days in permitting
```
=QUERY(
  'Project Pipeline'!A:BY,
  "select avg(BQ) where BQ is not null label avg(BQ) 'Avg days in permitting'",
  1
)
```

---

## 7) Scheduling Readiness - Ready to Schedule
```
=QUERY(
  'Project Pipeline'!A:BY,
  "select A,B,AI,AS,L,BF where I='Permitting' and J='Approved' and AI = TRUE and BF = 100 and AS is not null and L is not null order by R asc",
  1
)
```

## 8) Scheduling Readiness - Not Ready (Fix These)
```
=QUERY(
  'Project Pipeline'!A:BY,
  "select A,B,AI,AS,L,BF where I='Permitting' and J='Approved' and (AI = FALSE or BF < 100 or AS is null or L is null) order by R asc",
  1
)
```

---

## 9) Execution Watchlist
```
=QUERY(
  'Project Pipeline'!A:BY,
  "select A,B,BL,BV,BB,Q,H where (BW = TRUE or AC is not null) order by AC desc, R asc",
  1
)
```
