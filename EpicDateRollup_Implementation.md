# Epic Date Rollup — Implementation Guide

## Architecture Decision

The existing `_DeliverableData` sheet is the **SharePoint master source** fed by `qry_SPO_Deliverables`. To preserve that single source of truth, this implementation:

1. Creates a **new Power Query** (`qry_ADO_EpicData`) that pulls Epics and merges aggregated Feature dates
2. Outputs to a **new hidden sheet** (`_EpicData`) with an Excel table (`EpicDataTable`)
3. Adds **calculated columns** to `_DeliverableData` that XLOOKUP into `_EpicData` via `DevOpsID`

This keeps SharePoint as the deliverable master while layering on DevOps date intelligence.

```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────────┐
│ qry_ADO_Epics   │     │ qry_ADO_Features │     │ qry_SPO_Deliverables│
│ (OData v2)      │     │ DateSummary      │     │ (SharePoint)        │
│                 │     │ (OData v2)       │     │                     │
│ EpicID, Title,  │     │ ParentEpicID,    │     │ DeliverableID,      │
│ State, AreaPath,│     │ MIN(StartDate),  │     │ DevOpsID, Status,   │
│ StartDate,      │     │ MAX(TargetDate), │     │ Programme, etc.     │
│ TargetDate,     │     │ FeatureCount     │     │                     │
│ OriginalStart,  │     │                  │     │                     │
│ OriginalEnd,    │     │                  │     │                     │
│ Priority, Effort│     │                  │     │                     │
└────────┬────────┘     └────────┬─────────┘     └──────────┬──────────┘
         │                       │                           │
         │  Table.NestedJoin     │                           │
         └───────────┬───────────┘                           │
                     ▼                                       ▼
            ┌─────────────────┐                    ┌─────────────────┐
            │ _EpicData sheet │                    │ _DeliverableData│
            │ (EpicDataTable) │◄── XLOOKUP ───────│ + Calc Columns  │
            │                 │    via DevOpsID    │                 │
            └─────────────────┘                    └─────────────────┘
```

---

## DevOps Field Reference

### Epic OData Properties Used

| DevOps Property Name | OData Select Name | Renamed To (in EpicDataTable) | Description |
|----------------------|-------------------|-------------------------------|-------------|
| WorkItemId | WorkItemId | EpicID | Unique Epic identifier |
| Title | Title | EpicTitle | Epic display name |
| State | State | EpicState | New, Active, Resolved, Closed |
| AreaPath | AreaPath | AreaPath | Programme/team hierarchy |
| IterationPath | IterationPath | IterationPath | Sprint/iteration context |
| StartDate | StartDate | EpicStartDate | Current planned start |
| TargetDate | TargetDate | EpicTargetDate | Current planned end |
| Custom_OriginalStartDate | Custom_OriginalStartDate | OriginalStartDate | Baselined original start |
| Custom_OriginalEndDate | Custom_OriginalEndDate | OriginalEndDate | Baselined original end |
| Priority | Priority | EpicPriority | Epic priority (1-4) |
| Effort | Effort | EpicEffort | Estimated effort value |
| Custom_KeyDeliverable | Custom_KeyDeliverable | KeyDeliverable | Key deliverable flag |
| Custom_CriticalDate | Custom_CriticalDate | CriticalDate | Critical/immovable date |
| ChangedDate | ChangedDate | EpicLastUpdated | Last modification timestamp |
| AssignedTo | AssignedTo($expand) | EpicOwner | Assigned person display name |

### Feature OData Properties Used (in FeatureDateSummary)

| DevOps Property Name | Purpose in Aggregation |
|----------------------|----------------------|
| WorkItemId | Feature identifier |
| Title | Feature name (for debug/tracing) |
| State | Filter active vs Done/Closed |
| StartDate | MIN() → RolledUpStart |
| TargetDate | MAX() → RolledUpEnd |
| Parent.WorkItemId | Join key to parent Epic |

---

## 1. Power Query Implementation

### 1.1 Helper Query: `qry_ADO_FeatureDateSummary` (Connection Only)

This query aggregates Feature dates by parent Epic. Set to **Connection Only** — it feeds into the main Epic query, not a sheet.

```powerquery
let
    // Read config from Excel
    DevOpsOrg = Excel.CurrentWorkbook(){[Name="ConfigTable"]}[Content]{0}[DevOpsOrg],
    DevOpsProject = Excel.CurrentWorkbook(){[Name="ConfigTable"]}[Content]{0}[DevOpsProject],

    BaseURL = "https://analytics.dev.azure.com/" & DevOpsOrg & "/" & DevOpsProject &
              "/_odata/v4.0-preview/WorkItems",

    // Pull Features with Parent reference and dates
    Source = OData.Feed(
        BaseURL & "?" &
        "$filter=WorkItemType eq 'Feature' and State ne 'Removed'" &
        "&$select=WorkItemId,Title,State,StartDate,TargetDate" &
        "&$expand=Parent($select=WorkItemId)",
        null,
        [Implementation="2.0"]
    ),

    // Extract ParentEpicID from expanded Parent record
    ExpandParent = Table.ExpandRecordColumn(Source, "Parent", {"WorkItemId"}, {"ParentEpicID"}),

    // Remove Features with no parent Epic (orphans)
    FilterHasParent = Table.SelectRows(ExpandParent, each [ParentEpicID] <> null),

    // Ensure date columns are typed correctly
    TypedDates = Table.TransformColumnTypes(FilterHasParent, {
        {"StartDate", type date},
        {"TargetDate", type date}
    }),

    // Group by ParentEpicID and aggregate
    Grouped = Table.Group(TypedDates, {"ParentEpicID"}, {
        // Earliest Feature start = rolled-up Epic start
        {"RolledUpStart", each List.Min([StartDate]), type nullable date},

        // Latest Feature end = rolled-up Epic end
        {"RolledUpEnd", each List.Max([TargetDate]), type nullable date},

        // Count of linked Features
        {"FeatureCount", each Table.RowCount(_), Int64.Type},

        // Count of Features with dates populated (data quality indicator)
        {"FeaturesWithDates", each Table.RowCount(
            Table.SelectRows(_, each [StartDate] <> null and [TargetDate] <> null)
        ), Int64.Type},

        // Earliest Feature start among active (non-Done) Features only
        {"ActiveRolledUpStart", each List.Min(
            Table.SelectRows(_, each [State] <> "Done" and [State] <> "Closed")[StartDate]
        ), type nullable date},

        // Latest Feature end among active (non-Done) Features only
        {"ActiveRolledUpEnd", each List.Max(
            Table.SelectRows(_, each [State] <> "Done" and [State] <> "Closed")[TargetDate]
        ), type nullable date}
    })
in
    Grouped
```

**Why two sets of rolled-up dates?** The `RolledUpStart`/`RolledUpEnd` pair gives the full historical span across all Features. The `ActiveRolledUpStart`/`ActiveRolledUpEnd` pair excludes completed Features, giving you the current "remaining work window." Both are useful for different reporting scenarios.

---

### 1.2 Main Query: `qry_ADO_EpicData` (Loads to `_EpicData` Sheet)

This query pulls Epic-level fields directly, then merges the aggregated Feature dates.

```powerquery
let
    // Read config from Excel
    DevOpsOrg = Excel.CurrentWorkbook(){[Name="ConfigTable"]}[Content]{0}[DevOpsOrg],
    DevOpsProject = Excel.CurrentWorkbook(){[Name="ConfigTable"]}[Content]{0}[DevOpsProject],

    BaseURL = "https://analytics.dev.azure.com/" & DevOpsOrg & "/" & DevOpsProject &
              "/_odata/v4.0-preview/WorkItems",

    // Pull Epics with all required fields including custom properties
    Source = OData.Feed(
        BaseURL & "?" &
        "$filter=WorkItemType eq 'Epic' and State ne 'Removed'" &
        "&$select=WorkItemId,Title,State,AreaPath,IterationPath," &
        "StartDate,TargetDate," &
        "Custom_OriginalStartDate,Custom_OriginalEndDate," &
        "Priority,Effort," &
        "Custom_KeyDeliverable,Custom_CriticalDate," &
        "ChangedDate" &
        "&$expand=AssignedTo($select=UserName,UserEmail)",
        null,
        [Implementation="2.0"]
    ),

    // Flatten AssignedTo
    ExpandAssignedTo = Table.ExpandRecordColumn(Source, "AssignedTo",
        {"UserName"}, {"EpicOwner"}),

    // Rename for clarity
    RenamedColumns = Table.RenameColumns(ExpandAssignedTo, {
        {"WorkItemId", "EpicID"},
        {"Title", "EpicTitle"},
        {"State", "EpicState"},
        {"StartDate", "EpicStartDate"},
        {"TargetDate", "EpicTargetDate"},
        {"Custom_OriginalStartDate", "OriginalStartDate"},
        {"Custom_OriginalEndDate", "OriginalEndDate"},
        {"Priority", "EpicPriority"},
        {"Effort", "EpicEffort"},
        {"Custom_KeyDeliverable", "KeyDeliverable"},
        {"Custom_CriticalDate", "CriticalDate"},
        {"ChangedDate", "EpicLastUpdated"}
    }),

    // Type the date columns
    TypedDates = Table.TransformColumnTypes(RenamedColumns, {
        {"EpicStartDate", type date},
        {"EpicTargetDate", type date},
        {"OriginalStartDate", type date},
        {"OriginalEndDate", type date},
        {"CriticalDate", type date},
        {"EpicLastUpdated", type datetime}
    }),

    // Merge with Feature date summary
    MergedFeatureDates = Table.NestedJoin(
        TypedDates, {"EpicID"},
        qry_ADO_FeatureDateSummary, {"ParentEpicID"},
        "FeatureSummary",
        JoinKind.LeftOuter
    ),

    // Expand the merged columns
    ExpandedFeatureDates = Table.ExpandTableColumn(MergedFeatureDates, "FeatureSummary", {
        "RolledUpStart", "RolledUpEnd", "FeatureCount", "FeaturesWithDates",
        "ActiveRolledUpStart", "ActiveRolledUpEnd"
    }),

    // Add DevOps hyperlink
    AddDevOpsLink = Table.AddColumn(ExpandedFeatureDates, "EpicDevOpsLink",
        each "https://dev.azure.com/" & DevOpsOrg & "/" & DevOpsProject &
             "/_workitems/edit/" & Text.From([EpicID]),
        type text),

    // Variance: RolledUp vs Current Epic dates (schedule drift from Features)
    AddStartVarianceCurrent = Table.AddColumn(AddDevOpsLink, "StartDateVariance",
        each if [EpicStartDate] <> null and [RolledUpStart] <> null
             then Duration.Days([RolledUpStart] - [EpicStartDate])
             else null,
        type nullable number),

    AddEndVarianceCurrent = Table.AddColumn(AddStartVarianceCurrent, "EndDateVariance",
        each if [EpicTargetDate] <> null and [RolledUpEnd] <> null
             then Duration.Days([RolledUpEnd] - [EpicTargetDate])
             else null,
        type nullable number),

    // Variance: Current Epic dates vs Original baseline (Epic-level schedule movement)
    AddOrigStartVariance = Table.AddColumn(AddEndVarianceCurrent, "OriginalStartVariance",
        each if [OriginalStartDate] <> null and [EpicStartDate] <> null
             then Duration.Days([EpicStartDate] - [OriginalStartDate])
             else null,
        type nullable number),

    AddOrigEndVariance = Table.AddColumn(AddOrigStartVariance, "OriginalEndVariance",
        each if [OriginalEndDate] <> null and [EpicTargetDate] <> null
             then Duration.Days([EpicTargetDate] - [OriginalEndDate])
             else null,
        type nullable number),

    // Replace nulls in counts with 0
    ReplacedNulls = Table.ReplaceValue(AddOrigEndVariance, null, 0,
        Replacer.ReplaceValue, {"FeatureCount", "FeaturesWithDates"}),

    // Sort by AreaPath then EpicTitle
    Sorted = Table.Sort(ReplacedNulls, {
        {"AreaPath", Order.Ascending},
        {"EpicTitle", Order.Ascending}
    })
in
    Sorted
```

---

### 1.3 Power Query Setup Steps

1. **Open Power Query Editor** — Data → Get Data → Launch Power Query Editor

2. **Create `qry_ADO_FeatureDateSummary`:**
   - New Query → Blank Query → paste the M code from section 1.1
   - Right-click the query → **Properties** → rename to `qry_ADO_FeatureDateSummary`
   - Right-click → **Load To...** → select **Only Create Connection**

3. **Create `qry_ADO_EpicData`:**
   - New Query → Blank Query → paste the M code from section 1.2
   - Right-click → **Properties** → rename to `qry_ADO_EpicData`
   - Right-click → **Load To...** → select **Table** and choose **New worksheet**
   - Rename the sheet to `_EpicData` and hide it

4. **Privacy Levels** — Ensure Azure DevOps Analytics OData source is set to **Organisational** (consistent with your existing `qry_DevOps_Milestones` setup).

5. **Name the output table** — Click into the loaded table on `_EpicData`, then in the Table Design tab rename it to `EpicDataTable`.

---

## 2. `_EpicData` Sheet Schema (EpicDataTable)

The Power Query output produces this table structure:

### Direct from DevOps (Epic-level fields)

| Column | Type | DevOps Property | Description |
|--------|------|-----------------|-------------|
| EpicID | Number | WorkItemId | Links to `_DeliverableData.DevOpsID` |
| EpicTitle | Text | Title | Epic name |
| EpicState | Text | State | New, Active, Resolved, Closed |
| AreaPath | Text | AreaPath | Programme/team hierarchy |
| IterationPath | Text | IterationPath | Sprint/iteration context |
| EpicStartDate | Date | StartDate | Current planned start |
| EpicTargetDate | Date | TargetDate | Current planned end |
| OriginalStartDate | Date | Custom_OriginalStartDate | Baselined original start |
| OriginalEndDate | Date | Custom_OriginalEndDate | Baselined original end |
| EpicPriority | Number | Priority | Priority level (1-4) |
| EpicEffort | Number | Effort | Estimated effort value |
| KeyDeliverable | Text/Bool | Custom_KeyDeliverable | Key deliverable flag |
| CriticalDate | Date | Custom_CriticalDate | Critical/immovable target date |
| EpicLastUpdated | DateTime | ChangedDate | Last change timestamp |
| EpicOwner | Text | AssignedTo.UserName | Assigned person display name |

### Calculated from Feature Aggregation

| Column | Type | Calculation | Description |
|--------|------|-------------|-------------|
| RolledUpStart | Date | MIN(Feature.StartDate) | Earliest Feature start across all child Features |
| RolledUpEnd | Date | MAX(Feature.TargetDate) | Latest Feature end across all child Features |
| FeatureCount | Number | COUNT(Features) | Total linked Features |
| FeaturesWithDates | Number | COUNT(Features where both dates populated) | Data quality indicator |
| ActiveRolledUpStart | Date | MIN(StartDate) for non-Done Features | Remaining work window start |
| ActiveRolledUpEnd | Date | MAX(TargetDate) for non-Done Features | Remaining work window end |

### Calculated Variance & Links

| Column | Type | Calculation | Description |
|--------|------|-------------|-------------|
| EpicDevOpsLink | Text | URL construction | Clickable hyperlink to DevOps work item |
| StartDateVariance | Number | RolledUpStart − EpicStartDate (days) | Feature drift from Epic start (+ve = Features start later) |
| EndDateVariance | Number | RolledUpEnd − EpicTargetDate (days) | Feature drift from Epic end (+ve = Features end later) |
| OriginalStartVariance | Number | EpicStartDate − OriginalStartDate (days) | Epic start drift from baseline (+ve = slipped) |
| OriginalEndVariance | Number | EpicTargetDate − OriginalEndDate (days) | Epic end drift from baseline (+ve = slipped) |

---

## 3. Calculated Columns on `_DeliverableData`

Add these columns to `_DeliverableData` to enrich deliverables with their parent Epic's data. These use XLOOKUP against `EpicDataTable` via the `DevOpsID` field.

### 3.1 Column Definitions

Add these as new columns at the end of the `_DeliverableData` table:

**Column: EpicRolledUpStart**
```excel
=IFERROR(
    XLOOKUP([@DevOpsID], EpicDataTable[EpicID], EpicDataTable[RolledUpStart], ""),
    ""
)
```

**Column: EpicRolledUpEnd**
```excel
=IFERROR(
    XLOOKUP([@DevOpsID], EpicDataTable[EpicID], EpicDataTable[RolledUpEnd], ""),
    ""
)
```

**Column: EpicActiveStart**
```excel
=IFERROR(
    XLOOKUP([@DevOpsID], EpicDataTable[EpicID], EpicDataTable[ActiveRolledUpStart], ""),
    ""
)
```

**Column: EpicActiveEnd**
```excel
=IFERROR(
    XLOOKUP([@DevOpsID], EpicDataTable[EpicID], EpicDataTable[ActiveRolledUpEnd], ""),
    ""
)
```

**Column: EpicOriginalStart**
```excel
=IFERROR(
    XLOOKUP([@DevOpsID], EpicDataTable[EpicID], EpicDataTable[OriginalStartDate], ""),
    ""
)
```

**Column: EpicOriginalEnd**
```excel
=IFERROR(
    XLOOKUP([@DevOpsID], EpicDataTable[EpicID], EpicDataTable[OriginalEndDate], ""),
    ""
)
```

**Column: EpicCriticalDate**
```excel
=IFERROR(
    XLOOKUP([@DevOpsID], EpicDataTable[EpicID], EpicDataTable[CriticalDate], ""),
    ""
)
```

**Column: EpicKeyDeliverable**
```excel
=IFERROR(
    XLOOKUP([@DevOpsID], EpicDataTable[EpicID], EpicDataTable[KeyDeliverable], ""),
    ""
)
```

**Column: EpicPriority**
```excel
=IFERROR(
    XLOOKUP([@DevOpsID], EpicDataTable[EpicID], EpicDataTable[EpicPriority], ""),
    ""
)
```

**Column: EpicEffort**
```excel
=IFERROR(
    XLOOKUP([@DevOpsID], EpicDataTable[EpicID], EpicDataTable[EpicEffort], ""),
    ""
)
```

**Column: EpicFeatureCount**
```excel
=IFERROR(
    XLOOKUP([@DevOpsID], EpicDataTable[EpicID], EpicDataTable[FeatureCount], 0),
    0
)
```

**Column: EpicStartVariance** (days — positive = Features start later than Epic)
```excel
=IFERROR(
    XLOOKUP([@DevOpsID], EpicDataTable[EpicID], EpicDataTable[StartDateVariance], ""),
    ""
)
```

**Column: EpicEndVariance** (days — positive = Features end later than Epic)
```excel
=IFERROR(
    XLOOKUP([@DevOpsID], EpicDataTable[EpicID], EpicDataTable[EndDateVariance], ""),
    ""
)
```

**Column: EpicOrigStartVariance** (days — positive = Epic start slipped from baseline)
```excel
=IFERROR(
    XLOOKUP([@DevOpsID], EpicDataTable[EpicID], EpicDataTable[OriginalStartVariance], ""),
    ""
)
```

**Column: EpicOrigEndVariance** (days — positive = Epic end slipped from baseline)
```excel
=IFERROR(
    XLOOKUP([@DevOpsID], EpicDataTable[EpicID], EpicDataTable[OriginalEndVariance], ""),
    ""
)
```

**Column: EpicDateHealth** (traffic light indicator)
```excel
=LET(
    _devOpsID, [@DevOpsID],
    _endVar, IFERROR(XLOOKUP(_devOpsID, EpicDataTable[EpicID], EpicDataTable[EndDateVariance]), ""),
    _featureCount, IFERROR(XLOOKUP(_devOpsID, EpicDataTable[EpicID], EpicDataTable[FeatureCount]), 0),
    _featuresWithDates, IFERROR(XLOOKUP(_devOpsID, EpicDataTable[EpicID], EpicDataTable[FeaturesWithDates]), 0),

    IF(_devOpsID = "", "",
    IF(_featureCount = 0, "No Features",
    IF(_featuresWithDates < _featureCount, "⚠ Missing Dates",
    IF(_endVar = "", "N/A",
    IF(_endVar > 14, "🔴 Slipping",
    IF(_endVar > 0, "🟡 At Risk",
    "🟢 On Track"))))))
)
```

**Column: CriticalDateAlert** (flags if rolled-up end breaches the Critical Date)
```excel
=LET(
    _devOpsID, [@DevOpsID],
    _critDate, IFERROR(XLOOKUP(_devOpsID, EpicDataTable[EpicID], EpicDataTable[CriticalDate]), ""),
    _rolledEnd, IFERROR(XLOOKUP(_devOpsID, EpicDataTable[EpicID], EpicDataTable[RolledUpEnd]), ""),

    IF(OR(_critDate = "", _rolledEnd = ""), "",
    IF(_rolledEnd > _critDate, "🔴 CRITICAL DATE AT RISK",
    IF(_rolledEnd > _critDate - 14, "🟡 Approaching Critical Date",
    "")))
)
```

### 3.2 Adding the Columns

In Excel:
1. Click into the `_DeliverableData` table
2. Navigate to the last column
3. Type the header name in the next empty column header cell (e.g., `EpicRolledUpStart`)
4. Enter the formula in the first data row — it will auto-fill down the table
5. Format date columns as `DD/MM/YYYY`
6. Format variance columns as `0` (whole number of days)

### 3.3 Conditional Formatting

**EpicDateHealth column:**

| Condition | Format |
|-----------|--------|
| Cell contains "🔴 Slipping" | Red fill (#FFC7CE), dark red text (#9C0006) |
| Cell contains "🟡 At Risk" | Orange fill (#FFEB9C), dark orange text (#9C6500) |
| Cell contains "🟢 On Track" | Green fill (#C6EFCE), dark green text (#006100) |
| Cell contains "⚠ Missing Dates" | Grey fill (#D9D9D9), dark grey text (#808080) |

**CriticalDateAlert column:**

| Condition | Format |
|-----------|--------|
| Cell contains "🔴 CRITICAL" | Red fill (#FFC7CE), dark red text (#9C0006), **Bold** |
| Cell contains "🟡 Approaching" | Orange fill (#FFEB9C), dark orange text (#9C6500) |

**Variance columns (EpicStartVariance, EpicEndVariance, EpicOrigStartVariance, EpicOrigEndVariance):**

| Condition | Format |
|-----------|--------|
| Value > 14 | Red text (#9C0006) |
| Value > 0 | Orange text (#9C6500) |
| Value = 0 | Green text (#006100) |
| Value < 0 | Blue text (#0070C0) — ahead of schedule |

---

## 4. How It All Flows Together

```
1. PM sets Feature dates in Azure DevOps
   Feature 1: Start 01/01/26, End 30/01/26
   Feature 2: Start 01/02/26, End 28/02/26
   Feature 3: Start 01/03/26, End 30/03/26
        │
2. Data Refresh (manual or Power Automate scheduled)
        │
3. qry_ADO_FeatureDateSummary aggregates:
   Epic 123 → RolledUpStart: 01/01/26, RolledUpEnd: 30/03/26, FeatureCount: 3
        │
4. qry_ADO_EpicData merges Epic fields + Feature summary
   Epic 123 → OriginalStartDate: 15/12/25, OriginalEndDate: 15/03/26
              EpicStartDate: 01/01/26, EpicTargetDate: 28/03/26
              CriticalDate: 31/03/26, KeyDeliverable: Yes
              RolledUpStart: 01/01/26, RolledUpEnd: 30/03/26
              EndDateVariance: +2 days, OriginalEndVariance: +13 days
   → Loads to _EpicData sheet (EpicDataTable)
        │
5. _DeliverableData calculated columns XLOOKUP via DevOpsID
   DEL-001 (DevOpsID=123):
     EpicRolledUpStart: 01/01/26
     EpicRolledUpEnd: 30/03/26
     EpicOriginalStart: 15/12/25
     EpicOriginalEnd: 15/03/26
     EpicCriticalDate: 31/03/26
     EpicDateHealth: 🟡 At Risk
     CriticalDateAlert: 🟡 Approaching Critical Date
        │
6. Dashboard, Milestones, and other sheets can reference
   the enriched _DeliverableData columns
```

---

## 5. Integration with Existing Sheets

### Dashboard
Reference the new columns for KPI cards:

```excel
' Count of deliverables with slipping Epics
=COUNTIF(_DeliverableData[EpicDateHealth], "*Slipping*")

' Count of deliverables with missing Feature dates
=COUNTIF(_DeliverableData[EpicDateHealth], "*Missing*")

' Count of deliverables with Critical Date at risk
=COUNTIF(_DeliverableData[CriticalDateAlert], "*CRITICAL*")

' Count of Key Deliverables
=COUNTIF(_DeliverableData[EpicKeyDeliverable], "Yes")
```

### Milestones Sheet
The existing `qry_DevOps_Milestones` already pulls Feature dates with `Parent.WorkItemId`. You can add columns to the Milestones sheet showing parent Epic context:

```excel
' Parent Epic's overall rolled-up end date
=IFERROR(
    XLOOKUP([@ParentID], EpicDataTable[EpicID], EpicDataTable[RolledUpEnd], ""),
    ""
)

' Parent Epic's Critical Date
=IFERROR(
    XLOOKUP([@ParentID], EpicDataTable[EpicID], EpicDataTable[CriticalDate], ""),
    ""
)

' Is parent a Key Deliverable?
=IFERROR(
    XLOOKUP([@ParentID], EpicDataTable[EpicID], EpicDataTable[KeyDeliverable], ""),
    ""
)
```

### Suggestions Sheet
Use rolled-up dates to flag deliverables where the Epic timeline is drifting:

```excel
' Flag deliverables where Epic end is >14 days past the Epic's target
=IFERROR(
    IF(XLOOKUP([@DevOpsID], EpicDataTable[EpicID], EpicDataTable[EndDateVariance], 0) > 14,
       "Timeline review recommended", ""),
    ""
)

' Flag deliverables where rolled-up end breaches Critical Date
=IFERROR(
    LET(
        _crit, XLOOKUP([@DevOpsID], EpicDataTable[EpicID], EpicDataTable[CriticalDate]),
        _rolled, XLOOKUP([@DevOpsID], EpicDataTable[EpicID], EpicDataTable[RolledUpEnd]),
        IF(AND(_crit <> "", _rolled <> "", _rolled > _crit),
           "Critical date breach — escalation required", "")
    ),
    ""
)
```

---

## 6. Design Notes

### Why Power Query Aggregation (Not Excel Formulas)?

The aggregation (MIN/MAX/COUNT) happens in Power Query rather than with Excel MINIFS/MAXIFS because:

- **Performance** — Aggregation runs once at refresh time, not on every recalc
- **Clean separation** — `_EpicData` is a flat, pre-computed table; no volatile array formulas scanning `_MilestoneData`
- **Offline resilience** — Once refreshed, the rolled-up dates are static values that work without live connections
- **Consistency** — Follows the existing pattern where Power Query handles transformation and Excel handles presentation

### Epic Dates: Three Layers of Dates

This design now surfaces three distinct date layers per Epic:

| Layer | Fields | DevOps Properties | Purpose |
|-------|--------|-------------------|---------|
| **Original Baseline** | OriginalStartDate, OriginalEndDate | Custom_OriginalStartDate, Custom_OriginalEndDate | What was originally committed |
| **Current Epic** | EpicStartDate, EpicTargetDate | StartDate, TargetDate | What the PM has currently set on the Epic |
| **Rolled-Up from Features** | RolledUpStart, RolledUpEnd | Aggregated from child Feature StartDate/TargetDate | What the Features are actually showing |

The two variance pairs make the gaps visible:
- `OriginalStartVariance` / `OriginalEndVariance` — Has the Epic itself moved from baseline?
- `StartDateVariance` / `EndDateVariance` — Are the Features aligned with the Epic's current dates?

### Critical Date vs Target Date

`Custom_CriticalDate` represents an immovable external deadline (e.g., regulatory, contractual, ministerial). The `CriticalDateAlert` column compares the rolled-up Feature end against this date, giving early warning when actual delivery trajectory threatens a hard deadline — distinct from the softer `TargetDate` which PMs may adjust.

### Data Quality: FeaturesWithDates Column

Not all Features may have dates populated in DevOps. The `FeaturesWithDates` column compared against `FeatureCount` gives an immediate data quality signal. The `EpicDateHealth` formula uses this to flag "⚠ Missing Dates" before it can even assess timeline health.

---

## 7. Maintenance & Testing

### Refresh Order
Power Query evaluates dependencies automatically. Ensure:
1. `qry_ADO_FeatureDateSummary` refreshes first (Connection Only)
2. `qry_ADO_EpicData` refreshes second (depends on step 1)
3. `qry_SPO_Deliverables` can refresh independently

### Test Cases

| # | Test | Expected Result |
|---|------|-----------------|
| 1 | Epic with 3 Features, all with dates | RolledUpStart = MIN, RolledUpEnd = MAX |
| 2 | Epic with 1 Feature having null StartDate | RolledUpStart = MIN of non-null dates |
| 3 | Epic with no linked Features | FeatureCount = 0, all rolled-up dates = null |
| 4 | Epic with all Features in "Done" state | ActiveRolledUpStart/End = null |
| 5 | Feature dates exactly match Epic dates | StartDateVariance = 0, EndDateVariance = 0 |
| 6 | Feature end date 20 days past Epic target | EndDateVariance = 20, Health = "🔴 Slipping" |
| 7 | Deliverable with no DevOpsID | All XLOOKUP columns return "" or 0 |
| 8 | Data refresh with new Feature added | FeatureCount increments, dates recalculate |
| 9 | Epic with OriginalEndDate earlier than TargetDate | OriginalEndVariance = positive (slipped) |
| 10 | Epic with OriginalEndDate = TargetDate (no movement) | OriginalEndVariance = 0 |
| 11 | RolledUpEnd exceeds CriticalDate | CriticalDateAlert = "🔴 CRITICAL DATE AT RISK" |
| 12 | RolledUpEnd within 14 days of CriticalDate | CriticalDateAlert = "🟡 Approaching Critical Date" |
| 13 | RolledUpEnd well before CriticalDate | CriticalDateAlert = "" (blank) |
| 14 | Epic with CriticalDate = null | CriticalDateAlert = "" (blank, no false alarm) |
| 15 | Epic with KeyDeliverable = Yes | EpicKeyDeliverable shows "Yes" on _DeliverableData |
| 16 | Epic with Effort = 100 | EpicEffort shows 100 on _DeliverableData |

### Security Compliance
- ✅ No VBA macros
- ✅ No ActiveX controls
- ✅ Power Query only (approved connection type)
- ✅ OData v2 stable endpoint (not preview for production)
- ✅ .xlsx format maintained
