# Quarterly Plan Report — Business Rules & Logic

**Resource Capacity Planner** | Cloud & Platforms PMO

| | |
|---|---|
| **Version** | 1.0 |
| **Date** | 22 February 2026 |

---

## 1. Purpose

The Quarterly Plan Report provides a forward-looking view of **PMP milestones** for a given quarter. It is designed to support quarterly planning cycles and project plan submissions, answering two questions:

- What milestones are scheduled for delivery this quarter?
- What milestones were originally planned for this quarter but have been pushed beyond it?

Each project is required to submit a plan outlining the milestones scheduled for the upcoming quarter, including those that were planned but delayed. This report serves as the consolidated view across all projects.

Milestones are sourced from Azure DevOps Features via Power Query and enriched with calculated columns in the `_MilestoneData` table. Only features flagged as PMP milestones (`IsPMPMilestone` = Yes) are included in this report.

---

## 2. Filters

### 2.1 Pre-Filter (Automatic)

Before any user-selectable filters are applied, the report excludes all features where `_MilestoneData.IsPMPMilestone` ≠ Yes. This is a hard filter that cannot be overridden by the user.

Only DevOps features/milestones with `IsPMPMilestone` = Yes appear in the quarterly plan view. Non-PMP features are excluded entirely — they are not evaluated against any category logic and do not appear in any filter combination.

### 2.2 User-Selectable Filters

Three filters control the report. Quarter and Project filters support an "All" option that disables that filter axis. The Year filter defaults to the current Australian Financial Year and does not have an "All" option. Filters combine using **AND** logic — a milestone must satisfy all active filters to appear.

| Filter | Format | Examples |
|---|---|---|
| Year | Australian FY | FY25, FY26, FY27 (defaults to current FY) |
| Quarter | FY quarter | Q1, Q2, Q3, Q4, All (defaults to current quarter) |
| Project | Project name text | "Network Refresh", All |

### Australian Financial Year Quarter Convention

Quarters align to the Australian Financial Year:

| Quarter | Months | Example (FY26) |
|---|---|---|
| Q1 | July – September | Jul 2025 – Sep 2025 |
| Q2 | October – December | Oct 2025 – Dec 2025 |
| Q3 | January – March | Jan 2026 – Mar 2026 |
| Q4 | April – June | Apr 2026 – Jun 2026 |

### Default Behaviour

- **Year** defaults to the current Australian Financial Year based on today's date. Users can manually select a different financial year from the dropdown.
- **Quarter** defaults to the current quarter based on today's date. Users can manually select a different quarter from the dropdown.

Because Year always resolves to a specific FY and Quarter always resolves to a specific three-month period, the reporting period is always a known, concrete date range.

### Reporting Period Derivation

The reporting period is derived from the selected Year and Quarter:

```
Period Start  =  First day of the quarter's first month
Period End    =  Last day of the quarter's last month
```

| Example | Period Start | Period End |
|---|---|---|
| FY26 + Q1 | 1 July 2025 | 30 September 2025 |
| FY26 + Q3 | 1 January 2026 | 31 March 2026 |

When Quarter is "All" and a specific Year is selected, the report shows the full financial year (1 July to 30 June).

---

## 3. Milestone Categories

The classification logic operates on the **pre-filtered dataset** — only milestones where `IsPMPMilestone` = Yes (see [Section 2.1](#21-pre-filter-automatic)).

Every milestone is classified into exactly one of two categories based on a strict evaluation order. Once a milestone matches a category, it is excluded from subsequent checks. This guarantees no duplicates.

**Evaluation order:** Scheduled → Delayed

Milestones that do not match either category are excluded from the report.

The category label is output to the `Status` column in the quarterly plan report (see [Section 6](#6-output-columns)).

### 3.1 Scheduled

**Definition:** The milestone's effective date range overlaps with the reporting period (the selected quarter).

#### Date Fallback Logic

The system uses "effective dates" with a two-tier fallback:

1. **First preference:** `PlannedStartDate` and `TargetEndDate` — the current working dates maintained in DevOps.
2. **Fallback:** `OriginalStartDate_PMP` and `OriginalEndDate_PMP` — the baseline dates from the Project Management Plan.

Start and end dates are resolved independently — e.g. `PlannedStartDate` may be available while `TargetEndDate` is missing, in which case the system uses `PlannedStartDate` for the start and falls back to `OriginalEndDate_PMP` for the end.

If neither tier provides a usable date for both start and end, the milestone cannot be classified as Scheduled and falls through to the Delayed check.

#### Overlap Test

A milestone overlaps the reporting period when **both** of these are true:

```
Effective Start Date  <=  Last day of the quarter
Effective End Date    >=  First day of the quarter
```

This standard range-overlap test correctly captures milestones that:

- Start before and end during the quarter
- Start and end entirely within the quarter
- Start during and end after the quarter
- Span the entire quarter

#### Completed Milestones Within the Quarter

Milestones with a `ClosedDate` that falls within the selected quarter still appear as **Scheduled** provided their effective dates overlap the period. Unlike the Monthly Milestone Report, there is no separate Complete category in the quarterly plan — the focus is on what is planned, not what has been delivered.

> **Note:** If a milestone was completed before the quarter began (i.e. `ClosedDate` is before the period start), it will not appear as Scheduled unless its effective dates still overlap the quarter. Completed milestones whose dates no longer overlap are naturally excluded.

### 3.2 Delayed

**Definition:** The milestone is NOT scheduled in the reporting period, BUT its original PMP baseline dates indicate it SHOULD have been active during this quarter, AND its planned start has been pushed beyond the quarter (or no planned start exists at all).

> *"According to the original plan, this milestone should be happening this quarter. But it hasn't been scheduled yet, or it's been pushed out."*

**Conditions** (all must be true):

1. Not already classified as Scheduled.
2. `OriginalStartDate_PMP` and `OriginalEndDate_PMP` both exist.
3. The original PMP date range overlaps the reporting period (same overlap test as Scheduled).
4. **Either:**
   - (a) `PlannedStartDate` exists but is **after** the quarter end date (milestone has been explicitly pushed out), **or**
   - (b) `PlannedStartDate` does not exist (milestone has never been scheduled).

**What this catches:**

- Milestones where the PM has pushed the start date beyond the current quarter.
- Milestones that were baselined for this quarter but never had working dates assigned.

**What this does NOT catch:**

- Milestones that started on time but are overrunning — these appear as Scheduled because their effective dates still overlap the quarter.
- Overrunning milestones are better identified via the `IsOverdue` calculated column (`TargetEndDate` < `TODAY` and Status ≠ Complete) which is handled separately by conditional formatting on the Milestones sheet.

---

## 4. Sort Order

Results are sorted in two tiers:

- **Primary sort:** Status value order — Scheduled first, then Delayed.
- **Secondary sort:** Date ascending within each category.

The date used for secondary sorting depends on the category:

| Category | Sort Date |
|---|---|
| Scheduled | `Effective End` (when it's due) |
| Delayed | `OriginalEndDate_PMP` (when it was due) |

If no usable sort date exists, the milestone sorts to the bottom of its category group.

---

## 5. Edge Cases & Defensive Handling

### 5.1 Non-PMP Milestones

Features where `IsPMPMilestone` ≠ Yes are excluded from the quarterly plan entirely. They are not evaluated against any category logic and do not appear in any filter combination. This is applied as a hard pre-filter before all other processing (see [Section 2.1](#21-pre-filter-automatic)).

### 5.2 Missing Dates

- If all four date fields (`PlannedStartDate`, `TargetEndDate`, `OriginalStartDate_PMP`, `OriginalEndDate_PMP`) are blank, the milestone is excluded from the report entirely.
- If planned dates are blank but PMP originals exist, the milestone can still qualify as Scheduled (via fallback) or Delayed.

### 5.3 Invalid Date Values

All date columns are validated using `ISNUMBER()` before comparison. Text values, errors, or blank cells are treated as "no date available" and trigger the fallback cascade.

### 5.4 No Matches

- If the `_MilestoneData` table contains no rows (e.g. Power Query has not been refreshed), the report displays:
  > `No milestone data available — refresh Power Query`
- If milestones exist but none match the combined filters, the report displays the active filter values in the message:
  > `No milestones found for FY26 · Q3 · All Projects`
- When a specific Project filter is active and returns no results, the message includes a hint:
  > `No milestones found for FY26 · Q3 · Network Refresh — try setting Project to All`

### 5.5 Milestone Closed and Reopened

DevOps updates `ClosedDate` to reflect the most recent closure. If a milestone is reopened, `ClosedDate` is cleared and the milestone is re-evaluated as Scheduled or Delayed based on its dates.

### 5.6 Quarter Boundary Milestones

Milestones that span multiple quarters will appear in each quarter where their effective dates overlap. For example, a milestone running from August to November (FY26) will appear as Scheduled in both Q1 and Q2. This is correct behaviour — it reflects that the milestone is active work in both periods.

---

## 6. Output Columns

The quarterly plan report displays a simplified subset of `_MilestoneData` fields, tailored for quarterly planning and project plan submissions.

| Column | Source | Description |
|---|---|---|
| DevOpsID | `WorkItemId` + `DevOpsLink` | Numeric ID (e.g. 45367) displayed as a clickable hyperlink to the DevOps work item |
| Title | `Title` | Milestone name |
| Effective Start | `PlannedStartDate`, falling back to `OriginalStartDate_PMP` | Resolved start date after fallback |
| Effective End | `TargetEndDate`, falling back to `OriginalEndDate_PMP` | Resolved end date after fallback |
| Status | Calculated category label | "Scheduled" or "Delayed" |
| Critical Milestone | `IsCriticalDate` | Yes/No flag |
| Key Milestone | `IsKeyFeature` | Yes/No flag |

> **Note:** The Status column does not show the DevOps state (New, Active, Resolved, Closed). It is populated with the calculated category label from the milestone classification logic described in [Section 3](#3-milestone-categories).

---

## 7. Relationship to Other Views

This report complements the Monthly Milestone Report and the Milestones sheet. The key differences:

| Aspect | Milestones Sheet | Monthly Report | Quarterly Plan |
|---|---|---|---|
| Scope | All milestones | PMP milestones only | PMP milestones only |
| Period | N/A (current state) | Single month | Full quarter (3 months) |
| Categories | N/A | Complete, Scheduled, Delayed | Scheduled, Delayed |
| Filtering | By individual field values (PM, Area, Status, Month, Quarter) | By reporting period (FY + Month) with category logic | By reporting period (FY + Quarter) with category logic |
| View | Current state of all milestones | What happened / should have happened in a specific month | What is planned / should have been planned for a quarter |
| Sort | By Target End Date | By Status then Date | By Status then Date |
| Use case | Day-to-day milestone tracking | Monthly resourcing meetings | Quarterly planning cycles and project plan submissions |
