=LET(
    ' ═══════════════════════════════════════════════════════
    ' INPUTS (dropdown cells)
    ' ═══════════════════════════════════════════════════════
    _fy,        $B$2,
    _mth,       $D$2,
    _proj,      $F$2,
    _data,      _MilestoneData,

    ' ═══════════════════════════════════════════════════════
    ' RESOLVE REPORTING PERIOD
    '   Australian FY: FY26 = Jul 2025 → Jun 2026
    '   Months Jul–Dec fall in calendar year FY-1
    '   Months Jan–Jun fall in calendar year FY
    ' ═══════════════════════════════════════════════════════
    _fyYear,    IF(_fy = "All", 0,
                    VALUE("20" & RIGHT(_fy, 2))),
    _mthIdx,    IF(_mth = "All", 0,
                    MATCH(_mth,
                        {"Jan","Feb","Mar","Apr","May","Jun",
                         "Jul","Aug","Sep","Oct","Nov","Dec"}, 0)),
    _calYr,     IF(OR(_fyYear = 0, _mthIdx = 0), 0,
                    IF(_mthIdx >= 7, _fyYear - 1, _fyYear)),

    ' Period boundaries — wide ranges for "All" options
    _pStart,    SWITCH(TRUE,
                    AND(_fy="All", _mth="All"),  DATE(2020, 1, 1),
                    _mth = "All",                DATE(_fyYear - 1, 7, 1),
                    _fy = "All",                 DATE(2020, 1, 1),
                                                 DATE(_calYr, _mthIdx, 1)),
    _pEnd,      SWITCH(TRUE,
                    AND(_fy="All", _mth="All"),  DATE(2035, 12, 31),
                    _mth = "All",                DATE(_fyYear, 6, 30),
                    _fy = "All",                 DATE(2035, 12, 31),
                                                 EOMONTH(DATE(_calYr, _mthIdx, 1), 0)),

    ' ═══════════════════════════════════════════════════════
    ' EXTRACT & VALIDATE COLUMNS
    '   ISNUMBER guards against blanks, text, and errors
    '   in date columns — Excel stores dates as numbers
    ' ═══════════════════════════════════════════════════════
    _closed,     _data[ClosedDate],
    _plStart,    _data[PlannedStartDate],
    _targEnd,    _data[TargetEndDate],
    _oStart,     _data[OriginalStartDate_PMP],
    _oEnd,       _data[OriginalEndDate_PMP],
    _projCol,    _data[Project],

    _hasClosed,  ISNUMBER(_closed),
    _hasPlStart, ISNUMBER(_plStart),
    _hasTargEnd, ISNUMBER(_targEnd),
    _hasOStart,  ISNUMBER(_oStart),
    _hasOEnd,    ISNUMBER(_oEnd),

    ' Effective dates: planned first, fall back to PMP originals
    _effStart,   IF(_hasPlStart, _plStart,
                     IF(_hasOStart, _oStart, 0)),
    _effEnd,     IF(_hasTargEnd, _targEnd,
                     IF(_hasOEnd, _oEnd, 0)),
    _hasEffStart, ISNUMBER(_effStart) * (_effStart > 0),
    _hasEffEnd,   ISNUMBER(_effEnd) * (_effEnd > 0),
    _hasEff,      _hasEffStart * _hasEffEnd,

    ' ═══════════════════════════════════════════════════════
    ' CATEGORY 1 │ COMPLETE
    '   ClosedDate exists AND falls within reporting period
    '   Uses FY/Month independently so "All" works per axis
    ' ═══════════════════════════════════════════════════════
    _closedFY,   IF(_hasClosed,
                     IF(MONTH(_closed) >= 7,
                         YEAR(_closed) + 1,
                         YEAR(_closed)),
                     0),
    _isComplete, _hasClosed
                 * IF(_fy = "All",  TRUE, _closedFY = _fyYear)
                 * IF(_mth = "All", TRUE, MONTH(_closed) = _mthIdx),

    ' ═══════════════════════════════════════════════════════
    ' CATEGORY 2 │ SCHEDULED
    '   Effective date range overlaps the reporting period
    '   AND milestone is NOT already complete in this period
    '
    '   Overlap test: StartA <= EndB AND EndA >= StartB
    ' ═══════════════════════════════════════════════════════
    _isScheduled, NOT(_isComplete)
                  * _hasEff
                  * (_effStart <= _pEnd)
                  * (_effEnd >= _pStart),

    ' ═══════════════════════════════════════════════════════
    ' CATEGORY 3 │ DELAYED
    '   Original PMP dates overlap the period (should be active)
    '   BUT planned start is pushed beyond period end
    '       — or no planned start exists at all (unscheduled)
    '   Excludes anything already Complete or Scheduled
    ' ═══════════════════════════════════════════════════════
    _isDelayed,  NOT(_isComplete)
                 * NOT(_isScheduled)
                 * _hasOStart * _hasOEnd
                 * (_oStart <= _pEnd)
                 * (_oEnd >= _pStart)
                 * IF(_hasPlStart, _plStart > _pEnd, TRUE),

    ' Category assignment
    _cat,        IF(_isComplete, 1,
                     IF(_isScheduled, 2,
                         IF(_isDelayed, 3, 0))),
    _catLabel,   SWITCH(_cat,
                     1, "Complete",
                     2, "Scheduled",
                     3, "Delayed",
                     ""),

    ' Sort key: ClosedDate for complete, effective end for others
    _sortDate,   IF(_cat = 1, _closed,
                     IF(_hasEffEnd, _effEnd,
                         IF(_hasOEnd, _oEnd, DATE(2099, 12, 31)))),

    ' ═══════════════════════════════════════════════════════
    ' PROJECT FILTER
    ' ═══════════════════════════════════════════════════════
    _projOK,     IF(_proj = "All", TRUE, _projCol = _proj),

    ' ═══════════════════════════════════════════════════════
    ' COMBINE → FILTER → SORT → OUTPUT
    '   Sort: Category ASC (Complete → Scheduled → Delayed)
    '         then SortDate ASC within each category
    ' ═══════════════════════════════════════════════════════
    _match,      (_cat > 0) * _projOK,

    IFERROR(
        SORTBY(
            FILTER(HSTACK(_catLabel, _data), _match),
            FILTER(_cat, _match), 1,
            FILTER(_sortDate, _match), 1
        ),
        "No milestones match filters"
    )
)
