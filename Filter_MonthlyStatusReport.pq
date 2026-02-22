=LET(
    _fy,        $D$3,
    _mth,       $E$3,
    _proj,      $F$3,

    _fyYear,    VALUE("20" & RIGHT(_fy, 2)),

    _mthIdx,    IF(_mth = "All", 0,
                    MATCH(_mth,
                        {"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"}, 0)),

    _calYr,     IF(_mthIdx = 0, 0, IF(_mthIdx >= 7, _fyYear - 1, _fyYear)),

    _pStart,    IF(_mth = "All", DATE(_fyYear - 1, 7, 1), DATE(_calYr, _mthIdx, 1)),
    _pEnd,      IF(_mth = "All", DATE(_fyYear, 6, 30), EOMONTH(DATE(_calYr, _mthIdx, 1), 0)),

    _closed,     _MilestoneData[ClosedDate],
    _plStart,    _MilestoneData[PlannedStartDate],
    _targEnd,    _MilestoneData[TargetEndDate],
    _oStart,     _MilestoneData[OriginalStartDate_PMP],
    _oEnd,       _MilestoneData[OriginalEndDate_PMP],
    _projCol,    _MilestoneData[AreaPath],
    _isPMP,      _MilestoneData[IsPMPMilestone],

    _hasClosed,  ISNUMBER(_closed),
    _hasPlStart, ISNUMBER(_plStart),
    _hasTargEnd, ISNUMBER(_targEnd),
    _hasOStart,  ISNUMBER(_oStart),
    _hasOEnd,    ISNUMBER(_oEnd),

    _effStart,   IF(_hasPlStart, _plStart, IF(_hasOStart, _oStart, 0)),
    _effEnd,     IF(_hasTargEnd, _targEnd, IF(_hasOEnd, _oEnd, 0)),
    _hasEff,     (_effStart > 0) * (_effEnd > 0),

    _closedFY,   IF(_hasClosed,
                     IF(MONTH(_closed) >= 7, YEAR(_closed) + 1, YEAR(_closed)),
                     0),

    _isComplete, _hasClosed
                 * (_closedFY = _fyYear)
                 * IF(_mth = "All", TRUE, MONTH(_closed) = _mthIdx),

    _isScheduled, NOT(_isComplete)
                  * _hasEff
                  * (_effStart <= _pEnd)
                  * (_effEnd >= _pStart),

    _isDelayed,  NOT(_isComplete)
                 * NOT(_isScheduled)
                 * _hasOStart * _hasOEnd
                 * (_oStart <= _pEnd)
                 * (_oEnd >= _pStart)
                 * IF(_hasPlStart, _plStart > _pEnd, TRUE),

    _cat,        IF(_isComplete, 1, IF(_isScheduled, 2, IF(_isDelayed, 3, 0))),
    _catLabel,   SWITCH(_cat, 1, "Complete", 2, "Scheduled", 3, "Delayed", ""),

    _sortDate,   IFERROR(
                     CHOOSE(_cat,
                         _closed,
                         IF(_effEnd > 0, _effEnd, DATE(2099, 12, 31)),
                         IF(_hasOEnd, _oEnd, DATE(2099, 12, 31))
                     ),
                     DATE(2099, 12, 31)),

    _projOK,     IF(_proj = "All", TRUE, ISNUMBER(SEARCH(_proj, _projCol))),
    _pmpOK,      _isPMP = "Yes",

    _outCols,    HSTACK(
                     _MilestoneData[MilestoneID],
                     _MilestoneData[Title],
                     IF(_effStart > 0, _effStart, "Missing Dates"),
                     IF(_effEnd > 0, _effEnd, "Missing Dates"),
                     _catLabel,
                     IF(_MilestoneData[IsCriticalDate]="", "Not set", _MilestoneData[IsCriticalDate]),
                     IF(_MilestoneData[IsKeyFeature]="", "Not set", _MilestoneData[IsKeyFeature])
                 ),

    _match,      (_cat > 0) * _projOK * _pmpOK,

    _noData,     ROWS(_MilestoneData[MilestoneID]) = 0,
    _projLabel,  IF(_proj = "All", "All Projects", _proj),

    IF(_noData,
        "No milestone data available — refresh Power Query",
        IFERROR(
            SORTBY(
                FILTER(_outCols, _match),
                FILTER(_cat, _match), 1,
                FILTER(_sortDate, _match), 1
            ),
            "No milestones found for " & _fy & " · " & _mth & " · " & _projLabel & IF(_proj <> "All", " — try setting Project to All", "")
        )
    )
)
