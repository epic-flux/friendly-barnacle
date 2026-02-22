=LET(
    _fy,        $D$3,
    _qtr,       $E$3,
    _proj,      $F$3,

    _fyYear,    VALUE("20" & RIGHT(_fy, 2)),

    _qIdx,      IF(_qtr = "All", 0,
                    MATCH(_qtr, {"Q1","Q2","Q3","Q4"}, 0)),

    _pStart,    IF(_qtr = "All", DATE(_fyYear - 1, 7, 1),
                    CHOOSE(_qIdx,
                        DATE(_fyYear - 1, 7, 1),
                        DATE(_fyYear - 1, 10, 1),
                        DATE(_fyYear, 1, 1),
                        DATE(_fyYear, 4, 1))),
    _pEnd,      IF(_qtr = "All", DATE(_fyYear, 6, 30),
                    CHOOSE(_qIdx,
                        DATE(_fyYear - 1, 9, 30),
                        DATE(_fyYear - 1, 12, 31),
                        DATE(_fyYear, 3, 31),
                        DATE(_fyYear, 6, 30))),

    _plStart,    _MilestoneData[PlannedStartDate],
    _targEnd,    _MilestoneData[TargetEndDate],
    _oStart,     _MilestoneData[OriginalStartDate_PMP],
    _oEnd,       _MilestoneData[OriginalEndDate_PMP],
    _projCol,    _MilestoneData[AreaPath],
    _isPMP,      _MilestoneData[IsPMPMilestone],

    _hasPlStart, ISNUMBER(_plStart),
    _hasTargEnd, ISNUMBER(_targEnd),
    _hasOStart,  ISNUMBER(_oStart),
    _hasOEnd,    ISNUMBER(_oEnd),

    _effStart,   IF(_hasPlStart, _plStart, IF(_hasOStart, _oStart, 0)),
    _effEnd,     IF(_hasTargEnd, _targEnd, IF(_hasOEnd, _oEnd, 0)),
    _hasEff,     (_effStart > 0) * (_effEnd > 0),

    _isScheduled, _hasEff
                  * (_effStart <= _pEnd)
                  * (_effEnd >= _pStart),

    _isDelayed,  NOT(_isScheduled)
                 * _hasOStart * _hasOEnd
                 * (_oStart <= _pEnd)
                 * (_oEnd >= _pStart)
                 * IF(_hasPlStart, _plStart > _pEnd, TRUE),

    _cat,        IF(_isScheduled, 1, IF(_isDelayed, 2, 0)),
    _catLabel,   SWITCH(_cat, 1, "Scheduled", 2, "Delayed", ""),

    _sortDate,   IFERROR(
                     CHOOSE(_cat,
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
            "No milestones found for " & _fy & " · " & _qtr & " · " & _projLabel & IF(_proj <> "All", " — try setting Project to All", "")
        )
    )
)
