=LET(
    _fy,        $C$4,
    _mth,       $E$4,
    _proj,      $G$4,

    _fyYear,    IF(_fy = "All", 0,
                    VALUE("20" & RIGHT(_fy, 2))),
    _mthIdx,    IF(_mth = "All", 0,
                    MATCH(_mth,
                        {"Jan","Feb","Mar","Apr","May","Jun",
                         "Jul","Aug","Sep","Oct","Nov","Dec"}, 0)),
    _calYr,     IF(OR(_fyYear = 0, _mthIdx = 0), 0,
                    IF(_mthIdx >= 7, _fyYear - 1, _fyYear)),

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

    _closed,     _MilestoneData[ClosedDate],
    _plStart,    _MilestoneData[PlannedStartDate],
    _targEnd,    _MilestoneData[TargetEndDate],
    _oStart,     _MilestoneData[OriginalStartDate_PMP],
    _oEnd,       _MilestoneData[OriginalEndDate_PMP],
    _projCol,    _MilestoneData[Project],
    _isPMP,      _MilestoneData[IsPMPMilestone],

    _hasClosed,  ISNUMBER(_closed),
    _hasPlStart, ISNUMBER(_plStart),
    _hasTargEnd, ISNUMBER(_targEnd),
    _hasOStart,  ISNUMBER(_oStart),
    _hasOEnd,    ISNUMBER(_oEnd),

    _effStart,   IF(_hasPlStart, _plStart,
                     IF(_hasOStart, _oStart, 0)),
    _effEnd,     IF(_hasTargEnd, _targEnd,
                     IF(_hasOEnd, _oEnd, 0)),
    _hasEffStart, ISNUMBER(_effStart) * (_effStart > 0),
    _hasEffEnd,   ISNUMBER(_effEnd) * (_effEnd > 0),
    _hasEff,      _hasEffStart * _hasEffEnd,

    _closedFY,   IF(_hasClosed,
                     IF(MONTH(_closed) >= 7,
                         YEAR(_closed) + 1,
                         YEAR(_closed)),
                     0),
    _isComplete, _hasClosed
                 * IF(_fy = "All",  TRUE, _closedFY = _fyYear)
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

    _cat,        IF(_isComplete, 1,
                     IF(_isScheduled, 2,
                         IF(_isDelayed, 3, 0))),
    _catLabel,   SWITCH(_cat,
                     1, "Complete",
                     2, "Scheduled",
                     3, "Delayed",
                     ""),

    _sortDate,   IF(_cat = 1, _closed,
                     IF(_hasEffEnd, _effEnd,
                         IF(_hasOEnd, _oEnd, DATE(2099, 12, 31)))),

    _projOK,     IF(_proj = "All", TRUE, _projCol = _proj),
    _pmpOK,      _isPMP = "Yes",

    _outCols,    HSTACK(
                     _catLabel,
                     _MilestoneData[MilestoneID],
                     _MilestoneData[Title],
                     _plStart,
                     _targEnd,
                     _MilestoneData[Status],
                     _MilestoneData[IsCriticalDate],
                     _MilestoneData[KeyMilestone]
                 ),

    _match,      (_cat > 0) * _projOK * _pmpOK,

    IFERROR(
        SORTBY(
            FILTER(_outCols, _match),
            FILTER(_cat, _match), 1,
            FILTER(_sortDate, _match), 1
        ),
        "No milestones match filters"
    )
)
