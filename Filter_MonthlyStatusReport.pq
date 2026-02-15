=LET(
    _year, $B$2,
    _month, $D$2,
    _area, $F$2,
    _targetDate, _MilestoneData[TargetEndDate],
    _originalDate, _MilestoneData[OriginalEndDate_PMP],
    _completedDate, _MilestoneData[LastUpdated],
    _isComplete, _MilestoneData[Status]="Done",
    _targetFY, IF(MONTH(_targetDate)>=7,
        "FY"&RIGHT(YEAR(_targetDate),2)&"-"&RIGHT(YEAR(_targetDate)+1,2),
        "FY"&RIGHT(YEAR(_targetDate)-1,2)&"-"&RIGHT(YEAR(_targetDate),2)),
    _originalFY, IF(MONTH(_originalDate)>=7,
        "FY"&RIGHT(YEAR(_originalDate),2)&"-"&RIGHT(YEAR(_originalDate)+1,2),
        "FY"&RIGHT(YEAR(_originalDate)-1,2)&"-"&RIGHT(YEAR(_originalDate),2)),
    _completedFY, IF(MONTH(_completedDate)>=7,
        "FY"&RIGHT(YEAR(_completedDate),2)&"-"&RIGHT(YEAR(_completedDate)+1,2),
        "FY"&RIGHT(YEAR(_completedDate)-1,2)&"-"&RIGHT(YEAR(_completedDate),2)),
    _yearOK, IF(_year="All", TRUE,
        (_targetFY=_year) + (_originalFY=_year) + (_isComplete * (_completedFY=_year))),
    _monthNum, MATCH(_month, {"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"}, 0),
    _monthOK, IF(_month="All", TRUE,
        (MONTH(_targetDate)=_monthNum) + (MONTH(_originalDate)=_monthNum) + (_isComplete * (MONTH(_completedDate)=_monthNum))),
    _areaOK, IF(_area="All", TRUE, ISNUMBER(SEARCH(_area, _MilestoneData[AreaPath]))),
    _pmpOK, _MilestoneData[IsPMPMilestone]="Yes",
    SORT(
        CHOOSECOLS(
            FILTER(_MilestoneData, _yearOK * _monthOK * _areaOK * _pmpOK, "No milestones match filters"),
            MATCH({"WorkItemId","Title","OriginalEndDate_PMP","TargetEndDate","Status","IsCriticalDate","IsKeyMilestone"},
                  _MilestoneData[#Headers], 0)
        ),
        4, 1
    )
)
