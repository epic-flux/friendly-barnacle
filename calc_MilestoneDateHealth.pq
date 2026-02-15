=LET(
    _state, [@State],
    _targetEnd, [@TargetEnd],
    _lastUpdated, [@LastUpdated],
    _daysUntilDue, IF(_targetEnd = "", "", _targetEnd - TODAY()),
    _daysSinceUpdate, IF(_lastUpdated = "", "", TODAY() - _lastUpdated),
    _staleThreshold, Config!$B$10,

    IF(OR(_state = "Done", _state = "Closed"), "✅ Complete",
    IF(_targetEnd = "", "⚠ Missing Dates",
    IF(_targetEnd < TODAY(), "🔴 Overdue",
    IF(AND(_daysSinceUpdate <> "", _daysSinceUpdate > _staleThreshold), "⚪ Stale",
    IF(_daysUntilDue <= 7, "🟠 Due Soon",
    IF(_daysUntilDue <= 30, "🟡 Approaching",
    "🟢 On Track"))))))
)