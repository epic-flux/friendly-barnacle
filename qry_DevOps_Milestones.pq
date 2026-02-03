let
    // ===== READ CONFIGURATION FROM EXCEL =====
    ConfigTable = Excel.CurrentWorkbook(){[Name="ConfigTable"]}[Content],
    GetConfig = (param) => Table.SelectRows(ConfigTable, each [Parameter] = param){0}[Value],
    
    Organisation = GetConfig("Organisation"),
    Project = GetConfig("Project"),
    
    // ===== READ AREA FILTERS FROM EXCEL =====
    AreaFilterTable = Excel.CurrentWorkbook(){[Name="AreaFilterTable"]}[Content],
    AreaNames = Table.Column(AreaFilterTable, "AreaName"),
    
    // ===== BUILD DYNAMIC AREA FILTER =====
    AreaPathList = List.Transform(AreaNames, each Project & "\" & _),
    AreaFilters = Text.Combine(
        List.Transform(AreaPathList, each "startswith(Area/AreaPath, '" & _ & "')"),
        " or "
    ),
    
    // ===== BUILD ODATA URL =====
    BaseUrl = "https://analytics.dev.azure.com/" & Organisation & "/" & Project & "/_odata/v2.0/WorkItems?",
    Filter = "$filter=WorkItemType eq 'Feature' and State ne 'Removed' and (" & AreaFilters & ")",
    Expand = "$expand=AssignedTo($select=UserName),Area($select=AreaPath)",
    OrderBy = "$orderby=TargetDate asc",
    FullUrl = BaseUrl & Filter & "&" & Expand & "&" & OrderBy,
    
    // ===== EXECUTE QUERY =====
    Source = OData.Feed(FullUrl),
    
    // ===== TRANSFORM DATA =====
    #"Expanded AssignedTo" = Table.ExpandRecordColumn(Source, "AssignedTo", {"UserName"}, {"Assignee"}),
    #"Expanded Area" = Table.ExpandRecordColumn(#"Expanded AssignedTo", "Area", {"AreaPath"}, {"AreaPath"}),
    
    // Rename columns to friendly names
    #"Renamed Columns" = Table.RenameColumns(#"Expanded Area", {
        {"WorkItemId", "FeatureID"},
        {"State", "Status"},
        {"ChangedDate", "LastUpdated"},
        {"StartDate", "ForecastActualStart"},
        {"TargetDate", "ForecastEnd"},
        {"Custom_OriginalStartDate", "OriginalStartDate"},
        {"Custom_OriginalEndDate", "OriginalEndDate"},
        {"Custom_KeyFeature", "IsKeyFeature"},
        {"Custom_CriticalDate", "IsCriticalDate"},
        {"ParentWorkItemId", "EpicID"}
    }),
    
    // Select only the columns we need
    #"Selected Columns" = Table.SelectColumns(#"Renamed Columns", {
        "FeatureID",
        "Title",
        "Assignee",
        "AreaPath",
        "Status",
        "LastUpdated",
        "OriginalStartDate",
        "OriginalEndDate",
        "ForecastActualStart",
        "ForecastEnd",
        "IsKeyFeature",
        "IsCriticalDate",
        "EpicID"
    })
in
    #"Selected Columns"
