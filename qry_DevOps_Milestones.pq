let
    // ===== CONFIGURATION - UPDATE THESE VALUES =====
    Organisation = "YourOrganisation",
    Project = "YourProject",
    
    // Specify your AreaLevel2 values (add or remove as needed)
    AreaLevel2List = {"Programme1", "Programme2", "Programme3"},
    
    // ===== BUILD ODATA URL =====
    BaseUrl = "https://analytics.dev.azure.com/" & Organisation & "/" & Project & "/_odata/v2.0/WorkItems?",
    
    // Build AreaLevel2 filter dynamically
    AreaFilters = Text.Combine(
        List.Transform(AreaLevel2List, each "Area/AreaLevel2 eq '" & _ & "'"),
        " or "
    ),
    
    // Filter: Features only, not removed, matching AreaLevel2 values
    Filter = "$filter=WorkItemType eq 'Feature' and State ne 'Removed' and (" & AreaFilters & ")",
    
    // Select: Core fields and custom fields (using your exact field names)
    Select = "$select=" & 
        "WorkItemId," &
        "Title," &
        "State," &
        "Priority," &
        "ChangedDate," &
        "TargetDate," &
        "Custom_OriginalStartDate," &
        "Custom_OriginalEndDate," &
        "Custom_EndDate," &
        "Custom_KeyFeature," &
        "Custom_CriticalDate," &
        "Custom_RAGStatus," &
        "ParentWorkItemId",
    
    // Expand: Get AssignedTo details and Area details
    Expand = "$expand=AssignedTo($select=UserName),Area($select=AreaLevel2)",
    
    // Order by target date
    OrderBy = "$orderby=TargetDate asc",
    
    // Combine URL
    FullUrl = BaseUrl & Filter & "&" & Select & "&" & Expand & "&" & OrderBy,
    
    // ===== EXECUTE QUERY =====
    Source = OData.Feed(FullUrl),
    
    // ===== TRANSFORM DATA =====
    // Expand AssignedTo to get Assignee name
    #"Expanded AssignedTo" = Table.ExpandRecordColumn(Source, "AssignedTo", {"UserName"}, {"Assignee"}),
    
    // Expand Area to get Programme (AreaLevel2)
    #"Expanded Area" = Table.ExpandRecordColumn(#"Expanded AssignedTo", "Area", {"AreaLevel2"}, {"Programme"}),
    
    // Rename columns for clarity
    #"Renamed Columns" = Table.RenameColumns(#"Expanded Area", {
        {"WorkItemId", "FeatureID"},
        {"State", "Status"},
        {"ChangedDate", "LastUpdated"},
        {"TargetDate", "TargetStart"},
        {"Custom_EndDate", "TargetEnd"},
        {"Custom_OriginalStartDate", "OriginalStartDate"},
        {"Custom_OriginalEndDate", "OriginalEndDate"},
        {"Custom_KeyFeature", "IsKeyFeature"},
        {"Custom_CriticalDate", "IsCriticalDate"},
        {"Custom_RAGStatus", "RAGStatus"},
        {"ParentWorkItemId", "ParentID"}
    }),
    
    // Reorder columns
    #"Reordered Columns" = Table.ReorderColumns(#"Renamed Columns", {
        "FeatureID",
        "Title",
        "Assignee",
        "Programme",
        "Status",
        "Priority",
        "RAGStatus",
        "LastUpdated",
        "OriginalStartDate",
        "OriginalEndDate",
        "TargetStart",
        "TargetEnd",
        "IsKeyFeature",
        "IsCriticalDate",
        "ParentID"
    })
in
    #"Reordered Columns"
