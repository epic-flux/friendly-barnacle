=IF(B7="","",IF(ISNUMBER(B7),IFERROR(HYPERLINK(INDEX(_MilestoneData[DevOpsLink],MATCH(B7,_MilestoneData[MilestoneID],0)),B7),B7),B7))
