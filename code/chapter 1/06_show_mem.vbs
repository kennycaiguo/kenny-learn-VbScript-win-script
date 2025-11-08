strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfOS_Memory",,48)
strService = ""

For Each objItem in colItems
    strService = strService &  "AvailableBytes: " & objItem.AvailableBytes & vbCrLf
    strService = strService &  "AvailableKBytes: " & objItem.AvailableKBytes  & vbCrLf
    strService = strService &  "AvailableMBytes: " & objItem.AvailableMBytes  & vbCrLf
    strService = strService &  "CacheBytes: " & objItem.CacheBytes  & vbCrLf
    strService = strService &  "CacheBytesPeak: " & objItem.CacheBytesPeak  & vbCrLf
    strService = strService &  "CacheFaultsPersec: " & objItem.CacheFaultsPersec  & vbCrLf
    strService = strService &  "Caption: " & objItem.Caption  & vbCrLf
    strService = strService &  "CommitLimit: " & objItem.CommitLimit  & vbCrLf
    strService = strService &  "CommittedBytes: " & objItem.CommittedBytes  & vbCrLf
    strService = strService &  "DemandZeroFaultsPersec: " & objItem.DemandZeroFaultsPersec  & vbCrLf
    strService = strService &  "Description: " & objItem.Description  & vbCrLf
    strService = strService &  "FreeSystemPageTableEntries: " & objItem.FreeSystemPageTableEntries  & vbCrLf
    strService = strService &  "Frequency_Object: " & objItem.Frequency_Object  & vbCrLf
    strService = strService &  "Frequency_PerfTime: " & objItem.Frequency_PerfTime  & vbCrLf
    strService = strService &  "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS  & vbCrLf
    strService = strService &  "Name: " & objItem.Name  & vbCrLf
    strService = strService &  "PageFaultsPersec: " & objItem.PageFaultsPersec  & vbCrLf
    strService = strService &  "PageReadsPersec: " & objItem.PageReadsPersec  & vbCrLf
    strService = strService &  "PagesInputPersec: " & objItem.PagesInputPersec  & vbCrLf
    strService = strService &  "PagesOutputPersec: " & objItem.PagesOutputPersec  & vbCrLf
    strService = strService &  "PagesPersec: " & objItem.PagesPersec  & vbCrLf 
    strService = strService &  "PageWritesPersec: " & objItem.PageWritesPersec  & vbCrLf
    strService = strService &  "PercentCommittedBytesInUse: " & objItem.PercentCommittedBytesInUse  & vbCrLf
    strService = strService &  "PoolNonpagedAllocs: " & objItem.PoolNonpagedAllocs  & vbCrLf
    strService = strService &  "PoolNonpagedBytes: " & objItem.PoolNonpagedBytes  & vbCrLf
    strService = strService &  "PoolPagedAllocs: " & objItem.PoolPagedAllocs  & vbCrLf
    strService = strService &  "PoolPagedBytes: " & objItem.PoolPagedBytes  & vbCrLf
    strService = strService &  "PoolPagedResidentBytes: " & objItem.PoolPagedResidentBytes  & vbCrLf
    strService = strService &  "SystemCacheResidentBytes: " & objItem.SystemCacheResidentBytes  & vbCrLf
    strService = strService &  "SystemCodeResidentBytes: " & objItem.SystemCodeResidentBytes  & vbCrLf
    strService = strService &  "SystemCodeTotalBytes: " & objItem.SystemCodeTotalBytes  & vbCrLf
    strService = strService &  "SystemDriverResidentBytes: " & objItem.SystemDriverResidentBytes  & vbCrLf
    strService = strService &  "SystemDriverTotalBytes: " & objItem.SystemDriverTotalBytes  & vbCrLf
    strService = strService &  "Timestamp_Object: " & objItem.Timestamp_Object  & vbCrLf
    strService = strService &  "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime  & vbCrLf
    strService = strService &  "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS  & vbCrLf
    strService = strService &  "TransitionFaultsPersec: " & objItem.TransitionFaultsPersec  & vbCrLf
    strService = strService &  "TransitionPagesRePurposedPersec: " & objItem.TransitionPagesRePurposedPersec  & vbCrLf
    strService = strService &  "WriteCopiesPersec: " & objItem.WriteCopiesPersec  & vbCrLf
Next

WScript.Echo strService

