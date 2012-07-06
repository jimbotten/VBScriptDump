Set objSearcher = CreateObject("Microsoft.Update.Searcher")
Set objResults = objSearcher.Search("Type='Software'")
Set colUpdates = objResults.Updates

For i = 0 to colUpdates.Count - 1
    Wscript.Echo "Title: " & colUpdates.Item(i).Title
    Wscript.Echo "Autoselect on Web sites: " & _
        colUpdates.Item(i).AutoSelectOnWebSites

    For Each strUpdate in colUpdates.Item(i).BundledUpdates
        Wscript.Echo "Bundled update: " & strUpdate
    Next
    Wscript.Echo "Can require source: " & colUpdates.Item(i).CanRequireSource
    Set objCategories = colUpdates.Item(i).Categories

    For z = 0 to objCategories.Count - 1
        Wscript.Echo "Category name: " & objCategories.Item(z).Name
        Wscript.Echo "Category ID: " & objCategories.Item(z).CategoryID
        For Each strChild in objCategories.Item(z).Children
            Wscript.Echo "Child category: " & strChild
        Next
        Wscript.Echo "Category description: " & _
            objCategories.Item(z).Description
        Wscript.Echo "Category order: " & objCategories.Item(z).Order
        Wscript.Echo "Category type: " & objCategories.Item(z).Type
    Next

    Wscript.Echo "Deadline: " & colUpdates.Item(i).Deadline
    Wscript.Echo "Delta compressed content available: " & _
        colUpdates.Item(i).DeltaCompressedContentAvailable
    Wscript.Echo "Delta compressed content preferred: " & _
        colUpdates.Item(i).DeltaCompressedContentPreferred
    Wscript.Echo "Description: " & colUpdates.Item(i).Description
    Wscript.Echo "EULA accepted: " & colUpdates.Item(i).EULAAccepted
    Wscript.Echo "EULA text: " & colUpdates.Item(i).EULAText
    Wscript.Echo "Handler ID: " & colUpdates.Item(i).HandlerID

    Set objIdentity = colUpdates.Item(i).Identity
    Wscript.Echo "Revision number: " & objIdentity.RevisionNumber
    Wscript.Echo "Update ID: " & objIdentity.UpdateID

    Set objInstallationBehavior = colUpdates.Item(i).InstallationBehavior
    Wscript.Echo "Can request user input: " & _
        objInstallationBehavior.CanRequestUserInput
    Wscript.Echo "Impact: " & objInstallationBehavior.Impact
    Wscript.Echo "Reboot behavior: " & objInstallationBehavior.RebootBehavior
    Wscript.Echo "Requires network connectivity: " & _
        objInstallationBehavior.RequiresNetworkConnectivity
    Wscript.Echo "Is beta: " & colUpdates.Item(i).IsBeta
    Wscript.Echo "Is hidden: " & colUpdates.Item(i).IsHidden
    Wscript.Echo "Is installed: " & colUpdates.Item(i).IsInstalled
    Wscript.Echo "Is mandatory: " & colUpdates.Item(i).IsMandatory
    Wscript.Echo "Is uninstallable: " & colUpdates.Item(i).IsUninstallable

    For Each strLanguage in colUpdates.Item(i).Languages
        Wscript.Echo "Supported language: " & strLanguage
    Next

    Wscript.Echo "Last deployment change time: " & _
        colUpdates.Item(i).LastDeploymentChangeTime
    Wscript.Echo "Maximum download size: " & colUpdates.Item(i).MaxDownloadSize
    Wscript.Echo "Minimum download size: " & colUpdates.Item(i).MinDownloadSize
    Wscript.Echo "Microsoft Security Response Center severity: " & _
        colUpdates.Item(i).MsrcSeverity
    Wscript.Echo "Recommended CPU speed: " & _
        colUpdates.Item(i).RecommendedCPUSpeed
    Wscript.Echo "Recommended hard disk space: " & _
        colUpdates.Item(i).RecommendedHardDiskSpace
    Wscript.Echo "Recommended memory: " & colUpdates.Item(i).RecommendedMemory
    Wscript.Echo "Release notes: " & colUpdates.Item(i).ReleaseNotes
    Wscript.Echo "Support URL: " & colUpdates.Item(i).SupportURL
    Wscript.Echo "Type: " & colUpdates.Item(i).Type
    Wscript.Echo "Uninstallation notes: " & _
        colUpdates.Item(i).UninstallationNotes

    x = 1
    For Each strStep in colUpdates.Item(i).UninstallationSteps
        Wscript.Echo x & " -- " & strStep
        x = x + 1
    Next

    For Each strArticle in colUpdates.Item(i).KBArticleIDs
        Wscript.Echo "KB article: " & strArticle
    Next

    Wscript.Echo "Deployment action: " & colUpdates.Item(i).DeploymentAction
    Wscript.Echo
Next