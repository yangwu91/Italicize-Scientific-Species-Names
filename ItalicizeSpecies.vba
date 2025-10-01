Sub ItalicizeSpeciesSpecies()
    ' =================================================================
    '
    ' NAME: Italicize Scientific Species Species
    '
    ' AUTHOR: Yang WU
    '
    ' DESCRIPTION:
    ' This macro automatically finds and italicizes scientific names
    ' (species, genera, and abbreviations) within a Word document.
    ' It is designed to be robust, intelligent, and safe to run
    ' multiple times without causing errors.
    '
    ' FEATURES:
    ' 1.  Handles full species names (e.g., "Aedes albopictus").
    ' 2.  Automatically extracts and processes standalone genus names
    '     (e.g., "Aedes") from the full species list.
    ' 3.  Handles abbreviated names (e.g., "Ae. albopictus") based on
    '     a user-defined list.
    ' 4.  Prevents re-processing of already italicized text, making it
    '     safe to run in documents with "Track Changes" enabled.
    ' 5.  Searches the entire document, regardless of the current cursor
    '     position or selection.
    ' 6.  Reports the total number of modifications made upon completion.
    '
    ' =================================================================

    ' --- PART 1: USER CONFIGURATION - FULL SPECIES LIST ---
    ' Add any full scientific names (species, viruses, etc.) to this list.
    ' The macro will automatically derive the standalone genus from these names.
	' 这里自定义需要斜体的物种名全称：
    ' =================================================================
    Dim speciesList As Variant
    speciesList = Array( _
        "Aedes albopictus", "Stegomyia albopicta", "Aedes aegypti", _
        "Culex pipiens", "Culex quinquefasciatus", "Culex tarsalis", _
        "Anopheles sinensis", "Anopheles gambiae", "Anopheles stephensi", _
        "Drosophila melanogaster" _
        )

    ' --- PART 2: USER CONFIGURATION - ABBREVIATION LIST ---
    ' Add any genus abbreviations you want to handle.
    ' IMPORTANT: The abbreviation must include the period (e.g., "Ae.").
    ' This list is case-sensitive.
	' 这里自定义需要斜体的属名简写：
    ' =================================================================
    Dim abbreviationsList As Variant
    abbreviationsList = Array( _
        "Ae.", "Cx.", "An.", _
        "D." _
        )
        
    ' --- PART 3: MACRO EXECUTION ---
    ' No modifications are needed below this line.
    ' =================================================================
    Dim species As Variant
    Dim abbr As Variant
    Dim genus As Variant
    Dim modificationCount As Long
    Dim finalMessage As String
    Dim searchRange As Range
    
    ' Use a Dictionary object to store and de-duplicate the extracted genus names.
    Dim genusDict As Object
    Set genusDict = CreateObject("Scripting.Dictionary")
    
    ' Initialize the counter for total changes made.
    modificationCount = 0
    
    ' Turn off screen updating to speed up the macro.
    Application.ScreenUpdating = False
    
    ' === TASK 1: Process full species names and extract genera ===
    For Each species In speciesList
        ' Set the search scope to the entire document content.
        Set searchRange = ActiveDocument.Content
        With searchRange.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = species
            .Font.Italic = False ' CRITICAL: Only find non-italicized text to prevent re-processing.
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = False
            .MatchWildcards = False
            .MatchWholeWord = False ' A species name is a phrase, not a single "whole word".
            
            ' Loop through all found instances to count them.
            Do While .Execute
                modificationCount = modificationCount + 1
                searchRange.Font.Italic = True ' Apply italic formatting.
                searchRange.Collapse wdCollapseEnd ' Move to the end of the found range to continue searching.
            Loop
        End With
        
        ' Extract the genus (the first word) from the species name.
        genus = Split(species, " ")(0)
        ' Add the genus to the dictionary, which automatically handles duplicates.
        If Not genusDict.Exists(genus) Then
            genusDict.Add genus, True
        End If
    Next species

    ' === TASK 2: Process the auto-extracted standalone genus names ===
    For Each genus In genusDict.Keys
        Set searchRange = ActiveDocument.Content
        With searchRange.Find
            .ClearFormatting
            .Text = genus
            .Font.Italic = False ' CRITICAL: Only find non-italicized text.
            .MatchWholeWord = True ' CRITICAL: Must be a whole word to not match inside another word.
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = True ' Genus names are proper nouns and should be case-sensitive.
            .MatchWildcards = False

            Do While .Execute
                modificationCount = modificationCount + 1
                searchRange.Font.Italic = True
                searchRange.Collapse wdCollapseEnd
            Loop
        End With
    Next genus

    ' === TASK 3: Process abbreviated names using wildcards ===
    For Each abbr In abbreviationsList
        Set searchRange = ActiveDocument.Content
        With searchRange.Find
            .ClearFormatting
            ' Build the wildcard pattern, e.g., "(Ae. [a-z]{1,})"
            .Text = "(" & abbr & " [a-z]{1,})"
            .Font.Italic = False ' CRITICAL: Only find non-italicized text.
            .MatchWildcards = True ' IMPORTANT: Enable wildcard searching.
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = True ' Abbreviations are typically case-sensitive.
            
            Do While .Execute
                modificationCount = modificationCount + 1
                searchRange.Font.Italic = True
                searchRange.Collapse wdCollapseEnd
            Loop
        End With
    Next abbr

    ' Restore screen updating.
    Application.ScreenUpdating = True
    
    ' Prepare the final report message.
    If modificationCount > 0 Then
        finalMessage = "Operation complete! A total of " & modificationCount & " modifications were made."
    Else
        finalMessage = "Operation complete! No items were found that needed modification."
    End If

    ' Display the final report to the user.
    MsgBox finalMessage, vbInformation, "Macro Finished"
End Sub