on run argv
    if (count argv) < 3 or ((count argv) mod 2) is 0 then error "usage: output-directory id input-path [id input-path ...]"

    set outputDirectory to item 1 of argv
    set failures to {}
    do shell script "mkdir -p " & quoted form of outputDirectory

    tell application "Microsoft Word"
        launch
        repeat with argumentIndex from 2 to (count argv) by 2
            set outputId to item argumentIndex of argv
            set inputPath to item (argumentIndex + 1) of argv
            set outputPath to outputDirectory & "/" & outputId & ".pdf"
            set inputFile to my posixFile(inputPath)
            set openedDoc to missing value

            try
                with timeout of 120 seconds
                    open inputFile
                    set openedDoc to active document
                    my createEmptyFile(outputPath)
                    save as openedDoc file name outputPath file format format PDF
                    my waitForNonEmptyFile(outputPath)
                    close openedDoc saving no
                end timeout
            on error errorMessage number errorNumber
                if openedDoc is not missing value then
                    try
                        close openedDoc saving no
                    end try
                end if
                set end of failures to outputId & ": " & errorMessage & " (" & errorNumber & ")"
            end try
        end repeat
        quit
    end tell

    if (count failures) > 0 then error my joinLines(failures)
end run

on posixFile(inputPath)
    return inputPath as POSIX file
end posixFile

on createEmptyFile(outputPath)
    do shell script ": > " & quoted form of outputPath
end createEmptyFile

on waitForNonEmptyFile(outputPath)
    repeat 120 times
        if (do shell script "test -s " & quoted form of outputPath & " && echo yes || true") is "yes" then return
        delay 1
    end repeat
    error "Word did not create a non-empty PDF"
end waitForNonEmptyFile

on joinLines(values)
    set AppleScript's text item delimiters to linefeed
    set joined to values as text
    set AppleScript's text item delimiters to ""
    return joined
end joinLines
