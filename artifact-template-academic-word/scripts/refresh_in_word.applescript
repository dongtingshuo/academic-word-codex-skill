on run argv
    if (count of argv) is not 1 then error "usage: refresh_in_word.applescript input.docx"
    set inputPath to item 1 of argv
    with timeout of 45 seconds
        tell application "Microsoft Word"
            set openedDocument to open (POSIX file inputPath)
            save openedDocument
            close openedDocument saving no
        end tell
    end timeout
    return inputPath
end run
