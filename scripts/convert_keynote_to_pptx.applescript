on run argv
    if (count of argv) < 2 then
        return "错误：请提供输入 Keynote 文件路径和输出 PPTX 文件路径作为参数。"
    end if

    set inputKeynotePath to item 1 of argv
    set outputPptxPath to item 2 of argv

    try
        tell application "Keynote"
            activate
            set keynoteDocument to open POSIX file inputKeynotePath
            export keynoteDocument to POSIX file outputPptxPath as Microsoft PowerPoint
            close keynoteDocument saving no
        end tell
        return "成功将 '" & inputKeynotePath & "' 转换为 '" & outputPptxPath & "'"
    on error errMsg number errorNumber
        return "AppleScript 错误: " & errMsg & " (错误码: " & errorNumber & ")"
    end try
end run
