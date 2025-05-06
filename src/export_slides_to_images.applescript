on run {keynoteFilePath, tempOutputFolderPath}
    -- 确保临时输出文件夹存在
    tell application "System Events"
        if not (exists folder tempOutputFolderPath) then
            do shell script "mkdir -p " & quoted form of tempOutputFolderPath
        else
            -- 如果文件夹已存在，清空其内容以避免旧文件干扰
            do shell script "rm -rf " & quoted form of (tempOutputFolderPath & "/*")
        end if
    end tell

    tell application "Keynote"
        activate
        try
            open POSIX file keynoteFilePath
            set doc to front document
            
            -- 将所有幻灯片导出为JPEG图片到临时文件夹
            -- Keynote 会使用默认的文件名（如：幻灯片 1.jpeg, 幻灯片 2.jpeg）
            export doc to POSIX file tempOutputFolderPath as slide images with properties {image format:JPEG, export style:IndividualSlides, all stages:false, skipped slides:false}
            
            close doc saving no
            return "成功：幻灯片已临时导出到 " & tempOutputFolderPath
        on error errMsg number errorNum
            try
                if (exists front document) then
                    close front document saving no
                end if
            end try
            return "AppleScript 错误 (导出图片): " & errMsg & " (错误号: " & errorNum & ")"
        end try
    end tell
end run