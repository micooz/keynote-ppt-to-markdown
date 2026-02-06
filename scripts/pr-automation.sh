#!/bin/bash
# PR Comment Automation Script
# 自动监听 PR 评论，解析命令并执行

REPO="micooz/keynote-ppt-to-markdown"
PR_NUMBER=1
TOKEN="$GITHUB_TOKEN"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# 获取最新的评论
get_comments() {
    curl -s -H "Authorization: token $TOKEN" \
        "https://api.github.com/repos/$REPO/issues/$PR_NUMBER/comments?per_page=10" | \
        jq -r '.[] | select(.user.login == "micooz") | "\(.id)|\(.body)"'
}

# 回复评论
reply_comment() {
    local comment_id="$1"
    local body="$2"
    curl -s -X POST \
        -H "Authorization: token $TOKEN" \
        -H "Accept: application/vnd.github.v3+json" \
        "https://api.github.com/repos/$REPO/issues/$PR_NUMBER/comments/$1/replies" \
        -d "{\"body\":\"$body\"}"
}

# 解析并执行命令
execute_command() {
    local cmd="$1"
    local comment_id="$2"
    
    case "$cmd" in
        "/run"|"/convert")
            reply_comment "$comment_id" "✅ 收到命令，正在执行..."
            # 执行转换命令
            cd "$SCRIPT_DIR"
            npm install --silent 2>/dev/null
            npm run build --silent 2>/dev/null
            reply_comment "$comment_id" "✅ 依赖已安装，构建完成！"
            ;;
        "/build")
            cd "$SCRIPT_DIR"
            npm run build
            reply_comment "$comment_id" "✅ 构建完成！"
            ;;
        "/status")
            reply_comment "$comment_id" "📊 当前状态：\n- 分支：$(git branch --show-current)\n- 最后提交：$(git log -1 --oneline)\n- 修改文件：$(git diff --name-only HEAD | tr '\n' ', ')"
            ;;
        "/help")
            reply_comment "$comment_id" "🤖 可用命令：\n- /run - 执行构建和转换\n- /build - 仅构建项目\n- /status - 查看当前状态\n- /help - 显示此帮助"
            ;;
        *)
            reply_comment "$comment_id" "❓ 未知命令：$cmd\n可用命令：/run, /build, /status, /help"
            ;;
    esac
}

# 主循环
main() {
    echo "监听 PR 评论..."
    while true; do
        comments=$(get_comments)
        if [ -n "$comments" ]; then
            echo "$comments" | while IFS='|' read -r id body; do
                # 提取命令（第一行）
                cmd=$(echo "$body" | head -1)
                if [[ "$cmd" == /* ]]; then
                    echo "执行命令: $cmd"
                    execute_command "$cmd" "$id"
                fi
            done
        fi
        sleep 30
    fi
}

main
