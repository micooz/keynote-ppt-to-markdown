#!/usr/bin/env node
/**
 * PR Comment Handler for Agent
 * 自动监听 PR 评论，解析命令并执行，然后回复
 */

const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const REPO = 'micooz/keynote-ppt-to-markdown';
const PR_NUMBER = 1;
const TOKEN = process.env.GITHUB_TOKEN || execSync('echo $GITHUB_TOKEN').toString().trim();
const SCRIPT_DIR = __dirname;

async function apiRequest(url, options = {}) {
    const response = await fetch(url, {
        ...options,
        headers: {
            'Authorization': `token ${TOKEN}`,
            'Accept': 'application/vnd.github.v3+json',
            'Content-Type': 'application/json',
            ...options.headers
        }
    });
    return response.json();
}

async function getComments() {
    const data = await apiRequest(
        `https://api.github.com/repos/${REPO}/issues/${PR_NUMBER}/comments?per_page=10`
    );
    return data.filter(c => c.user.login === 'micooz');
}

async function replyComment(commentId, body) {
    await apiRequest(
        `https://api.github.com/repos/${REPO}/issues/comments/${commentId}/replies`,
        {
            method: 'POST',
            body: JSON.stringify({ body })
        }
    );
}

async function executeCommand(cmd, commentId) {
    console.log(`执行命令: ${cmd}`);
    
    switch(cmd.trim()) {
        case '/run':
        case '/convert':
            await replyComment(commentId, '✅ 收到命令，正在执行安装和构建...');
            
            // 执行构建
            execSync('npm install --silent', { cwd: SCRIPT_DIR, stdio: 'pipe' });
            execSync('npm run build', { cwd: SCRIPT_DIR, stdio: 'pipe' });
            
            await replyComment(commentId, '✅ 依赖安装完成，构建成功！\n脚本已就绪，可供 Agent 使用。');
            break;
            
        case '/build':
            execSync('npm run build', { cwd: SCRIPT_DIR, stdio: 'pipe' });
            await replyComment(commentId, '✅ 构建完成！');
            break;
            
        case '/status':
            const branch = execSync('git branch --show-current', { cwd: SCRIPT_DIR }).toString().trim();
            const lastCommit = execSync('git log -1 --oneline', { cwd: SCRIPT_DIR }).toString().trim();
            await replyComment(commentId, `📊 当前状态：
- 分支：${branch}
- 最后提交：${lastCommit}
- 状态：就绪`);
            break;
            
        case '/help':
            await replyComment(commentId, `🤖 可用命令：
- /run - 安装依赖并构建
- /build - 仅构建项目  
- /status - 查看当前状态
- /help - 显示此帮助

Agent 可直接调用 scripts/convert.sh 进行转换。`);
            break;
            
        default:
            await replyComment(commentId, `❓ 未知命令：${cmd}\n可用命令：/run, /build, /status, /help`);
    }
}

async function main() {
    console.log('监听 PR 评论...');
    
    const comments = await getComments();
    
    for (const comment of comments) {
        const firstLine = comment.body.split('\n')[0].trim();
        
        if (firstLine.startsWith('/')) {
            console.log(`发现命令: ${firstLine} (comment #${comment.id})`);
            await executeCommand(firstLine, comment.id);
        }
    }
}

main().catch(console.error);
