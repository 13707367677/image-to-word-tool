#!/bin/bash
# 添加任务到队列的脚本
# 用法: ./add-task.sh "任务描述"

QUEUE_FILE="C:\Users\Administrator\lobsterai\project\task-queue.json"
LOG_FILE="C:\Users\Administrator\lobsterai\project\task-queue.log"

if [ -z "$1" ]; then
    echo "用法: $0 <任务描述>"
    exit 1
fi

TASK_DESC="$1"

# 添加任务到队列
node -e "
const fs=require('fs');
const q=JSON.parse(fs.readFileSync('$QUEUE_FILE', 'utf8'));
q.queue.push({
    id: Date.now(),
    task: '$TASK_DESC',
    addedAt: new Date().toISOString(),
    status: 'pending'
});
fs.writeFileSync('$QUEUE_FILE', JSON.stringify(q, null, 2));
console.log('任务已添加到队列，当前队列长度: ' + q.queue.length);
"

echo "任务 '$TASK_DESC' 已添加到队列"
