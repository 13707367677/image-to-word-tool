#!/bin/bash
# 自动任务调度器 - 每2小时运行一次

QUEUE_FILE="C:\Users\Administrator\lobsterai\project\task-queue.json"
LOG_FILE="C:\Users\Administrator\lobsterai\project\task-queue.log"

log() {
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] $1" >> "$LOG_FILE"
}

# 读取当前小时
CURRENT_HOUR=$(date +%H)
log "检查调度... 当前时间: $CURRENT_HOUR"

# 检查是否在允许的时间段 (21:00-09:00)
ALLOWED_HOURS=("21" "22" "23" "00" "01" "02" "03" "04" "05" "06" "07" "08")

IS_ALLOWED=0
for h in "${ALLOWED_HOURS[@]}"; do
    if [ "$CURRENT_HOUR" == "$h" ]; then
        IS_ALLOWED=1
        break
    fi
done

if [ $IS_ALLOWED -eq 0 ]; then
    log "当前时段 ($CURRENT_HOUR) 不在允许时间范围内，跳过"
    exit 0
fi

# 读取队列
if [ ! -f "$QUEUE_FILE" ]; then
    log "队列文件不存在"
    exit 1
fi

# 检查是否有待执行任务
TASK_COUNT=$(node -e "const q=require('$QUEUE_FILE'); console.log(q.queue.length)")

if [ "$TASK_COUNT" -eq 0 ]; then
    log "队列为空，无需执行"
    exit 0
fi

log "队列中有 $TASK_COUNT 个任务，开始执行..."

# 执行第一个任务（可以扩展为批量执行）
node -e "
const fs=require('fs');
const q=JSON.parse(fs.readFileSync('$QUEUE_FILE', 'utf8'));
if(q.queue.length > 0) {
    const task=q.queue[0];
    console.log(JSON.stringify(task));
}
"

log "任务调度完成"
