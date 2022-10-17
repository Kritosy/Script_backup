#/bin/bash

botPid=$(ps -ef | grep -v grep | grep python3 | awk '{print $2}')
if [ $botPid ]; then
        echo "started... ready to stop"
        kill -9 $botPid
	      screen -r bot -X quit
else
        echo "stoped... ready to start"
fi
#创建screen
screen -dm bot

#在screen中执行命令
#需要执行的命令最好都放在另外一个脚本中执行
screen -x -S bot -p 0 -X stuff "bash start_bot.sh\n"
