#!/bin/bash
temp=$(sensors | grep 'Core 3' | awk '{print substr($3,2,2)}')
RecordTime=$(date +"[%Y-%m-%d %H:%M.%S]")
echo "$RecordTime" >> /root/temp/temp.log
sensors | grep Core >> /root/temp/temp.log
if [ "$temp" -gt 70 ];then
	echo "poweroff" >> /root/temp/temp.log
	`poweroff`
else
	echo "no" >> /root/temp/temp.log
fi
