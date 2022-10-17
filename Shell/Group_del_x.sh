#!/bin/bash
##此脚本用来对文件批量删除组执行权限
##脚本需要在目标路径下执行
##自定义第二个函数，可是实现任何对文件的批量操作
function read_dir(){  #遍历文件夹和目录
for file in `ls $1` 
do
 if [ -d $1"/"$file ] 
 then
 read_dir $1"/"$file #递归读取目录
 else
 chk_group "$1/$file" #将两个参数合并成一个参数。此处对文件的进行操作，不能对文件夹操作
 fi
done
} 
function chk_group(){
        GP_name=$(ls -l $1 | awk '{print $4}')  #awk以空格或table分隔，组名为第四个（$4）
        if [ "$GP_name"x == "Group_name"x ] #将获取的组名与预设的Group_name比较
        then
                chmod g-x $1
                echo $1 has delete x
        else
                echo $1 is not belong to Group_name
        fi
}
dir=`pwd` #获取当前目录
read_dir $dir 
