#!/bin/bash
PingInfo=`ping -c 6 10.9.20.1 >> Ping.log`
tail -n 1 $PingInfo