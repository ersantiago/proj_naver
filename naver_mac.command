#!/bin/bash
cwd=`pwd`
navs=`find . * | grep naver_gui_v3_mac.py | tail -1`
navp="${cwd}/${navs}"
echo $navp
python $navp
