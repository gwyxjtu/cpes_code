'''
Author: gwyxjtu
Date: 2022-06-18 15:11:07
LastEditors: gwyxjtu 867718012@qq.com
LastEditTime: 2022-06-18 15:11:39
FilePath: /git_code/log/grblogtool.py
Description: 人一生会遇到约2920万人,两个人相爱的概率是0.000049,所以你不爱我,我不怪你.

Copyright (c) 2022 by gwyxjtu 867718012@qq.com, All Rights Reserved. 
'''
import grblogtools as glt

summary = glt.parse("mip.log").summary()
glt.plot(summary)
