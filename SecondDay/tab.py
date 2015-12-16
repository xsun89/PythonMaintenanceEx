#!/usr/bin/env python 
# this content comes from oldboy trainning.
# e_mail:31333741@qq.com
# qqinfo:49000448
# function: python tab config.
# version:1.1 
################################################
# oldboy trainning info.      
# QQ 1986787350 70271111
# site:http://www.etiantian.org
# blog:http://oldboy.blog.51cto.com
# oldboy trainning QQ group: 208160987 45039636
################################################
# python startup file
import sys
import readline
import rlcompleter
import atexit
import os
# tab completion
readline.parse_and_bind('tab: complete')
# history file
histfile = os.path.join(os.environ['HOME'], '.pythonhistory')
try:
    readline.read_history_file(histfile)
except IOError:
    pass
atexit.register(readline.write_history_file, histfile)


del os, histfile, readline, rlcompleter

