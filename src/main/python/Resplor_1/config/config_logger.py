# config file that configures the logger for the whole project
# import this file into any project to have logging available with those settings


import logging

logger = logging.getLogger('MainLog')
logger.setLevel(logging.DEBUG)
logfile = 'logfile.log'

# empty file first
open(logfile, 'w').close
logfilehandler = logging.FileHandler(logfile)

# set format for file logging
logformat = logging.Formatter('%(asctime)s | %(name)s | %(levelname)s | %(funcName)s:%(lineno)d |%(message)s')
logfilehandler.setFormatter(logformat)

#set level for file logging
logfilehandler.setLevel(logging.DEBUG)

# set logger to use this logfile
logger.addHandler(logfilehandler)