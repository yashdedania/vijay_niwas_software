from distutils.core import setup
import py2exe
import sys
#import numpy
sys.setrecursionlimit(5000)

#setup(console=['vijay_niwas.py'])

setup(windows=[{"script":"vijay_niwas.py"}],
      options={"py2exe":
               {"excludes":["matplotlib","pandas","scipy","numpy"],
                "includes":["sip"]
                }})

'''from distutils.filelist import findall
import os
import matplotlib
matplotlibdatadir = matplotlib.get_data_path()
matplotlibdata = findall(matplotlibdatadir)
matplotlibdata_files = []
for f in matplotlibdata:
    dirname = os.path.join('matplotlibdata', f[len(matplotlibdatadir)+1:])
    matplotlibdata_files.append((os.path.split(dirname)[0], [f]))

setup(
    console=['test.py'],
    options={
        'py2exe': {
            'packages' : ['matplotlib', 'pytz'],
            }
        },
    data_files=matplotlibdata_files
    )
'''
