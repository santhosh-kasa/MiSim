'''
Created on Jan 5, 2015

@author: santhosh
'''
f = open('myfile','w')
f.write('hi there\n') # python will convert \n to os.linesep
f.close()