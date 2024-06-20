import logging as L
import os

directory = 'TNR-71'
parent_dir = "C:/Users/"+os.getlogin()+"/Desktop/"
Directory = os.path.join(parent_dir, directory)
#now we will Create and configure logger
L.basicConfig(filename = Directory + "/logs.txt",
              format = '%(asctime)s %(message)s',
              filemode = 'a')

#Let us Create an object
logger = L.getLogger()

#Now we are going to Set the threshold of logger to DEBUG
logger.setLevel(L.INFO)
