#!/usr/bin/python
# Copyright 2010 Google Inc.
# Licensed under the Apache License, Version 2.0
# http://www.apache.org/licenses/LICENSE-2.0

# Google's Python Class
# http://code.google.com/edu/languages/google-python-class/

import sys
import re
import os
import shutil
import subprocess
from zipfile import ZipFile
#import zipfile

"""Copy Special exercise
"""

# +++your code here+++
# Write functions and modify main() to call them
def list_dir(dir):
  filenames = os.listdir(dir)
  filename_list = []
  for filename in filenames:
    match = re.search(r'[\w.]+__[\w.]+__', filename)
    if match:
      abspath = os.path.abspath(os.path.join(dir, filename))
      filename_list.append(abspath)
    else:
      print ('filename doesn\'t match', filename)
  return filename_list
  
def copy_to_dir(filename_list,todir):
  if not os.path.exists(todir):
    os.mkdir(todir)
  for filename in filename_list:
    shutil.copy(filename, todir)
    
def zip_to(filename_list, zip_file):
  zipobject = ZipFile(zip_file, 'w')
  for filename in filename_list:
    dirname = os.path.dirname(filename)
    print (dirname)
    print (os.path.basename(filename))
    zipobject.write(os.path.basename(filename))
  zipobject.close()

def main():
  # This basic command line argument parsing code is provided.
  # Add code to call your functions below.

  # Make a list of command line arguments, omitting the [0] element
  # which is the script itself.
  args = sys.argv[1:]
  if not args:
    print ("usage: [--todir dir][--tozip zipfile] dir [dir ...]")
    sys.exit(1)

  # todir and tozip are either set from command line
  # or left as the empty string.
  # The args array is left just containing the dirs.
  todir = ''
  if args[0] == '--todir':
    todir = args[1]
    del args[0:2]

  tozip = ''
  if args[0] == '--tozip':
    tozip = args[1]
    del args[0:2]

  if len(args) == 0:
    print ("error: must specify one or more dirs")
    sys.exit(1)

  # +++your code here+++
  # Call your functions
  for arg in args:
    filename_list = list_dir(arg)
  
  if todir:
    copy_to_dir(filename_list, todir)
  elif tozip:
    zip_to(filename_list, tozip)
  else:
    print ('\n'.join(filename_list))
  
if __name__ == "__main__":
  main()
