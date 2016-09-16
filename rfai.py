#! /usr/bin/env python3

# Michael Fruhnert, Sep 2016
# file used as launcher for actual scripts
# Can place it anywhere on engr14x

import os, sys

gsMainScripts = "/share/engr14x/12 Grading Systems 2.0/06 14xScripts/"
#if Mac terminal is detected, then paths need to change
if not(os.path.exists("/share/")):
    print("Mac detected") #change /share/ to /Volumes/
    gsMainScripts = "/Volumes" + gsMainScripts[6:]

sys.path.insert(0, gsMainScripts) # so it finds setupDirAndStuff

import setupDirAndStuff as dir14x # import first
dir14x.addPathsToLibraries(False) # add paths to other libraries

# read command line parameter
if ((len(sys.argv) > 1) and (sys.argv[1].lower() == 'admin')):
    import rfaiAdmin # will actually launch it too
else:
    import rfaiPTA # will launch it