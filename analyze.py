#!/usr/bin/env python

import sys
import xlrd

def main():
  print sys.argv[1]
  workbook = xlrd.open_workbook(sys.argv[1])

if __name__ == '__main__':
  main()
