#-*- coding: utf-8 -*-

from os import path, remove, renames
from openpyxl import load_workbook
from shutil import rmtree
import sys

VERSION = "1.1"
FILE_NAME = 'renamer.xlsx'

args = sys.argv[1:]

if args:
    if path.isdir(args[0]):
        arg_path = args[0]
        if arg_path[-1] != '\\':
            arg_path += '\\'
        file = arg_path+FILE_NAME
    elif path.isfile(args[0]):
        file = args[0]
else:
    file = FILE_NAME

print(file)


try:
    wb = load_workbook(file, data_only=True)
except:
    print("renamer.xlsx not found")
    quit()
ws = wb['rename']

names = []
rows = list(ws.rows)[1:]
for row in rows:
    before_path = row[1].value
    before_name = row[2].value
    after_path = row[3].value
    after_name = row[4].value
    if after_path:
        if after_path[-1] != '\\':
            after_path += '\\'
    before = before_path + before_name if before_name else before_path
    after = after_path + after_name if after_name else after_path
    if '-u' in args:
        names.append([after, before])
    else:
        names.append([before, after])

names.reverse()

for name in names:
    if name[0]==name[1]:
        print("[remained]",name[0])
        continue
    else:
        if name[1]==None:
            try:
                if path.isdir(name[0]):
                    rmtree(name[0])
                else:
                    remove(name[0])
                print("[removed]",name[0])
            except:
                print("[error removing]",name[0])
            continue
        if name[0] and name[1]:
            try:
                renames(name[0], name[1])
                print("[renamed]",name[0])
            except:
                print("[error renaming]",name[0])
                continue
    print("[skipped]",name[0])
    continue