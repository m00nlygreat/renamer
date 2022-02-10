#-*- coding: utf-8 -*-

from os import path, remove, renames
from openpyxl import load_workbook
from shutil import rmtree
import sys

args = sys.argv[1:]

try:
    wb = load_workbook("renamer.xlsx")
except:
    print("renamer.xlsx not found")
    quit()
ws = wb['rename']

names = []
rows = list(ws.rows)[1:]
for row in rows:
    before = row[1].value+row[2].value if row[2].value else row[1].value
    after = row[3].value+row[4].value if row[4].value else row[3].value
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