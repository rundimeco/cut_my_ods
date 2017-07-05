#coding:utf-8
import re
import sys
sys.path.append("../../../libs/")
import glob
import os
from pyexcel_ods import save_data
from pyexcel_ods import get_data

def write_ods(l1, lines, next_cut, fic):
  window = "%s-%s"%(str(next_cut-len(lines)+1),str(next_cut))
  Fname = re.sub("\.ods", "_%s.ods"%(window), fic)
  data= {"Feuil1":[]}
  for l in lines:
    data["Feuil1"].append(l)
  save_data(Fname, data)

def cut_wb(fic, cut=1000):
  print "\nWill cut the original file (%s) in files with %s lines at most\n"%(fic, str(cut))
  data = get_data(fic)
  NB_lignes = 0
  ligne_debut= []
  current_cut = []
  next_cut = cut
  for feuille, rows in data.iteritems():
    current_line = []
    for row in rows:
      if NB_lignes==0:
          ligne_debut = row
      else: 
          current_line=row
      if NB_lignes>0:
        current_cut.append(current_line)
      if NB_lignes==next_cut:
        write_ods(ligne_debut, current_cut, next_cut, fic)
        next_cut+=cut
        current_cut = []
      NB_lignes+=1
  if len(current_cut)>0:
    write_ods(ligne_debut, current_cut, NB_lignes-1, fic)

if __name__=="__main__":
  cut = 1000
  if len(sys.argv)<2:
    print "Usage : path-to-file number-of-lines(default=%s)"%str(cut)
    exit()
  elif len(sys.argv)>2:
    cut = int(sys.argv[2])
  fic = sys.argv[1]
  cut_wb(fic, cut)
