#!/usr/bin/python3
import sys
from rand_class import *
from load_data import *
from write_data import *
#import pdb

orig_data = []
group_data = []
rule = []
if len(sys.argv) != 2:

    print("%d\n" %len(sys.argv))
else:
    filename = sys.argv[1]
    #pdb.set_trace()
    teacher_list = load_data(filename)
    rand_teacher_list = rand_class(teacher_list)
    #writefile(data, filename, total_class, boy_class, girl_class)
