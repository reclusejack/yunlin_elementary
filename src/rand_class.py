import argparse
import copy
import random
import re
import pprint
#import pdb

def rand_class(teacher_list):
    rand_th_list = []

    pprint.pprint(teacher_list)
    rand_th_list.append([x for x in teacher_list if "一" in x[1]])
    rand_th_list.append([x for x in teacher_list if "二" in x[1]])
    rand_th_list.append([x for x in teacher_list if "三" in x[1]])
    rand_th_list.append([x for x in teacher_list if "四" in x[1]])
    rand_th_list.append([x for x in teacher_list if "五" in x[1]])
    rand_th_list.append([x for x in teacher_list if "六" in x[1]])

    for num, grade in enumerate(rand_th_list):
        if grade != []:
            unassigned_class = list(range(1, len(grade) + 1))
            for seq in range(0, len(grade)):
                sec_rand = random.SystemRandom()
                rand_th_list[num][seq][4] = sec_rand.choice(unassigned_class)
                unassigned_class.remove(rand_th_list[num][seq][4])

    return rand_th_list
