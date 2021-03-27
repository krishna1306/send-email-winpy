import csv
import os

if __name__=="__main__":

    with open('new_mem.txt','r') as f:
        csv_python = csv.reader(f)
        
        for row in csv_python:
            print(row)