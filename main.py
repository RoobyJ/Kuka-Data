#!Python3.8.10
import os
from DataMiner import DataMiner


def mean(time_list):
    temp = float
    for var in time_list:
        temp += var
    return temp


def main():
    time_list = []
    while True:
        input_string = input("Enter path to backups:")
        if os.path.exists(os.path.dirname(input_string)):
            break

    for backup in os.listdir(input_string):
        resp = DataMiner(backup, input_string).create_file()
        time_list.append(resp)

    print(mean(time_list))


if __name__ == "__main__":
    main()
