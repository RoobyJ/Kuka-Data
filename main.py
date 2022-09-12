#!Python3.8.10
import os
import asyncio
import sys
from DataMiner import DataMiner

time_list = []


def mean():
    temp = 0
    for var in time_list:
        temp += var
    return temp


async def config_to_excel(_backup, _input_string):
    obj = DataMiner(_backup, _input_string)
    await obj.create_file()
    resp = obj.absolute_time()
    if resp is not None:
        time_list.append(resp)


async def main():
    input_string = sys.argv[1]
    if os.path.exists(os.path.dirname(input_string)):
        for backup in os.listdir(input_string):
            await config_to_excel(backup, input_string)
        print(mean())
    else:
        print('given path is incorrect \n should look like this example "C:\\Projects\\NameofProject\\Backups"')

if __name__ == "__main__":
    asyncio.run(main())
