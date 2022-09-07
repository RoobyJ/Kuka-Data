#!Python3.8.10
import os
import asyncio
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
    while True:
        input_string = input("Enter path to backups:")
        if os.path.exists(os.path.dirname(input_string)):
            break

    for backup in os.listdir(input_string):
        await config_to_excel(backup, input_string)
    print(mean())


if __name__ == "__main__":
    asyncio.run(main())
