#!Python3.8.10
import os
import asyncio
import sys
from DataMiner import DataMiner


async def config_to_excel(_backup, _input_string, _time_list):
    obj = DataMiner(_backup, _input_string)
    await obj.create_file()
    resp = obj.absolute_time()
    if resp is not None:
        _time_list.append(resp)


# by making the program async we save much time example for 32 backups async does it in ~2.9s, synchronous needs ~5s
async def main():
    time_list = []
    input_string = sys.argv[1]
    if os.path.exists(os.path.dirname(input_string)):
        for backup in os.listdir(input_string):
            await config_to_excel(backup, input_string, time_list)
        print(sum(time_list))
    else:
        print('given path is incorrect \n should look like this example "C:\\Projects\\NameofProject\\Backups"')

if __name__ == "__main__":
    asyncio.run(main())
