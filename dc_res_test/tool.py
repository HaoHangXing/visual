import os
import shutil
import re


       

# 判断文件是否存在
def IsExist(file):
    # print("currunt path:", os.getcwd())
    if os.path.exists(file):
        print("found %s" % file)
        return True
    else:
        print("cannot find %s" % file)
        return False

# 删除原有的文件夹，包括里面的内容。重新创建一个dir。
def RemoveCreatFolder(dir):
    if IsExist(dir):
        print('delete previous : '+ dir)
        shutil.rmtree(dir)

    os.mkdir(dir)

# 删除指定字符串,pattern正则表达式格式
def DeleteStr(str, pattern):
    return re.sub(pattern, "", str)

# 找到指定字符串类型
def FindPatternStr(str, pattern):
    return re.findall(pattern, str)
    # return re.finditer(pattern, str)

# 找到提取数字
def GetNumformStr(str):
    num_list = re.findall('\d+', str)
    for i, num in enumerate(num_list):
        num_list[i] = int(num)
    return num_list

# 找到时间数字
def FindTime(str, pattern):
    time = FindPatternStr(str, pattern)
    return GetNumformStr(time[0])
