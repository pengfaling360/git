import os
key = 171
print("本软件只针对office办公软件打不开或乱码的情况进行恢复，恢复前请检查特征码。")
path1=input("请输入文件所在路径:")
path=path1
def file_move(path,target):
    try:
        os.chdir(path)
        all_file =os.listdir(os.curdir)
        for each in all_file:
            if os.path.isdir(each):
                file_move(each,target)
                os.chdir(os.pardir)
            else:
                ext = os.path.splitext(each)[1]
                print(os.path.splitext(each)[0]+"."+os.path.splitext(each)[1] +"  已成功恢复")
                if ext in target:
                    namel=(os.getcwd() +os.sep+each+ os.linesep).strip()
                    with open (namel, "rb+") as f:
                         data=f.read()
                         data_arr=bytearray(data)
                         for i,v in enumerate(data_arr):
                             data_arr[i]= v ^ key
                         with open (namel, 'wb') as f:
                             f.write(data_arr)
                             f.close()
    except:
        pass
path= path1
target=['.doc','.docx','.ppt','.pptx','.xlsx']
file_move(path,target)
