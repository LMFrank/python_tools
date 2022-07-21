import os
import shutil

# 对象文件的类型指定
file_type_list = ['pdf', 'txt', 'xls', 'xlsx', 'pptx', 'doc']


def get_file_list(folder):
    filelist = []
    for dirpath, dirnames, filenames in os.walk(folder):
        for file in filenames:
            file_type = file.split('.')[-1]
            if file_type in file_type_list:
                file_fullname = os.path.join(dirpath, file)
                filelist.append(file_fullname)
    return filelist


# 将文件list里面的文件拷贝到指定目录下
def copy_file(src_file_list, dst_folder):
    print('===========copy start===========')
    for file in src_file_list:
        shutil.copy(file, dst_folder)
    print('===========copy end!===========')


if __name__ == "__main__":
    # copy源所在目录
    src_folder = r'D:\新建文件夹\新建文件夹 (5)'
    # copy到的指定目录
    dst_folder = r'D:\新建文件夹\新建文件夹 (4)'

    # 取得文件夹下所有指定类型的文件全名
    filelist = get_file_list(src_folder)
    copy_file(filelist, dst_folder)
