from win32com.client import Dispatch
from os import walk
import os

wdFormatPDF = 17

def doc2pdf(input_file):
    word = Dispatch('Word.Application')
    doc = word.Documents.Open(input_file)
    if input_file.endswith(".doc"):
        doc.SaveAs(input_file.replace(".doc", ".pdf"), FileFormat=wdFormatPDF)
    if input_file.endswith(".docx"):
        doc.SaveAs(input_file.replace(".docx", ".pdf"), FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()

# 遍历文件夹及其子文件夹中的文件，并存储在一个列表中
# 输入文件夹路径、空文件列表[]
# 返回 文件列表Filelist,包含文件名（完整路径）
def get_filelist(dir, Filelist):
    newDir = dir
    if os.path.isfile(dir):
        Filelist.append(dir)
        # # 若只是要返回文件文，使用这个
        # Filelist.append(os.path.basename(dir))
    elif os.path.isdir(dir):
        for s in os.listdir(dir):
            # 如果需要忽略某些文件夹，使用以下代码
            # if s == "xxx":
            # continue
            newDir = os.path.join(dir, s)
            get_filelist(newDir, Filelist)
    return Filelist

def filename_find(filepath, return_type=0):
    basename = os.path.basename(filepath)
    extension = f'.{basename.split(".")[-1]}'
    if not basename.__contains__('.'):
        extension = ''
    filename_without_extension = basename[0:len(basename)-len(extension)]
    if return_type is 0:    # 文件名
        return basename
    if return_type is 1:    # 后缀名
        return extension
    if return_type is 2:    # 无后缀文件名
        return filename_without_extension

if __name__ == '__main__':

    #list = get_filelist(image_path, [])
    #################################################################
    # 只要把文件路径替换下边的Y:\word2pdf\上交，文件夹里的所有word都会变成pdf #
    #################################################################
    directory = "Y:\word2pdf\上交"

    for root, dirs, filenames in walk(directory):
        print(root)
        for file in filenames:
            '''
            #删除word文件
            if file.endswith(".doc") or file.endswith(".docx"):
                os.remove(str(root + "\\" + file))
            '''
            temp_file_name=filename_find(file,2)+'.pdf'
            print(temp_file_name)
            temp_filex_name = filename_find(file, 2) + '.pdfx'
            if temp_file_name in filenames or temp_filex_name in filenames:
                continue
            #word有时候抽风出的处理文件根本不可见，直接跳过
            if file.startswith("~$"):
                continue
            if file.endswith(".doc") or file.endswith(".docx"):
                print(str(root + "\\" + file))
                doc2pdf(str(root + "\\" + file))
    print('处理完成')

