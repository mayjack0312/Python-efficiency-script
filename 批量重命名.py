import office

path = 'r' + input('文件夹路径：')
office.file.replace4filename(
                              path=path,
                              del_content=input('要去掉的内容：'),
                              replace_content=input('要替换的内容：')
                              )