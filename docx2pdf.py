import office

path = input('输入存放word文件的位置(仅支持.docx)：')
office.word.docx2pdf(path=path)