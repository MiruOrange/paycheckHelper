from PyPDF2 import PdfReader, PdfWriter
from docx2pdf import convert

#將word檔轉換為pdf檔
#convert(x, y)
#x = 讀取的檔名; y= 輸出的檔名
convert("input.docx", "output.pdf")

#將讀取pdf文件的類別物件化
#PdfReader(x)
#x = 要讀取的pdf檔名
reader = PdfReader("output.pdf")

#設定pdf密碼
employeeData =[
    {'張無忌':'A12345678'},{'小昭':'B12345678'},{'趙敏':'C12345678'}
]

#將pdf文件逐一加密後另存新檔
i = 0
for page in reader.pages:
    for key, value in employeeData[i].items():
        fn = key+'_1月薪資明細.pdf'
        writer = PdfWriter()    #每次進迴圈就重新物件化，清除前一頁資料
        writer.add_page(page)   #每次只加入新一頁的資料
        writer.encrypt(value)   #設定文件加密密碼
        #將文件另存為新的pdf檔，隨頁數而變化，有n頁，就會產生n個新的檔
        with open(fn, "wb") as f:
            writer.write(f)
    i+=1