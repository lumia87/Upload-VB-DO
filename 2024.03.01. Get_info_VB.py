from docx import Document #pip install python-docx/conda install -c conda-forge python-docx
import datetime
import win32com.client #conda install pywin32
import re
import json
import time
import os
import sys

def saveasdocx(file_path):
     word = win32com.client.Dispatch('Word.Application')
     word.Visible = False
     docx_file = '{0}{1}'.format(file_path, 'x')
     try:
          wordDoc = word.Documents.Open(file_path, False, False, False)
          wordDoc.SaveAs2(docx_file, FileFormat = 16)
          wordDoc.Close()
     except Exception as e: 
          print('Failed to Convert: {0}'.format(file_path))
          print(e)

def getinfo(doc_path):     
#Lấy thông tin của Doucument ở doc_path
     #trả về 4 thông tin: Trích yếu;  Người ký;  Nơi nhận
     #1.Nơi nhận
     trichyeu=""
     noinhan_th=[] #nhận để thực hiện
     noinhan_db=[] #nhận để biết
     noinhan_tn=[] #nhận trong ngành
     nguoi_ky=""
     #bắt đầu chương trình
     document = Document(doc_path)
     if len(document.tables)>1:
        table=document.tables[1] #bảng ở cuối danh sách
        cell=table.rows[0].cells[len(table.columns)-1]
        last_p=len(cell.paragraphs)
        tmp=cell.paragraphs[last_p-1].text.strip()
        if tmp=="Phan Sỹ Bách":
            nguoi_ky="BACPS"
        elif tmp=="Lê Đặng Xuân Tân":
            nguoi_ky="TANLDX"

        #tìm nội dung nơi nhận ở ô Nơi nhận:        
        cell=table.rows[0].cells[0]
        for p in cell.paragraphs:
            if "A0" in p.text and "A0" not in noinhan_tn:
                        noinhan_tn.append("A0")

            if ("b/c" or "bc" or "báo cáo" or "biết") in p.text: #Chỉ có xuất hiện chữ báo cáo hay để biết mới cho vô List nơi nhận để biết, còn lại là cho vô nơi nhận để Thực hiện
                if "các phòng" in p.text.lower():
                    noinhan_db= noinhan_db+["TH","KH","TCKT","DD","PT","CN"]
                elif ("bgđ" in p.text.lower()) or ("ban giám đốc" in p.text.lower()) or ("ban gđ" in p.text.lower()):
                    noinhan_db.append("BACHPS","TANLDX")
                else:
                    if "TH" in p.text and "TH" not in noinhan_db:
                        noinhan_db.append("TH")
                    if "KH" in p.text and "KH" not in noinhan_db:
                        noinhan_db.append("KH")
                    if ( ("TCKT" in p.text) or ("TC-KT" in p.text)) and "TCKT" not in noinhan_db:
                        noinhan_db.append("TCKT")
                    if "DD" in p.text and "DD" not in noinhan_db:
                        noinhan_db.append("DD")
                    if "PT" in p.text and "PT" not in noinhan_db:
                        noinhan_db.append("PT")
                    if "CN" in p.text and "CN" not in noinhan_db:
                        noinhan_db.append("CN")
                    if "GĐ" in p.text and "GĐ" not in noinhan_db:
                        noinhan_db.append("GD")
                    if "Tân" in p.text and "TANLDX" not in noinhan_db:
                        noinhan_db.append("TANLDX")
                    if "Bách" in p.text and "Bách" not in noinhan_db:
                        noinhan_db.append("BACHPS") 
            else:
                if "các phòng" in p.text.lower():
                    for ph in ["TH","KH","KT","DD","PT","CN"]:
                        if ph not in noinhan_th:
                            noinhan_th.append(ph)
                elif ("bgđ" in p.text.lower()) or ("ban giám đốc" in p.text.lower()) or ("ban gđ" in p.text.lower()):
                    noinhan_th.append("BACHPS","TANLDX")
                else:
                    if "TH" in p.text and "TH" not in noinhan_th:
                        noinhan_th.append("TH")
                    if "KH" in p.text and "KH" not in noinhan_th:
                        noinhan_th.append("KH")
                    if "TCKT" in p.text and "TCKT" not in noinhan_th:
                        noinhan_th.append("TCKT")
                    if "DD" in p.text and "DD" not in noinhan_th:
                        noinhan_th.append("DD")
                    if "PT" in p.text and "PT" not in noinhan_th:
                        noinhan_th.append("PT")
                    if "CN" in p.text and "CN" not in noinhan_th:
                        noinhan_th.append("CN")
                    if "GĐ" in p.text and "GĐ" not in noinhan_th:
                        noinhan_th.append("BACHPS")
                    if "Tân" in p.text and "TANLDX" not in noinhan_th:
                        noinhan_th.append("TANLDX")
     else:
         #khong su dung bang
         s_noinhan=""
         ghi=False
         print("khong dung bang o chu ky")
         for p in document.paragraphs:
             if "Nơi nhận:" in p.text:
                 ghi=True
             if ghi==True:
                s_noinhan=s_noinhan+"\n"+(p.text).strip()
                if "A0" in p.text and "A0" not in noinhan_tn:
                    noinhan_tn.append("A0")

                if ("b/c" or "bc" or "báo cáo" or "biết") in p.text: #Chỉ có xuất hiện chữ báo cáo hay để biết mới cho vô List nơi nhận để biết, còn lại là cho vô nơi nhận để Thực hiện
                    if "các phòng" in p.text.lower():
                        noinhan_db= noinhan_db+["TH","KH","TCKT","DD","PT","CN"]
                    elif ("bgđ" in p.text.lower()) or ("ban giám đốc" in p.text.lower()) or ("ban gđ" in p.text.lower()):
                        noinhan_db.append("BACHPS","TANLDX")
                    else:
                        if "TH" in p.text and "TH" not in noinhan_db:
                            noinhan_db.append("TH")
                        if "KH" in p.text and "KH" not in noinhan_db:
                            noinhan_db.append("KH")
                        if ( ("TCKT" in p.text) or ("TC-KT" in p.text)) and "TCKT" not in noinhan_db:
                            noinhan_db.append("TCKT")
                        if "DD" in p.text and "DD" not in noinhan_db:
                            noinhan_db.append("DD")
                        if "PT" in p.text and "PT" not in noinhan_db:
                            noinhan_db.append("PT")
                        if "CN" in p.text and "CN" not in noinhan_db:
                            noinhan_db.append("CN")
                        if "GĐ" in p.text and "GĐ" not in noinhan_db:
                            noinhan_db.append("GD")
                        if "Tân" in p.text and "TANLDX" not in noinhan_db:
                            noinhan_db.append("TANLDX")
                        if "Bách" in p.text and "Bách" not in noinhan_db:
                            noinhan_db.append("BACHPS") 
                else:
                    if "các phòng" in p.text.lower():
                        for ph in ["TH","KH","KT","DD","PT","CN"]:
                            if ph not in noinhan_th:
                                noinhan_th.append(ph)
                    elif ("bgđ" in p.text.lower()) or ("ban giám đốc" in p.text.lower()) or ("ban gđ" in p.text.lower()):
                        noinhan_th.append("BACHPS","TANLDX")
                    else:
                        if "TH" in p.text and "TH" not in noinhan_th:
                            noinhan_th.append("TH")
                        if "KH" in p.text and "KH" not in noinhan_th:
                            noinhan_th.append("KH")
                        if "TCKT" in p.text and "TCKT" not in noinhan_th:
                            noinhan_th.append("TCKT")
                        if "DD" in p.text and "DD" not in noinhan_th:
                            noinhan_th.append("DD")
                        if "PT" in p.text and "PT" not in noinhan_th:
                            noinhan_th.append("PT")
                        if "CN" in p.text and "CN" not in noinhan_th:
                            noinhan_th.append("CN")
                        if "GĐ" in p.text and "GĐ" not in noinhan_th:
                            noinhan_th.append("BACHPS")
                        if "Tân" in p.text and "TANLDX" not in noinhan_th:
                            noinhan_th.append("TANLDX")
        
     #Tên người ký
     
   

     ghi=False
 
     for p in document.paragraphs:
         if ("QUYẾT ĐỊNH" in p.text):
             ghi=True
         elif "GIÁM ĐỐC" in p.text: #dừng lại khi thấy chữ Giám đốc đối với Quyết định
             ghi=False
             break

         if ghi==True:
             if ("QUYẾT ĐỊNH" in p.text):
                trichyeu=(p.text).lower().capitalize()+" "                
             elif trichyeu.strip()=="Quyết định":
                trichyeu=trichyeu+" "+(p.text).lower()     #Dòng ngày sau quyết định
             else:
                trichyeu=trichyeu+" "+p.text       

     trichyeu = re.sub(r'[/\():?]', ' ', trichyeu)
     trichyeu = re.sub('\s+',' ',trichyeu) #chỉ để lại tối đa 1 ký tự trống
     trichyeu=trichyeu.strip()
     trichyeu
       

    
     #for i in range(0,len(noinhan)): #lưu ý là range thì sẽ chạy đến cận trên -1 nên lấy len luôn
     #    print("Nơi nhận",i,noinhan[i])
     with open('ds_vb.txt','w', encoding='utf8') as f:
         f.write(json.dumps({'path':doc_path, 'trich_yeu':trichyeu,'nguoi_ky':nguoi_ky,'nn_th':noinhan_th,'nn_db':noinhan_db, 'nn_tn':noinhan_tn}, ensure_ascii=False))
     return (doc_path, trichyeu,nguoi_ky,noinhan_th,noinhan_db, noinhan_tn)

print(getinfo('QD.docx'))