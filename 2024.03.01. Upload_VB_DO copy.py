# Create 01.03.2024

import requests
import os
import json
import time
from datetime import datetime
from datetime import datetime, timezone

from requests import Request, Session
from concurrent.futures import ThreadPoolExecutor, as_completed


# (23.02.2024.21h00) Sử dụng Request.prepare() thi duoc. Tham khao tai:
# https://experienceleaguecommunities.adobe.com/t5/adobe-experience-platform/400-bad-request-when-i-try-to-send-a-post-request-to-the-auth/m-p/593183
lst_vb=[
    {'stt': 1,'path_vb':'KH.docx','trich_yeu':'KH công tác quốc phòng, quân sự địa phương (A2)', 'nn_ct':'th', 'nn_ph':['cb', 'bchqs'], 'nn_db':['vt','bachps', 'tanldx'], 'nn_tn':['A0']
    },
    {
    'stt': 2,'path_vb':'QD.docx','trich_yeu':'QĐ công nhận xếp loại hoàn thành nhiệm vụ năm 2024'
    }
]

phongban = {
    "DD": 39594,     
    "CN": 39596,    
    "PT": 39595,   
    "KH": 39599,   
    "TCKT":39598,     
    "KSAT":39623,
    "TH": 39597,
    "BACHPS":44194,
    "TANLDX":39630,
    "CB": 35844,
    "VT": 39611,
    "BCHCD":39621,
    "BCHQS":43850
}

donvi={
     "A0": 110,
}


def noinhan_dv(ky_hieu):
    res=[]
    try:
        for dv in ky_hieu:
            res.append({'ID_DV':donvi[dv.upper()]})
    except:
        print(f"Khong co don vi ma: {ky_hieu}")
    return res

def noinhan_ct(ky_hieu):
    res=[]
    try:
        res = [{'ID_PB':phongban[ky_hieu.upper()]}]
    except:
        print(f"Khong co phong ban ma: {ky_hieu}")
    return res

def noinhan_ph(ky_hieu):
    res=[]
    try:
        for pb in ky_hieu:
            res.append({'ID_PB':phongban[pb.upper()]})
    except:
        print(f"Khong co phong ban ma: {ky_hieu}")
    return res

def get_authorization(username, password):
    url = "https://gwdoffice.evn.com.vn/v3/auth/Auth/DAuth"
    headers = {
        "content-type": "application/json",
        "Sec-Ch-Ua": "'Not A(Brand';v='99','Google Chrome';v='121', 'Chromium';v='121'",
        "Accept": "application/json, text/plain, */*",
        "Sec-Ch-Ua-Platform": "Windows",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
    }

    body = {
        "username": username,
        "password": password,
        "expiration": 60,
        "deviceInfo": {
            "deviceId": "9433b780-e98c-4e2b-8d60-eb61e12a649d",
            "deviceType": "windows-10/desktop/Chrome",
            "appId": "DOFFICE",
            "appVersion": "v2.0.0",
        },
    }

    request = Request(method="POST", url=url, headers=headers, json=body)
    session = requests.Session()
    prepped = session.prepare_request(request)

    r = session.send(prepped)
    result = []
    if r.status_code == 200:
        result = "Bearer " + r.json()["Data"]["accessToken"]
    return result



def upload_vb(username, password, vb):
    
    url ='https://gwdoffice.evn.com.vn/v1/duthao/VBDT/UploadFile'

    Authorization = get_authorization(username, password)

    payload={
            "id_DV":"114",
            "ID_NV":"100141",
            "File_Chinh":"true",
            "Check":"false", #neu true se tao ra ma filebase64 cua pdf; false se khong tao ma to ra Duong dan
            "nguoi_ky":"Phan Sỹ Bách",
            "check_kyCA":"false",
            "LOAI_DK":"KY_SO",
            "ID_PB_CURRENT":"39597",
            "NGHACH":"HC",
            "NGON_NGU":"VI",
            "MAU_DU_THAO":"",
            "ID_LOAI_VB":"140615",
        }

    # files=[('files',('QD.docx',open('./QD.docx','rb'),'application/vnd.openxmlformats-officedocument.wordprocessingml.document'))]
    files = {'files': open(vb['path_vb'], 'rb')}
    # headers = {"Content-Type": "application/json"}
    # The data parameter expects a byte string, which is why we convert the JSON string to a byte string using json.dumps().
    r = requests.post( url, headers={"Authorization": Authorization,}, data = payload, files=files)
    
    if r.status_code==200:
        # print(r.text)
        
        with open("response.txt", "w") as f:
            
            files_info=json.loads(r.content)['Data']['responseData']
            f.write(str(files_info))

            #chen don vi trinh
            
            headers = {
                "content-type": "application/json",
                "Authorization": Authorization,
            }

            url2='https://gwdoffice.evn.com.vn/v1/duthao/VBDT_LD/InsertDTVB_DangKy'

            payload2={
                "ID_DT": "0",
                "TRICH_YEU": vb['trich_yeu'],
                "SO_TRANG": "0",
                "NGAY_TRINH": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3]+"Z",
                "HAN_XULY": "", #sua tu null thanh ""
                "NOI_BO": "false",
                "NGON_NGU": "VI",
                "GHI_CHU": "",
                "VITRI_KY": "",
                "NOI_NHAN_NN": "",
                "KY_SO": "true",
                "MAU_DU_THAO": "",
                "TINH_TRANG": "",
                "CP_NHANH": "false",
                "VB_ID": "0",
                "KY_HIEU": "null",
                "MA_NGACH": "HC",
                "LOG_XULY": "",
                "DISABLE": "false",
                "MA_DM": "b",
                "MA_DK": "b",
                "ID_DV": "114",
                "ID_LOAI_VB": "140615",
                "ID_NV": "100141",
                "ID_PB": "39597",
                "CAN_CU": "",
                "LOAIVANBANDKY": "",
                "CV_KYSO": "false",
                "CHUCVU_NGUOIKY": "Phó giám đốc",
                "NGUOI_KY": "Phan Sỹ Bách",
                "SO_BAN_PH": "0",
                "DM_PHONG_BAN_PHOI_HOP": [],
                "LDAO_CQ_KYNHAY": [],
                "LDAO_CQ_KYBHANH": [
                    {
                        "ID_PB": "44194",
                        "ID_NV": "110808",
                        "FULLNAME": "Phan Sỹ Bách",
                        "FIRSTNAME": "Phan Sỹ Bách",
                        "LASTNAME": "Bách",
                        "ORDINAL": "0",
                        "TEN_CV": "Phó giám đốc"
                    }
                ],
                "LANH_DAO_PHONGDUYET": [
                    {
                        "ID_PB": "39597",
                        "ID_NV": "100053",
                        "FULLNAME": "Lê Đặng Hiệp Lê",
                        "FIRSTNAME": "Lê Đặng Hiệp",
                        "LASTNAME": "Lê",
                        "ORDINAL": "2"
                    }
                ],
                "DTVB_NOI_NHAN_TN": noinhan_dv(vb['nn_tn']),
                "DTVB_NOI_NHAN_XLY": noinhan_ph(vb['nn_ph']),
                "DTVB_NOI_NHAN_XLY_CTRI": noinhan_ct(vb['nn_ct']),
                "DTVB_NOI_NHAN_XLY_XDB": noinhan_ph(vb['nn_db']),
                "FILEDTVB": [],
                "FILEDTVB_FILECHINH": [
                    {
                        "ID_FILE": files_info[0]['ID_FILE'],
                        "TEN_FILE": files_info[0]['TEN_FILE'],
                        "LOAI_FILE": "docx",
                        "FILE_CHINH": "true",
                        "DUONG_DAN": files_info[0]['DUONG_DAN'],
                        "DISABLE": "false",
                        "NGAY_TAO": files_info[0]['NGAY_TAO'],
                        "TRANG_THAI": "null"
                    },
                    {
                        "ID_FILE": files_info[1]['ID_FILE'],
                        "TEN_FILE": files_info[1]['TEN_FILE'],
                        "LOAI_FILE": "pdf",
                        "FILE_CHINH": "true",
                        "DUONG_DAN": files_info[1]['DUONG_DAN'],
                        "DISABLE": "false",
                        "NGAY_TAO": files_info[1]['NGAY_TAO'],
                        "TRANG_THAI": "null"
                    }
                ],
                "FILEDTVB_GIAI_TRINH": [],
                "FILEDTVB_ROLE": [],
                "FILEDTVB_CANCU": [],
                "CAN_CU_CONGVIEC": [],
                "HOAN_THANH_CONGVIEC": [],
                "FILE_HDTV": [],
                "ID_VB_HDTV": []
            }  

            print(payload2)

            req = Request('POST', url2, headers=headers, json = payload2)
            session = requests.Session()
            prepped = session.prepare_request(req)

            r2 = session.send(prepped)

            print(r2.status_code)
            print(r2.text)
# Chạy chương trình chính

upload_vb("evnsrldc\\hanb", "Minminnguyen@175", lst_vb[0])
# print(noinhan_ph(['TH','KH']))