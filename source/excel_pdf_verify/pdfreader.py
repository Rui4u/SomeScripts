import os
import pdfplumber
import re
import shutil
import os

import pandas as pd
from pypinyin import pinyin, Style

def chinese_to_pinyin(text):
  if text == "陕西":
    return "SHAANXI"
  if text == "内蒙古":
    return "NEIMONGOL"
    # 将中文转换为拼音
  pinyin_list = pinyin(text, style=Style.NORMAL)
    # 将拼音列表转换为大写并连接成字符串
  return ''.join([x[0].upper() for x in pinyin_list])


from datetime import datetime

def format_date(date_str):
    # 解析日期字符串为 datetime 对象
    date_obj = datetime.strptime(date_str, "%Y-%m-%d")
    # 将日期对象格式化为指定格式
    formatted_date = date_obj.strftime("%d/%m/%Y")
    return formatted_date


current_dir = current_dir = os.path.dirname(os.path.abspath(__file__)) 

excel_file = "东方红.xlsx"
folder_path = "orginal_files"
newFloder_path = 'grouped_files'

excel_file = os.path.join(current_dir, excel_file)
folder_path = os.path.join(current_dir, folder_path)
newFloder_path = os.path.join(current_dir, newFloder_path)


df = pd.read_excel(excel_file, usecols=["收录编号","出生地","护照号", "出发口岸","英文拼音","性别","出生日期","有效期"])

def ensure_string(input_var):
    if not isinstance(input_var, str):
        input_var = str(input_var)
    return input_var

def checkNull(values):
  value = "未知"
  if len(values) > 0:
    value = values[0]
  return ensure_string(value)
  
def getDir(passport_number):
  # 读取Excel文件
  filtered_data = df[df["护照号"] == passport_number]
  departure_ports = checkNull(filtered_data["出发口岸"].values)
  if departure_ports == '未知':
     return {
    "port":"未知",
    "number":"未知",
    "name":"未知",
    "sex":"未知",
    "place":"未知",
    "birthday":"未知",
    "ExpiryDate":"未知"
  } 
  number = checkNull(filtered_data["收录编号"].values)

  name = checkNull(filtered_data["英文拼音"].values).replace("/", " ")
  sex = checkNull(filtered_data["性别"].values)
  if sex == "F":
    sex = "FEMALE"
  if sex == "M":
    sex = "MALE"
  birth_place = chinese_to_pinyin(checkNull(filtered_data["出生地"].values))
  date_str = checkNull(filtered_data["出生日期"].values)
  brith_date = format_date(date_str.split("T")[0])


  voidDate = checkNull(filtered_data["有效期"].values)
  voidDate = format_date(voidDate.split("T")[0])
  
 
  return {
    "port":departure_ports,
    "number":number,
    "name":name,
    "sex":sex,
    "place":birth_place,
    "birthday":brith_date,
    "ExpiryDate":voidDate
  }

def copyToFile(source_file_path, destination_folder, new_file_name):
  # 如果目标文件夹不存在，则创建它
  if not os.path.exists(destination_folder):
      os.makedirs(destination_folder, exist_ok=True)

  destination_file_path = os.path.join(destination_folder, new_file_name)
  if not os.path.exists(destination_file_path):
    shutil.copy2(source_file_path, destination_file_path)
    # print("文件已复制到目标文件夹并改名为:", destination_file_path)
  # else:
  #   print(f"目标文件 '{destination_file_path}' 已存在，跳过复制。")
  

def find_name(text):
  # 定义正则表达式模式
  pattern = r"Name : (\w+ \w+).*?Reference No. : (\w+)"

  # 使用正则表达式进行匹配
  match = re.search(pattern, text)

  # 如果找到匹配项，提取名字和参考号
  if match:
      name = match.group(1)
      # print(f"name:{name}")
      return (True,name)
  else:
      return (False, "")

def find_sex(text):
  # 定义正则表达式模式
  pattern = r"Sex : (\w+).*?Visa No\. : (\w+)"

  # 使用正则表达式进行匹配
  match = re.search(pattern, text)

  # 如果找到匹配项，提取名字和参考号
  if match:
      name = match.group(1)
      # print(f"sex:{name}")
      return (True,name)
  else:
      return (False, "")

def find_place(text):
  # 定义正则表达式模式
  pattern = r'Place of Birth : (.*?) Payment Receipt No'

  # 使用正则表达式进行匹配
  match = re.search(pattern, text)

  # 如果找到匹配项，提取名字和参考号
  if match:
      name = match.group(1).strip()
      return (True,name)
  else:
      return (False, "")

def find_birthday(text):
  # 定义正则表达式模式
  pattern = r"Date of Birth : (\d{2}/\d{2}/\d{4})"
  # 使用正则表达式进行匹配
  match = re.search(pattern, text)

  # 如果找到匹配项，提取名字和参考号
  if match:
      name = match.group(1)
      # print(f"birthday:{name}")
      return (True,name)
  else:
      return (False, "")

def find_expiry_date(text):
  # 定义正则表达式模式
  pattern = r"Passport Expiry Date : (\d{2}/\d{2}/\d{4})"
  # 使用正则表达式进行匹配
  match = re.search(pattern, text)

  # 如果找到匹配项，提取名字和参考号
  if match:
      name = match.group(1)
      # print(f"birthday:{name}")
      return (True,name)
  else:
      return (False, "") 

def find_nationality(text):
  # 定义正则表达式模式
  pattern = r"Nationality : (\w+) Date of issue : (\d{2}/\d{2}/\d{4})"
  # 使用正则表达式进行匹配
  match = re.search(pattern, text)

  # 如果找到匹配项，提取名字和参考号
  if match:
      name = match.group(1)
      # print(f"ExpiryDate:{name}")
      return (True,name)
  else:
      return (False, "")
# 定义一个函数来处理单个PDF文件
def process_pdf(file_path):
    with pdfplumber.open(file_path) as pdf:
        # 遍历每一页
      index = 0
      info_dict = {}
      text = ""
      for page in pdf.pages:
          # 提取文本
          text += page.extract_text() + "\n"
      text_array = [line.strip() for line in text.split('\n') if line.strip()]
      index = 0
      for page in text_array:
          if index == 4:
            info_dict["name"] = page
          else:
            passport_number_match = re.search(r'Passport No. : (\w+)', page)
            
            if passport_number_match:
              passport_number = passport_number_match.group(1)
              info_dict["passport_number"] = passport_number

            findName = find_name(page)
            if findName[0]:
              info_dict["name"] = findName[1]
            findSex = find_sex(page)
            if findSex[0]:
              info_dict["sex"] = findSex[1]

            findBirthday = find_birthday(page)
            if findBirthday[0]:
              info_dict["birthday"] = findBirthday[1]


            findPlace = find_place(page)
            if findPlace[0]:
              info_dict["place"] = findPlace[1].upper()


            findNationality = find_nationality(page)
            if findNationality[0]:
              info_dict["nationality"] = findNationality[1]

            findExpiryDate = find_expiry_date(page)
            if findExpiryDate[0]:
              info_dict["ExpiryDate"] = findExpiryDate[1]
              
              
          index = index + 1
      passport_number = info_dict["passport_number"]
      excel_info = getDir(passport_number)
      number = excel_info["number"]
      port = os.path.join(newFloder_path, excel_info["port"])
      if excel_info["port"] == "未知":
          print("无出发地" + excel_info["port"] + "  名字:" + info_dict["name"])
          print("-------------------------")

      if not excel_info["port"] == "未知" and not (excel_info["name"] == info_dict["name"] and excel_info["sex"] == info_dict["sex"] and excel_info["birthday"] == info_dict["birthday"] and   compair_place(excel_info["place"],info_dict["place"]) and excel_info["ExpiryDate"] == info_dict["ExpiryDate"] and excel_info["name"] == info_dict["name"]):
        print(excel_info["number"])
        print("出发口岸:" + excel_info["port"] + "  名字:" + excel_info["name"])

        if excel_info["name"] != info_dict["name"]:
          print("名字错了:" +  excel_info["name"],info_dict["name"])
        if excel_info["sex"] != info_dict["sex"]:
          print("性别错了:" + excel_info["sex"],info_dict["sex"])
        if excel_info["birthday"] != info_dict["birthday"]:
          print("出生日期错了:" + excel_info["birthday"],info_dict["birthday"])
        if excel_info["place"] != info_dict["place"]:
          print("出生地错了:" + excel_info["place"],info_dict["place"])
        if excel_info["ExpiryDate"] != info_dict["ExpiryDate"]:
          print("有效期错了:" + excel_info["ExpiryDate"],info_dict["ExpiryDate"])
        print("-------------------------")
      copyToFile(file_path, port ,f"{number}" + "_" + info_dict["name"] + ".pdf")
    pdf.close()     

# 定义一个函数来遍历文件夹内的所有PDF文件
def process_pdf_folder(folder_path):
    # 遍历文件夹中的每个文件
    for root, dirs, files in os.walk(folder_path):
        for file_name in files:
            if file_name.endswith('.pdf'):  # 确保是PDF文件
                file_path = os.path.join(root, file_name)
                process_pdf(file_path)

def compair_place(str1, str2):
   str2_temp = str2.replace(" ", "")
   return  str1 == str2 or str1 == str2_temp

# 调用函数并传入文件夹路径
process_pdf_folder(folder_path)
