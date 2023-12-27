import streamlit as st
from openpyxl import load_workbook
from datetime import datetime, timedelta
import glob, os

def delete_all_files_in_directory(directory):
    try:
        # 获取目录下的所有文件和子目录
        files = os.listdir(directory)
        
        # 遍历所有文件和子目录
        for file in files:
            # 构建文件的完整路径
            file_path = os.path.join(directory, file)
            
            # 判断是否是文件
            if os.path.isfile(file_path):
                # 如果是文件，删除之
                os.remove(file_path)
                print(f"Deleted file: {file_path}")
                
            # 如果是目录，递归调用函数删除其中的文件
            elif os.path.isdir(file_path):
                delete_all_files_in_directory(file_path)
        
        print(f"All files in {directory} deleted successfully.")
    
    except Exception as e:
        print(f"An error occurred: {e}")

datetime_format = "%Y-%m-%d %H:%M:%S"  # 对应的格式字符串
target = "output"

# Streamlit页面标题
st.title("Excel 文件处理")

# 接收用户上传的Excel文件
uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx"])


if uploaded_file is not None:

    if not os.path.exists( target): os.makedirs( target)
    else:  delete_all_files_in_directory( target)
    
    workbook = load_workbook( uploaded_file)
    sheet = workbook.active
    x, y, z = None, None, None
    for row in sheet:
        if x is None:
            for i, cell in enumerate( row):
                if cell.value == '入场时间': x = i 
                elif cell.value == r'停车时长/分': y = i
                elif cell.value == '出场时间': z = i
        else:
            try:
                one = row[x].value; two = row[y].value
                if not one: continue
                new_datetime = datetime.strptime( one, datetime_format)  + timedelta(minutes= int( two))
                row[z].value = str( new_datetime)
            except Exception:
                pass

    # file_name = uploaded_file.split( os.sep )[-1]
    datetime_format_ = "%Y%m%d_%H%M%S"  # 对应的格式字符串
    file_name = f"结果文件_{datetime.now().strftime( datetime_format_)}.xlsx"
    file_path =  os.path.join( target, file_name) 
    workbook.save( file_path) 

    # 读取上传的Excel文件为DataFrame

    # 提供处理后的Excel文件下载选项
    # st.markdown(f"[:floppy_disk: 下载处理后的文件]({ file_path})", unsafe_allow_html=True)
    # if st.button("下载处理后的文件"):
    st.download_button(
        label="下载处理后的文件",
        data=open(file_path, "rb").read(),
        file_name=file_name,
        key="download_button"
    )
    st.divider()
