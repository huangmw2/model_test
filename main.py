import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.styles import Alignment
import ctypes
import os
import threading
import time
import sys
from tkinter import Toplevel
import serial.tools.list_ports
import serial
import json
from typing import Any, Dict
# 获取运行目录
if getattr(sys, 'frozen', False):  # 检测是否是 PyInstaller 打包后的环境
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

#file_path = r"Data\.."
file_dir = os.path.join(base_path, "Data")

class JsonDataHandler:
    def __init__(self):
        self.user_data_path = os.path.join(file_dir,"user_data.json")
        self.user_data = self.read_json(self.user_data_path)
        #print(f"self.user_data={self.user_data}")
    def read_json(self,file_path: str) -> Dict[str, Any]:
        """读取JSON文件内容"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"文件 {file_path} 不存在")
            return {}
        except json.JSONDecodeError:
            print(f"文件 {file_path} 格式错误")
            return {}
        
    def write_json(self) -> None:
        # 先清空文件内容
        self.clear_json()

        with open(self.user_data_path, 'w', encoding='utf-8') as f:
            json.dump(self.user_data, f, ensure_ascii=False, indent=2)
        print(f"数据已保存至 {self.user_data_path}")

    def clear_json(self):
        open(self.user_data_path, 'w').close()

    def update_data(self, key: str, value: Any) -> None:
        """更新 self.user_data 中的内容，不立即写入文件"""
        self.user_data[key] = value  # 先存入内存
        print(f"数据已更新到内存: {key} = {value}") 

    def find_data(self, key: str) -> Any:
        """查找 self.user_data 中的某个键值"""
        return self.user_data.get(key, None)

    def delete_data(self, key: str) -> None:
        """从 self.user_data 中删除某个键值"""
        if key in self.user_data:
            del self.user_data[key]
            print(f"数据 {key} 已从内存中删除")
        else:
            print(f"键 {key} 不存在")

class USBComm:
    def __init__(self):
        "初始化，检查时否找到dll库"
        self.dll_flag = False
        self.rtn_data = None
        self.open_buffer_usb = None
        self.set_buffer_usb = None
        if self.get_dll_path():
            self.join_dll()
        else :
            pass

    def get_dll_path(self):
        "获取dll库路径，如果不存在返回flase"
        DLL_Name = r"DLL\CsnPrinterLibs.dll"
        
        self.dll_path = os.path.join(base_path, DLL_Name)
        if os.path.exists(self.dll_path):
            self.dll_flag = True
            return True
        else :
            log = "没有找到共享库的路径; 当前路径:{}".format(self.dll_path)
            print(f"log = {log}")
            return False
        
    def join_dll(self):
        "加载USB的接口"
        self.mylib = ctypes.CDLL(self.dll_path)
        #USB 的接口
        self.mylib.Port_EnumUSB.argtypes = [ctypes.c_char_p, ctypes.c_size_t]
        self.mylib.Port_EnumUSB.restype = ctypes.c_size_t
        self.mylib.Port_OpenUSBIO.argtypes = [ctypes.c_char_p]
        self.mylib.Port_OpenUSBIO.restype = ctypes.c_void_p
        self.mylib.Port_SetPort.argtypes = [ctypes.c_void_p]
        self.mylib.Port_SetPort.restype = ctypes.c_bool
        self.mylib.Port_ClosePort.argtypes = [ctypes.c_void_p]
        self.mylib.Pos_SelfTest.restype = ctypes.c_bool
        #self.buffer_usb = ctypes.create_string_buffer(self.buffer_size)
        self.mylib.WriteData.argtypes = [ctypes.POINTER(ctypes.c_ubyte), ctypes.c_size_t, ctypes.c_ulong]
        self.mylib.WriteData.restype = ctypes.c_int
        self.mylib.Read.argtypes = [ctypes.POINTER(ctypes.c_ubyte), ctypes.c_size_t, ctypes.c_ulong]
        self.mylib.Read.restype = ctypes.c_int
        self.mylib.ReadData.argtypes = [ctypes.POINTER(ctypes.c_ubyte), ctypes.c_size_t, ctypes.c_ulong]
        self.mylib.ReadData.restype = ctypes.c_int
        #打印图片
        self.mylib.Pos_ImagePrint.argtypes = [ctypes.c_wchar_p, ctypes.c_int, ctypes.c_int]
        self.mylib.Pos_ImagePrint.restype = ctypes.c_bool 

    def receive_thread(self):
        buf_size = 1024
        buf =  ctypes.create_string_buffer(buf_size)  # 创建字符串缓冲区
        count =  ctypes.c_size_t(buf_size)
        timeout = ctypes.c_ulong(2)  # 超时设置（例如 5000 毫秒）
        buf_ptr = ctypes.cast(buf, ctypes.POINTER(ctypes.c_ubyte))
        self.mylib.Pos_Reset()
        self.mylib.ReadInit()
        result = self.mylib.ReadData(buf_ptr, count, timeout)
        data_str = b''
        if result > 0:
            data_str = buf.raw[:result]
            hex_representation = ' '.join(f'{byte:02X}' for byte in data_str)  # 转换为十六进制字符串
            self.rtn_data = hex_representation
        else:
            print("没有数据返回")  
              

    #打开USB接口
    def open_usb(self):
        buffer_size = 256
        buffer_usb = ctypes.create_string_buffer(buffer_size)
        if not self.dll_flag :
            return False
        ctypes.memset(buffer_usb, 0, buffer_size)
        result_length = self.mylib.Port_EnumUSB(buffer_usb, buffer_size)
        if not result_length:
            messagebox.showerror("错误", "没有找到USB设备")
            return False
           
        self.open_buffer_usb = self.mylib.Port_OpenUSBIO(buffer_usb)  
        if not self.open_buffer_usb :
            messagebox.showerror("错误", "打开USB端口失败")
            return False    
        self.set_buffer_usb = self.mylib.Port_SetPort(self.open_buffer_usb)  
        if not self.set_buffer_usb :
            messagebox.showerror("错误", "设置USB端口失败")
            return False
        return True
    #关闭usb接口
    def close_usb(self):
        if self.set_buffer_usb:
            self.set_buffer_usb = self.mylib.Port_ClosePort(self.open_buffer_usb) 
        return self.rtn_data
    
    def recevice_data(self,data):
        ret = self.open_usb()
        if ret == False:
            return None
        if isinstance(data, str):
            encoded_data = data.encode('cp936')
        else :
            encoded_data = data
        #创建一个接收线程
        read_thread = threading.Thread(target=self.receive_thread)
        read_thread.start()
        time.sleep(0.1)
        #写入状态数据
        buf =  ctypes.create_string_buffer(bytes(encoded_data))  # 创建字符串缓冲区
        count =  ctypes.c_size_t(len(encoded_data))  # 数据长度
        timeout = ctypes.c_ulong(5000)  # 超时设置（例如 5000 毫秒）
        buf_ptr = ctypes.cast(buf, ctypes.POINTER(ctypes.c_ubyte))
        self.rtn_data = None
        #循环个3次
        for i in range(3):
            self.mylib.WriteData(buf_ptr, count, timeout)
            if self.rtn_data:
                break
            time.sleep(0.1)
        #销毁线程
        self.mylib.ReadClose()
        read_thread.join()
        self.close_usb()
        return self.rtn_data
    
    def send_long_data(self,data):
        
        if isinstance(data, str):
            encoded_data = data.encode('cp936')
        else :
            encoded_data = data
            
        buf =  ctypes.create_string_buffer(bytes(encoded_data))  # 创建字符串缓冲区
        count =  ctypes.c_size_t(len(encoded_data))  # 数据长度
        timeout = ctypes.c_ulong(5000)  # 超时设置（例如 5000 毫秒）
        buf_ptr = ctypes.cast(buf, ctypes.POINTER(ctypes.c_ubyte))
        result = self.mylib.WriteData(buf_ptr, count, timeout)
        return result
    
    def send_hex_data(self,data):
        ret = self.open_usb()
        if ret == False:
            return False
        if isinstance(data, str):
            encoded_data = data.encode('cp936')
        else :
            encoded_data = data
        
        buf =  ctypes.create_string_buffer(bytes(encoded_data))  # 创建字符串缓冲区
        count =  ctypes.c_size_t(len(encoded_data))  # 数据长度
        timeout = ctypes.c_ulong(5000)  # 超时设置（例如 5000 毫秒）
        buf_ptr = ctypes.cast(buf, ctypes.POINTER(ctypes.c_ubyte))
        result = self.mylib.WriteData(buf_ptr, count, timeout)   
        if result < 0:
            messagebox.showerror("错误", "USB端口发送数据失败")
        self.close_usb()
        return True
    
    def print_image(self,image_path,image_size):
        ret = self.open_usb()
        if ret == False:
            return False
        ret = self.mylib.Pos_ImagePrint(image_path,image_size,0)
        self.close_usb()
        return ret
class ExcelHandler:
    "处理Excel表格事件"
    def __init__(self):
        self.source_file = os.path.join(file_dir, "原文件.xlsx")
        
        self.search_term = "产品表面"
        self.found_cell = None
        self.ws = None
        self.max_cell = 100

        self.excel_test_list = {}
        try:
            wb = load_workbook(self.source_file)
            self.ws = wb.active  # 获取当前激活的工作表
            self.loading_test_data()
        except Exception as e:
            messagebox.showerror("错误", f"加载 Execel 数据失败：{e}")
        
    def get_excel_topic_data(self,cell_name):
        # 遍历表格查找 "测试项目" 的列索引
        if self.ws == None:
            return None
        col_idx = None
        row_idx = None

        for row in self.ws.iter_rows():
            for cell in row:
                if cell.value == "测试项目":
                    col_idx = cell.column  # 获取列索引（数字）
                    row_idx = cell.row     # 获取行索引
                    break
            if col_idx:
                break

        # 获取该列 "测试项目" 下面的所有数据
        column_data = [self.ws.cell(row=i, column=col_idx).value for i in range(row_idx + 1, self.ws.max_row + 1) 
                       if self.ws.cell(row=i, column=col_idx).value is not None]
        return column_data

    def get_excel_data(self,cell):
        if self.ws :
            cell_value = self.ws[cell].value
        else :
            cell_value = None
        return cell_value  
    def loading_test_data(self):
        step = 1
        if not self.ws:
            return 
        for row in self.ws.iter_rows():
            for cell in row:
                if cell.value == self.search_term:
                    self.found_cell = cell
                    break
            if self.found_cell :
                break

        for i in range(0, self.max_cell):  # 
            next_row = self.found_cell.row + i * step  # 计算下一行的行号
            next_cell = self.ws.cell(row=next_row, column=self.found_cell.column)  # 获取单元格
            if not next_cell.value:
                break
            self.excel_test_list[next_cell.coordinate]=next_cell.value
    def get_topic_range(self,topic_name):
        step = 1
        found_cell = None
        if not self.ws:
            return None,0
        for row in self.ws.iter_rows():
            for cell in row:
                if cell.value == topic_name:
                    found_cell = cell
                    break
            if found_cell :
                break
        if found_cell == None:
            return None,0
        
       # print(f"found_cell={found_cell}")
        for i in range(0, self.max_cell):  # 这里 5 只是示例，你可以改成你需要的次数
            next_row = found_cell.row + i * step  # 计算下一行的行号
            next_cell = self.ws.cell(row=next_row, column=found_cell.column)  # 获取单元格
            if  next_cell.value and i != 0:  #如果不为空就break掉 
                break;    
           # print(f"next_cell.value={next_cell.value}")
        if i >=99:
            return None,0
        # 提取字母部分
        column_letter = found_cell.coordinate[0]
        # 将字母转换为ASCII值并加1
        new_column_letter = chr(ord(column_letter) + 1)
        # 提取数字部分
        row_number = found_cell.coordinate[1:]
        # 返回新的单元格
        cell_coordinate = new_column_letter + row_number
        return cell_coordinate ,i
class EventHandler:
    """事件处理类，专门处理 UI 中的各种事件"""
    def __init__(self):
        self.usb_send_data = USBComm()

    #获取文件数据
    def get_file_data(self,file_name):
        file_path = os.path.join(file_dir,file_name)
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                data = file.read()
                return data  # 读取内容并返回
        except (FileNotFoundError, IOError) as e:
            messagebox.showerror("错误", f"无法读取文件 {file_path}: {e}") 
    #获取产品信息
    def get_driver_product(self,_entry=None):
        byte_data = bytes.fromhex("1D 49 43")
        data_hex_str = self.usb_send_data.recevice_data(byte_data)
        if data_hex_str == None:
            return
        bytes_data = bytes.fromhex(data_hex_str)
        result_str = bytes_data.decode('utf-8')
        _entry.delete(0, tk.END)
        _entry.insert(0,result_str)
        
        #print(f"data={result_str}")
    #测试时间
    def get_now_time(self,_entry):
        current_date = time.strftime("%Y-%m-%d")  # 格式化时间
        _entry.delete(0, tk.END)
        _entry.insert(0, current_date)  # 插入到 Entry
    #样机数量
    def change_machine_num(self,paras,_entry):
        value = _entry.get()
        num = int(value)
        if paras == "add":
            num = num + 1
            _entry.delete(0, tk.END)
            _entry.insert(0, str(num))  # 插入到 Entry
        else :
            if num > 1:
                num = num -1
            _entry.delete(0, tk.END)
            _entry.insert(0, str(num))  # 插入到 Entry      
    #获取版本信息
    def get_version_inf(self,_entry):
        byte_data = bytes.fromhex("1D 49 41")
        data_hex_str = self.usb_send_data.recevice_data(byte_data)
        if data_hex_str == None:
            return 
        bytes_data = bytes.fromhex(data_hex_str)
        result_str = bytes_data.decode('utf-8')
        _entry.delete(0, tk.END)
        _entry.insert(0,result_str)
        #print(f"data={result_str}")
    #切刀测试（发送全切，半切数据）
    def cut_test(self):
        byte_data = bytes.fromhex("1b 40 1b 61  00 1b 24 00  00 1b 4d 00  1b 45 00 1b  2d 00 1b 7b  00 1d 42 00  1b 56 00 1b  39 01 1d 21  00 48 61 6c  66 43 75 74  20 50 61 70  65 72 0a 1b 64 0a 1b 6d 1b 40 1b 61  00 1b 24 00  00 1b 4d 00  1b 45 00 1b  2d 00 1b 7b  00 1d 42 00  1b 56 00 1b  39 01 1d 21  00 46 75 6c  6c 43 75 74  20 50 61 70  65 72 0a  1b 64 0a 1b 69 ")
        self.usb_send_data.send_hex_data(byte_data)
    def feed_paper_test(self):
        byte_data = bytes.fromhex("0a 0a 0a 0a 0a 0a 0a 0a 0a 0a 0a 0a 0a 0a 0a 0a 0a 0a 0a 0a")
        self.usb_send_data.send_hex_data(byte_data)     
    #平均电流（发送打印速度的数据）
    def avg_I_test(self,_radiobutton):
        value = _radiobutton.get()
        print(f"value={value}")
        if value == "2寸":
            data = self.get_file_data("twoinches_speed.hex")
        else :
            data = self.get_file_data("threeinches_speed.hex")
        byte_data = bytes.fromhex(data)
        self.usb_send_data.send_hex_data(byte_data)
    #峰值电流（发送黑块图片数据）
    def black_print(self,_combbox):
        value = _combbox.get()
        print(f"value={value}")
        image_path =  os.path.join(file_dir,"black.bmp")
        self.usb_send_data.print_image(image_path,int(value))
    #usb接口测试
    def usb_comm_test(self):
        data = self.get_file_data("usb_data_test.hex")
        byte_data = bytes.fromhex(data)
        self.usb_send_data.send_hex_data(byte_data)
    #列出串口接口
    def list_serial_com(self):
        ports = serial.tools.list_ports.comports()
        available_ports = [port.device for port in ports]
        return available_ports
    #设置波特率
    def set_baud_rate(self,value):
        #状态查询
        data = "02 00 30 00  00 00 00 00  00 00 00 00  08 00 3a 00  01 02 04 08  10 20 40 80  ff"
        byte_data = bytes.fromhex(data)
        self.usb_send_data.send_hex_data(byte_data)
        time.sleep(0.05)
        if value == 9600:
            data = "02 00 81 00  00 00 00 00  00 00 00 00  01 00 82 00  00 00 "
        else :
            data = "02 00 81 00  00 00 00 00  00 00 00 00  01 00 82 00  04 04 "
        byte_data = bytes.fromhex(data)
        self.usb_send_data.send_hex_data(byte_data)
    #串口打印测试
    def serial_comm_test(self,com,baud_rate,flow_ctrl):
        print("串口接口测试")
        print(f"com={com},baud_rate={baud_rate},flow_ctrl={flow_ctrl}")

    def ttl_comm_test(self):
        print("TTL接口测试")
    def ethernet_comm_test(self,ip,port,para):
        print("网口接口测试")
    def _4g_comm_test(self):
        print("4g接口测试")
    def cashbox_comm_test(self):
        print("钱箱接口测试")
        byte_data = bytes.fromhex("1b 70 00 14 3c")
        ret = self.usb_send_data.send_hex_data(byte_data)
        if ret == False:
            return
        byte_data = bytes.fromhex("1b 70 01 14 3c")
        ret = self.usb_send_data.send_hex_data(byte_data)
    def cut_cmd(self):
        byte_data = bytes.fromhex("1b 69")
        self.usb_send_data.send_hex_data(byte_data)   
    def print_selftest(self):
        byte_data = bytes.fromhex()
        self.usb_send_data.send_hex_data(byte_data)             
    def print_page_mode(self):
        data = self.get_file_data("page_mode.hex")
        byte_data = bytes.fromhex(data)
        self.usb_send_data.send_hex_data(byte_data)  
class MainUI(tk.Tk):
    """
        根据样机出样的测试用例来设计UI
    """
    def __init__(self, event_handler,excel_handler,json_data_handler):
        super().__init__()
        self.event_handler = event_handler  # 接收事件处理类的实例
        self.excel_handler = excel_handler
        self.json_handler = json_data_handler
        self.title("样机出样测试工具")
        self.geometry("800x600+600+300")
        self.temp_configure_name = self.excel_handler.get_excel_topic_data("测试项目")
        self.configure_name =["确认结构和外观测试","确认各接口","确认按键功能","确认指示灯和蜂鸣器","确认侦测传感器","吐纸回收功能"
                              ,"确认打印效果及切刀","纸卷适用确认（票据纸）","确认显示屏","确认扫描头","确认设置项"," "]
        self.create_widgets()
        self.states = ["NA", "OK", "NG"]
        self.current_index = 0  # 初始状态为 "NA"  

    def create_widgets(self):
        # 创建一个主容器框架
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)
        container.config(borderwidth=2, relief="solid")
        # 创建 Canvas 和垂直滚动条
        self.canvas = tk.Canvas(container,  bd=0, highlightthickness=0)
        v_scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=v_scrollbar.set)
        v_scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        # 在 Canvas 内创建一个容器，用于放置所有控件
        self.main_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.main_frame, anchor="nw")

        # 当内部容器尺寸改变时更新 Canvas 的滚动区域
        self.main_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )


        # 绑定鼠标滚轮事件（Windows系统）
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        load_frame = tk.LabelFrame(self.main_frame, text="设置", width=650, height=80)
        load_frame.pack(side=tk.TOP, padx=5, pady=5)  
        #加载数据按钮
        load_data_button = tk.Button(load_frame, text="加载数据", width=10, command=self.load_data_button)
        load_data_button.place(x=10,y=10)
        #保存数据按钮
        save_button = tk.Button(load_frame, text="保存数据", width=10, command=self.save_data_button)
        save_button.place(x=110,y=10)
        #生成样机测试报告
        creat_report_button = tk.Button(load_frame, text="生成样机测试报告", width=14, command=self.creat_report_button)
        creat_report_button.place(x=210,y=10)

        basic_frame = tk.LabelFrame(self.main_frame, text="基本信息", width=650, height=200)
        basic_frame.pack(side=tk.TOP, padx=5, pady=5)  
        frame,self.salesman_combobox = self.configure_basic_info(basic_frame,"业务员",self.excel_handler.get_excel_data("B1") ,"样机版本外发测试员:(默认)",1,['庄思峰','张佰鹏'])
        frame,self.drive_type_entry = self.configure_basic_info(basic_frame,"设备型号",self.excel_handler.get_excel_data("B2"),"设备型号:(默认)",0,None)
        self.configure_button(frame,"设备型号","尝试获取(USB获取)",self.drive_type_entry)
        frame,self.test_time_entry = self.configure_basic_info(basic_frame,"测试时间",self.excel_handler.get_excel_data("B3") ,"测试时间:(默认)",0,None)
        self.configure_button(frame,"测试时间","当前时间",self.test_time_entry)
        frame,self.machine_quantity_entry = self.configure_basic_info(basic_frame,"样机数量",self.excel_handler.get_excel_data("B4"),"样机数量:(默认)",0,"1")
        self.configure_button(frame,"样机数量","↑",self.machine_quantity_entry)
        frame,self.power_source_combobox = self.configure_basic_info(basic_frame,"电源",self.excel_handler.get_excel_data("B5"),"电源:(默认)",1,['24V','12V','9V'])
        frame,self.version_entry = self.configure_basic_info(basic_frame,"样机/版本确认",self.excel_handler.get_excel_data("B6"),"样机/版本确认:(默认)",0,None)
        self.configure_button(frame,"样机/版本确认","尝试获取(USB获取)",self.version_entry) 

        self.items_list = list(self.excel_handler.excel_test_list.items()) 
        #print(self.items_list)
        #结构和外观测试
        self.appearance_test = []
        self.configure_test_component(self.configure_name[0],self.appearance_test)
        #确认各接口
        self.interface_test = []
        self.configure_test_component(self.configure_name[1],self.interface_test)
        #确认按键功能
        self.key_function_test = []
        self.configure_test_component(self.configure_name[2],self.key_function_test)
        #确认指示灯和蜂鸣器
        self.light_beep_test = []
        self.configure_test_component(self.configure_name[3],self.light_beep_test)
        #确认侦测传感器
        self.sensor_test = []
        self.configure_test_component(self.configure_name[4],self.sensor_test)
        #吐纸回收功能
        self.paper_recovery_test = []
        self.configure_test_component(self.configure_name[5],self.paper_recovery_test)
        #确认打印效果及切刀
        self.print_quality_test = []
        self.configure_test_component(self.configure_name[6],self.print_quality_test) 
        #纸卷适用确认（票据纸）
        self.paper_roll_test = []
        self.configure_test_component(self.configure_name[7],self.paper_roll_test)
        #确认显示屏
        self.display_device_test = []
        self.configure_test_component(self.configure_name[8],self.display_device_test)
        #确认扫描头
        self.scanner_test = []
        self.configure_test_component(self.configure_name[9],self.scanner_test)
        #确认设置项
        self.config_set = []
        self.configure_test_component(self.configure_name[10],self.config_set)

    def load_item_data(self,item_list):
        for item in item_list:
            key = item[1]
            data = self.json_handler.find_data(key)
            if data:
                if 'radiobutton' in data:
                    item[2].set(data['radiobutton'])
                if 'entry' in data:
                    item[3].insert(0, data['entry'])
                
    def load_data_button(self):
        data = None
        config = {
            "业务员": (self.excel_handler.get_excel_data("B1"),1,self.salesman_combobox),
            "设备型号": (self.excel_handler.get_excel_data("B2"),0,self.drive_type_entry),
            "测试时间": (self.excel_handler.get_excel_data("B3"),0,self.test_time_entry),
            "样机数量": (self.excel_handler.get_excel_data("B4"),0,self.machine_quantity_entry),
            "电源"    : (self.excel_handler.get_excel_data("B5"),1,self.power_source_combobox),
            "版本确认": (self.excel_handler.get_excel_data("B6"),0,self.version_entry),
        }
        for i, (key, flag, commp) in config.items():
            data = self.json_handler.find_data(key)
            if key and data:
                widget_class = commp.winfo_class()  # 获取组件的类名
                if widget_class == "Entry": 
                    commp.delete(0, tk.END)
                    commp.insert(0,data)
                elif widget_class == "TCombobox":  # 如果是 Combobox 组件
                    commp.set(data)

        #结构和外观测试
        self.load_item_data(self.appearance_test)
        #确认各接口
        self.load_item_data(self.interface_test)
        #确认按键功能
        self.load_item_data(self.key_function_test)
        #确认指示灯和蜂鸣器
        self.load_item_data(self.light_beep_test)
        #确认侦测传感器
        self.load_item_data(self.sensor_test)
        #吐纸回收功能
        self.load_item_data(self.paper_recovery_test)
        #确认打印效果及切刀
        self.load_item_data(self.print_quality_test) 
        #纸卷适用确认（票据纸）
        self.load_item_data(self.paper_roll_test)
        #确认显示屏
        self.load_item_data(self.display_device_test)
        #确认扫描头
        self.load_item_data(self.scanner_test)  

    def save_item_data(self,item_list):
        for item in item_list:
            key = item[1]
            if key:
                radiobutton = item[2].get()
                entry = item[3].get()
                value = {"radiobutton":radiobutton,"entry":entry}
                self.json_handler.update_data(key,value) 

    def save_data_button(self):
        config = {
            "config1" : (self.excel_handler.get_excel_data("B1"),self.salesman_combobox.get()),
            "config2" : (self.excel_handler.get_excel_data("B2"),self.drive_type_entry.get()),
            "config3" : (self.excel_handler.get_excel_data("B3"),self.test_time_entry.get()),
            "config4" : (self.excel_handler.get_excel_data("B4"),self.machine_quantity_entry.get()),
            "config5" : (self.excel_handler.get_excel_data("B5"),self.power_source_combobox.get()),
            "config6" : (self.excel_handler.get_excel_data("B6"),self.version_entry.get()),
        }
        for i, (key, ui_value) in config.items():
            if key and ui_value:
                self.json_handler.update_data(key,ui_value)
        #结构和外观测试
        self.save_item_data(self.appearance_test)
        #确认各接口
        self.save_item_data(self.interface_test)
        #确认按键功能
        self.save_item_data(self.key_function_test)
        #确认指示灯和蜂鸣器
        self.save_item_data(self.light_beep_test)
        #确认侦测传感器
        self.save_item_data(self.sensor_test)
        #吐纸回收功能
        self.save_item_data(self.paper_recovery_test)
        #确认打印效果及切刀
        self.save_item_data(self.print_quality_test) 
        #纸卷适用确认（票据纸）
        self.save_item_data(self.paper_roll_test)
        #确认显示屏
        self.save_item_data(self.display_device_test)
        #确认扫描头
        self.save_item_data(self.scanner_test)  
        self.json_handler.write_json()

    def creat_report_button(self):
        for i in range(len(self.appearance_test)):
            key = self.appearance_test[i][1]
            radiobutton = self.appearance_test[i][2].get()
            entry = self.appearance_test[i][3].get()
            value = {key:{radiobutton,entry}}

    def configure_button(self,frame,frame_name,_text,_commp):
        self.button_func = {
            "设备型号" : self.event_handler.get_driver_product,
            "测试时间": self.event_handler.get_now_time,
            "样机/版本确认":self.event_handler.get_version_inf,
            "样机数量":self.event_handler.change_machine_num
        }
        if frame_name == '样机数量':
            button = tk.Button(frame, text=_text, command=lambda :self.button_func[frame_name]("add",_commp))
            button.pack(side=tk.LEFT, padx=(20,5), pady=5)  
            button = tk.Button(frame, text="↓", command=lambda :self.button_func[frame_name]("dec",_commp))
            button.pack(side=tk.LEFT, padx=(20,5), pady=5)     
        else :
            button = tk.Button(frame, text=_text, command=lambda :self.button_func[frame_name](_commp))
            button.pack(side=tk.LEFT, padx=(20,5), pady=5)  
            
    def configure_basic_info(self, basic_frame,frame_name,_cell_name,default_cell_name,comp_type,default_value):
        frame = tk.LabelFrame(basic_frame, text=frame_name, width=650, height=50)
        frame.pack(side=tk.TOP, padx=5, pady=5)
        frame.pack_propagate(False)  # 禁止 LabelFrame 自动调整大小 
        if _cell_name:
            cell_name = _cell_name
        else :
            cell_name = default_cell_name
        tk.Label(frame, text=cell_name, anchor="w").pack(side=tk.LEFT, padx=5, pady=5) 
        if comp_type == 0:
            entry = tk.Entry(frame, width=30)
            entry.pack(side=tk.LEFT, padx=(20,5), pady=5)
            if default_value:
                entry.insert(0,default_value)
            temp = entry
        elif comp_type == 1:
            combobox = ttk.Combobox(frame, width=8, values=default_value)  # 根据实际情况修改串口号列表
            combobox.pack(side=tk.LEFT, padx=(20,5), pady=5)
            temp = combobox
        else :
            pass
        return frame,temp
    def configure_test_component(self,_name,_list):
        #配置测试组件
        start, range_row = self.excel_handler.get_topic_range(_name)
        #print(f"start={start},range_row={range_row}")
        if range_row != 0:
            frame = tk.LabelFrame(self.main_frame, text=_name, width=500, height=200)
            frame.pack(side=tk.TOP, padx=5, pady=5)  
            frame.pack_propagate(False)  # 禁止 LabelFrame 自动调整大小  
            items_index = 0
            index_flag = False
            for i in range(len(self.items_list)):
                key, value = self.items_list[items_index]
                if start == key:
                    index_flag = True
                    break
                items_index+=1
            if index_flag:
                for i in range(range_row):
                    key, value = self.items_list[items_index]
                    option, entry = self.standard_test_case(frame, value,i) 
                    _list.append((key,value,option,entry))
                    items_index+=1   
    #创建每个测试项的组件
    def standard_test_case(self,frame,label_name,_row):
        # 通用配置
        padx, pady = 5, 1
        grid_config = {"padx": padx, "pady": pady, "sticky": "w"}
        font_style = ("仿宋", 12)
        _column = 0
        # 输入框组件
        entry = tk.Entry(frame, width=20)
        entry.grid(row=_row, column=_column, **grid_config)
        entry.insert(0, label_name)
        _column += 1
        # 颜色映射
        color_map = {
            "NA": "yellow",
            "NG": "red"
        }
        radio_buttons = {}
        # 状态单选按钮
        status_var = tk.StringVar(value="OK")
        default_bg = None
        def update_bg():
            """更新当前行的 Radiobutton 背景颜色"""
            selected_value = status_var.get()
            for value, rb in radio_buttons.items():
                if value == selected_value:
                    # 选中按钮：使用映射中的颜色
                    rb.config(bg=color_map.get(value, default_bg),
                            activebackground=color_map.get(value, default_bg))
                else:
                    # 未选中按钮恢复为默认白色
                    rb.config(bg=default_bg, activebackground=default_bg)

        # 为状态变量添加追踪：当变量改变时自动调用 update_bg
        status_var.trace_add("write", lambda *args: update_bg())
        for text, value in [("NA", "NA"), ("OK", "OK"), ("NG", "NG")]:
            rb = tk.Radiobutton(frame, text=text, variable=status_var, 
                        value=value, font=font_style,  command=update_bg)
            rb.grid(row=_row, column=_column, **grid_config)
            _column += 1
            radio_buttons[value] = rb

        if default_bg == None:
            default_bg =  rb.cget("bg")
        # 描述组件
        tk.Label(frame, text="描述").grid(row=_row, column=_column, **grid_config)
        describe_entry = tk.Entry(frame, width=20)
        describe_entry.grid(row=_row, column=_column+1, **grid_config)
        _column += 2

        # 全局设置按钮（首行显示）
        if _row == 0:
            tk.Button(frame, text="全部设置成NA/OK/NG", 
                    command=lambda: self.set_all_na_ok_ng(frame.cget("text"))
            ).grid(row=_row, column=_column, **grid_config)
        # 按钮配置字典（核心映射关系）
        button_configs = {
            "电源线供电（走纸电流）": ("走纸20行(USB发送)", self.event_handler.feed_paper_test),
            "电源线供电（平均电流）": ("平均电流(USB发送)", lambda: self.select_inches_window(frame)),
            "电源线供电（峰值电流）": ("打印黑块(USB发送)", lambda:self.select_image_window(frame)),
            "切刀": ("切刀测试(USB发送)", self.event_handler.cut_test),
            "USB线": ("USB接口测试", self.event_handler.usb_comm_test),
            "串口线": ("串口接口测试", lambda:self.select_serial_window(frame)),
            "TTL口": ("TTL接口测试", lambda:self.select_serial_window(frame)),
            "网口": ("网口接口测试", lambda:self.select_ethernet_comm(frame)),
            "4G": ("4G测试",self.event_handler._4g_comm_test),
            "钱箱": ("钱箱接口测试", self.event_handler.cashbox_comm_test),
            "轴侦测（开合盖）": ("状态测试",lambda:self.status_check_comm(frame)),
            "黑标": ("黑标测试", lambda:self.black_label_window(frame)),
            "缝标": ("缝标测试", lambda:self.sew_label_window(frame)),
            "堵纸": ("堵纸/拽纸", lambda:self.jam_pull_paper_window(frame)),
            "过温": ("打印黑块", lambda:self.select_image_window(frame)),
            "切刀回到HOME点": ("发送切刀命令", self.event_handler.cut_cmd),
            "驱动打印票据回收": ("吐纸回收设置", lambda:self.spit_recycing_paper_window(frame)),
            "支持指令协议确认": ("打印自测页", self.event_handler.print_selftest),
            "驱动打自检页样张": ("驱动打印设置", lambda:self.driver_print_window(frame)),
            "小票样张": ("票据样张打印", lambda:self.receipt_print_window(frame)),
            "页模式": ("页模式打印", self.event_handler.print_page_mode),
            "打印宽度确认": ("打印宽度设置", lambda:self.width_test_window(frame)),
            "打印速度确认": ("打印速度设置", lambda:self.speed_test_window(frame)),
        }
        # 动态创建专用按钮
        if config := button_configs.get(label_name):
            text, command = config
            tk.Button(frame, text=text, command=command).grid(
                row=_row, column=_column, **grid_config)
            _column += 1
        return status_var,describe_entry
    #创建新窗口
    def creat_new_window(self,parent,_name,_size):
        # 创建模态窗口
        modal = Toplevel(parent)
        modal.title(_name)
        modal.geometry(_size)
        # 获取鼠标位置
        mouse_x = parent.winfo_pointerx()
        mouse_y = parent.winfo_pointery()
        # 禁止调整窗口大小
        modal.resizable(False, False)
        modal.geometry(f"+{mouse_x}+{mouse_y}")  
        return modal
    #锁定新窗口
    def lock_new_window(self,_modal,_parent):
        _modal.transient(_parent)  # 让对话框依附于父窗口
        _modal.grab_set()  # 模态窗口，锁定主窗口
        _parent.wait_window(_modal)  # 等待窗口关闭
    #选择尺寸窗口
    def select_inches_window(self,parent):
        #创建窗口
        modal = self.creat_new_window(parent,"尺寸选择","150x150",)
        #创建组件
        inches_option = tk.StringVar(value="3寸")
        for text in ("2寸", "3寸"):
            tk.Radiobutton(modal, text=text, variable=inches_option, value=text, font=("仿宋", 12)).pack(padx=5, pady=5)
        tk.Button(modal, text="确认", command=lambda: self.event_handler.avg_I_test(inches_option)).pack(padx=5, pady=5)
        #锁定父窗口
        self.lock_new_window(modal,parent)
    #选择图片大小窗口
    def select_image_window(self,parent):
        modal = self.creat_new_window(parent, "图片选择", "150x150")
        # 创建顶部容器
        top_frame = tk.Frame(modal)
        top_frame.pack(fill="x")
        tk.Label(top_frame, text="图片大小", anchor="w").pack(side="left", padx=5, pady=5)
        image_size = ttk.Combobox(top_frame, width=8, values=['384','576','832'])
        image_size.pack(side="left", padx=5, pady=5)
        image_size.set('576')
        # 独立打包按钮
        tk.Button(modal, text="打印", width=12,command=lambda: self.event_handler.black_print(image_size)).pack(fill="x", padx=5, pady=5)
        self.lock_new_window(modal, parent)
   #选择串口参数窗口
    def select_serial_window(self,parent):
        #创建窗口
        modal = self.creat_new_window(parent,"串口通信测试","150x200",)
        configs = [
            ("端口号", self.event_handler.list_serial_com(), 0),
            ("波特率", ['9600','115200'], '115200'),
            ("流量控制", ['NONE','DstDtr','CtsRts','Xon/Xoff'], 'CtsRts')
        ]
        entries = []
        for i, (text, values, default) in enumerate(configs):
            tk.Label(modal, text=text, anchor="w").grid(row=i,column=0,padx=2,pady=2,sticky="w")
            cb = ttk.Combobox(modal, width=8, values=values)
            cb.grid(row=i,column=0,padx=60,pady=2,sticky="w")
            if values: cb.set(values if text=="端口号" else default)
            entries.append(cb)
        cmd = [
            lambda:self.event_handler.serial_comm_test(entries[0].get(),entries[1].get(),entries[2].get()),
            lambda:self.event_handler.set_baud_rate(entries[1].get())
        ]
        tk.Button(modal, text="打印", width=18, command=cmd[0]).grid(row=i+1,column=0,padx=5,pady=2,sticky="w")
        tk.Button(modal, text="设置波特率(USB设置)",width=18,command=cmd[1]).grid(row=i+2,column=0,padx=5,pady=2,sticky="w")
        #锁定父窗口
        self.lock_new_window(modal,parent)

    #选择网口参数窗口
    def select_ethernet_comm(self,parent):
        modal = self.creat_new_window(parent,"网口测试","180x150")
        fields = [
            {"text": "ip地址:",  "default": "192.168.0.87"},{"text": "端口号:",  "default": "9100"}
        ]
        entries = []
        for idx, field in enumerate(fields):
            tk.Label(modal, text=field["text"], anchor="w").grid(row=idx,column=0,padx=2,pady=2)
            entry = tk.Entry(modal, width=12)
            entry.grid(row=idx,column=1,padx=2,pady=2)
            entry.insert(0, field["default"])
            entries.append(entry)
        # 按钮配置数据 
        func = self.event_handler.ethernet_comm_test
        tk.Button(modal, text="打印", width=8, command=lambda:func(entries[0],entries[1],0)).grid(row=idx+1,column=0,padx=2,pady=2)
        tk.Button(modal, text="设置Ip地址", width=8, command=lambda:func(entries[0],entries[1],1)).grid(row=idx+1,column=1,padx=2,pady=2)
        tk.Button(modal, text="开启DHCP", width=8, command=lambda:func(entries[0],entries[1],2)).grid(row=idx+2,column=0,padx=2,pady=2)
        tk.Button(modal, text="关闭DHCP", width=8, command=lambda:func(entries[0],entries[1],3)).grid(row=idx+2,column=1,padx=2,pady=2)
        #锁定父窗口
        self.lock_new_window(modal,parent)
    #状态检测窗口
    def status_check_comm(self,parent):
        # 创建模态窗口
        modal = Toplevel(parent)
        modal.title("状态检测")
        modal.geometry("200x200")
        # 获取鼠标位置
        mouse_x = parent.winfo_pointerx()
        mouse_y = parent.winfo_pointery()
        # 禁止调整窗口大小
        modal.resizable(False, False)
        modal.geometry(f"+{mouse_x}+{mouse_y}")   

        statu_value_label = tk.Label(modal, text="返回状态值:", anchor="w")
        statu_value_label.place(x=10,y=10)
        statu_value_entry = tk.Entry(modal, width=12) 
        statu_value_entry.place(x=80,y=10)
        status_button = [None]*7
        status_button[6] = tk.Button(modal, text="正常", bg="green",width=12,height=2)
        status_button[6].place(x=50,y=40) 
        
        status_button[0] = tk.Button(modal, text="发送10 04 01", width=12)
        status_button[0].place(x=0,y=100) 
        status_button[1] = tk.Button(modal, text="发送10 04 02", width=12)
        status_button[1].place(x=105,y=100) 
        status_button[2] = tk.Button(modal, text="发送10 04 03", width=12)
        status_button[2].place(x=0,y=130) 
        status_button[3] = tk.Button(modal, text="发送10 04 04", width=12)
        status_button[3].place(x=105,y=130) 
        status_button[4] = tk.Button(modal, text="发送10 04 05", width=12)
        status_button[4].place(x=0,y=160) 
        status_button[5] = tk.Button(modal, text="发送全部", width=12)
        status_button[5].place(x=105,y=160) 
        #statu_value_entry.insert(0,"9100")

        #锁定父窗口
        modal.transient(parent)  # 让对话框依附于父窗口
        modal.grab_set()  # 模态窗口，锁定主窗口
        parent.wait_window(modal)  # 等待窗口关闭    
    #黑标测试窗口
    def black_label_window(self,parent):
        # 创建模态窗口
        modal = Toplevel(parent)
        modal.title("黑标测试")
        modal.geometry("250x160")
        # 获取鼠标位置
        mouse_x = parent.winfo_pointerx()
        mouse_y = parent.winfo_pointery()
        # 禁止调整窗口大小
        modal.resizable(False, False)
        modal.geometry(f"+{mouse_x}+{mouse_y}")   
        black_set_button = [None] * 10 
        black_set_button[0] = tk.Button(modal, text="设置黑标", width=8)
        black_set_button[0].grid(row=0,column=0,pady=5)
        black_set_button[1] = tk.Button(modal, text="取消黑标", width=8)
        black_set_button[1].grid(row=0,column=1,pady=5)
        black_distance_label = tk.Label(modal, text="查找黑标后进纸:", anchor="w")
        black_distance_label.grid(row=1,column=0,pady=5)
        black_distance_entry = tk.Entry(modal, width=6) 
        black_distance_entry.grid(row=1,column=1,pady=5) 
        black_set_button[2] = tk.Button(modal, text="设置", width=6)
        black_set_button[2].grid(row=1,column=2,pady=5) 

        black_set_button[3] = tk.Button(modal, text="查找黑标", width=8)
        black_set_button[3].grid(row=2,column=0,pady=5) 
        black_set_button[4] = tk.Button(modal, text="查找黑标后切纸", width=12)
        black_set_button[4].grid(row=2,column=1,pady=5)  
        black_set_button[5] = tk.Button(modal, text="黑标打印测试", width=12)
        black_set_button[5].grid(row=3,column=0,pady=5)  
        #锁定父窗口
        modal.transient(parent)  # 让对话框依附于父窗口
        modal.grab_set()  # 模态窗口，锁定主窗口
        parent.wait_window(modal)  # 等待窗口关闭    
    #缝标测试窗口
    def sew_label_window(self,parent):
        # 创建模态窗口
        modal = Toplevel(parent)
        modal.title("黑标测试")
        modal.geometry("250x160")
        # 获取鼠标位置
        mouse_x = parent.winfo_pointerx()
        mouse_y = parent.winfo_pointery()
        # 禁止调整窗口大小
        modal.resizable(False, False)
        modal.geometry(f"+{mouse_x}+{mouse_y}")   
        sew_set_button = [None] * 10 
        sew_set_button[0] = tk.Button(modal, text="设置缝标", width=8)
        sew_set_button[0].grid(row=0,column=0,pady=5)
        sew_set_button[1] = tk.Button(modal, text="取消缝标", width=8)
        sew_set_button[1].grid(row=0,column=1,pady=5)
        sew_set_button[2] = tk.Button(modal, text="查找缝标", width=8)
        sew_set_button[2].grid(row=1,column=0,pady=5) 
        sew_set_button[3] = tk.Button(modal, text="查找缝标并切纸", width=12)
        sew_set_button[3].grid(row=1,column=1,pady=5) 

        #锁定父窗口
        modal.transient(parent)  # 让对话框依附于父窗口
        modal.grab_set()  # 模态窗口，锁定主窗口
        parent.wait_window(modal)  # 等待窗口关闭    
    #堵纸/拽纸
    def jam_pull_paper_window(self,parent):
        modal = self.creat_new_window(parent,"堵纸/拽纸测试","150x150")
        #锁定父窗口
        self.lock_new_window(modal,parent)
    #吐纸、回收
    def spit_recycing_paper_window(self,parent):
        modal = self.creat_new_window(parent,"吐纸、回收测试","150x150")
        #锁定父窗口
        self.lock_new_window(modal,parent)
    #驱动打印测试
    def driver_print_window(self,parent):
        modal = self.creat_new_window(parent,"驱动打印测试","150x150")
        #锁定父窗口
        self.lock_new_window(modal,parent)
    #驱动打印测试
    def receipt_print_window(self,parent):
        modal = self.creat_new_window(parent,"票据打印测试","150x150")
        #锁定父窗口
        self.lock_new_window(modal,parent)
    #打印宽度测试
    def width_test_window(self,parent):
        modal = self.creat_new_window(parent,"打印宽度测试","150x150")
        #锁定父窗口
        self.lock_new_window(modal,parent)    
    #打印速度测试
    def speed_test_window(self,parent):
        modal = self.creat_new_window(parent,"打印速度测试","150x150")
        #锁定父窗口
        self.lock_new_window(modal,parent)  
    #设置all窗口
    def set_all_na_ok_ng(self,name): 
        _list_radiobutton= {
            0: self.appearance_test,
            1: self.interface_test,
            2: self.key_function_test,
            3: self.light_beep_test,
            4: self.sensor_test,
            5: self.paper_recovery_test,
            6: self.print_quality_test,
            7: self.paper_roll_test,
            8: self.display_device_test,
            9: self.scanner_test,
        }
        index = self.configure_name.index(name)
        current_index = (self.current_index) % len(self.states)
        self.current_index = self.current_index + 1
        ret_list = _list_radiobutton.get(index, None)
        if ret_list :
            for i in range(len(ret_list)):
                ret_list[i][2].set(self.states[current_index])   
         

    def _on_mousewheel(self, event):
        # Windows 系统下，event.delta 通常为 120 的倍数
        self.canvas.yview_scroll(-1 * (event.delta // 120), "units")

if __name__ == "__main__":
    # 创建事件处理器实例
    event_handler = EventHandler()
    # 创建Excel处理
    excel_handler = ExcelHandler()
    # 加载json数据
    json_data_handler = JsonDataHandler()
    # 创建 UI 实例，并将事件处理器传递给 UI
    app = MainUI(event_handler,excel_handler,json_data_handler)
    app.mainloop()
