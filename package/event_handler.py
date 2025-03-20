import os
import time
import serial
import tkinter as tk
import serial.tools.list_ports
from tkinter import  messagebox
from package.communication import USBComm

class EventHandler:
    """事件处理类，专门处理 UI 中的各种事件"""
    def __init__(self,base_path,file_dir):
        self.file_dir = file_dir
        self.usb_send_data = USBComm(base_path)

    #获取文件数据
    def get_file_data(self,file_name):
        file_path = os.path.join(self.file_dir,file_name)
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
        image_path =  os.path.join(self.file_dir,"black.bmp")
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