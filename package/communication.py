import ctypes
import os
from package.logger import log_message
from tkinter import  messagebox
import threading
import time

class USBComm:
    def __init__(self,base_path):
        self.dll_path = None
        self.dll_name = r"DLL\CsnPrinterLibs.dll"
        self.base_path = base_path
        self.is_dll_loaded = False
        self.return_data = None
        self.usb_open_buffer = None
        self.usb_set_buffer = None
        
        if self.find_dll_path():
            self.init_dll()
        else :
            pass
    # 查找动态库路径
    def find_dll_path(self):
        self.dll_path = os.path.join(self.base_path, self.dll_name)
        if os.path.exists(self.dll_path):
            self.is_dll_loaded = True
            return True
        else :
            log = "没有找到共享库的路径; 当前路径:{}".format(self.dll_path)
            log_message(log,"ERROR")
            return False
        
    # 初始化动态库的接口
    def init_dll(self):
        self.mylib = ctypes.CDLL(self.dll_path)
        self.mylib.Port_EnumUSB.argtypes = [ctypes.c_char_p, ctypes.c_size_t]
        self.mylib.Port_EnumUSB.restype = ctypes.c_size_t
        self.mylib.Port_OpenUSBIO.argtypes = [ctypes.c_char_p]
        self.mylib.Port_OpenUSBIO.restype = ctypes.c_void_p
        self.mylib.Port_SetPort.argtypes = [ctypes.c_void_p]
        self.mylib.Port_SetPort.restype = ctypes.c_bool
        self.mylib.Port_ClosePort.argtypes = [ctypes.c_void_p]
        self.mylib.Pos_SelfTest.restype = ctypes.c_bool
        self.mylib.WriteData.argtypes = [ctypes.POINTER(ctypes.c_ubyte), ctypes.c_size_t, ctypes.c_ulong]
        self.mylib.WriteData.restype = ctypes.c_int
        self.mylib.Read.argtypes = [ctypes.POINTER(ctypes.c_ubyte), ctypes.c_size_t, ctypes.c_ulong]
        self.mylib.Read.restype = ctypes.c_int
        self.mylib.ReadData.argtypes = [ctypes.POINTER(ctypes.c_ubyte), ctypes.c_size_t, ctypes.c_ulong]
        self.mylib.ReadData.restype = ctypes.c_int
        self.mylib.Pos_ImagePrint.argtypes = [ctypes.c_wchar_p, ctypes.c_int, ctypes.c_int]
        self.mylib.Pos_ImagePrint.restype = ctypes.c_bool 

    #接收数据的线程
    #将接收到的数据转成16进制字符串
    #如果接收到数据，则将数据赋值给self.return_data
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
            self.return_data = hex_representation
        else:
            self.return_data = None
            log = "没有接收到返回的数据"
            log_message(log,"DEBUG")
    # 打开usb通信端口
    # 成功返回TRUE,失败返回FALSE  
    def open_usb(self):
        buffer_size = 256
        buffer_usb = ctypes.create_string_buffer(buffer_size)
        if not self.is_dll_loaded :
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
    #关闭usb通信接口
    def close_usb(self):
        if self.set_buffer_usb:
            self.set_buffer_usb = self.mylib.Port_ClosePort(self.open_buffer_usb) 
        return self.return_data
    #发送指令，并接收返回的结果
    def recevice_data(self,data):
        ret = self.open_usb()
        if ret == False:
            return None
        if isinstance(data, str):
            encoded_data = data.encode('cp936')
        else :
            encoded_data = data
        read_thread = threading.Thread(target=self.receive_thread)
        read_thread.start()
        time.sleep(0.1)
        buf =  ctypes.create_string_buffer(bytes(encoded_data))  # 创建字符串缓冲区
        count =  ctypes.c_size_t(len(encoded_data))  # 数据长度
        timeout = ctypes.c_ulong(5000)  # 超时设置（例如 5000 毫秒）
        buf_ptr = ctypes.cast(buf, ctypes.POINTER(ctypes.c_ubyte))
        self.rtn_data = None
        for i in range(3):
            self.mylib.WriteData(buf_ptr, count, timeout) #写入状态数据
            if self.rtn_data:
                break
            time.sleep(0.1)
        self.mylib.ReadClose()
        read_thread.join()
        self.close_usb()
        return self.return_data
    #发送16进制数据
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
    #打印图片
    def print_image(self,image_path,image_size):
        ret = self.open_usb()
        if ret == False:
            return False
        ret = self.mylib.Pos_ImagePrint(image_path,image_size,0)
        self.close_usb()
        return ret