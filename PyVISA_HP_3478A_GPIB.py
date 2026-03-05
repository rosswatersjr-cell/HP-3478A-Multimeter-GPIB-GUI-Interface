import tkinter as tk
from tkinter import *
from tkinter.scrolledtext import ScrolledText
from tkinter import ttk, StringVar, Menu, IntVar, font, messagebox
from win32api import GetMonitorInfo, MonitorFromPoint
from time import sleep
import pathlib
import os
import pyvisa

version = "2026.03.05"
class HP3478A():
    ########### REQUIREMENTS ##########
    # NI VISA For Windows 11 64 Bit, (2022 Q3) Or Later Installed
    # PyVISA 1.13.0 Or Latest Version. Last Tested Using PyVISA-1.16.2
    # HP 3478A Multimeter Or Equivalent With GPIB Interface
    # Tested With (NI GPIB-USB-HS+) Controller
    # Tested Using Python Version 3.14.2 64 Bit
    # This Class Assumes That Pyvisa And NI-VISA Are Installed And Working Correctly.
    def __init__(self):
        self.rm=pyvisa.ResourceManager()
        self.test_instrument=""
        self.args=[]
        self.controller_port=''
        self.controller_name=''
        self.controller_resources=None
        self.dvm=None
        self.dvm_name=''
        self.dvm_address='6'
        self.dvm_port=None
        self.rtn_val=''
        self.index=None
        self.dvm_commands=[]
        self.dvm_functions=["DC VOLTS","AC VOLTS","2-WIRE OHMS","4-WIRE OHMS","EXTENDED OHMS","DC CURRENT","AC CURRENT",
                            "DC VOLTS PRESET 0","DC VOLTS PRESET 1","AC VOLTS PRESET","2-WIRE OHMS PRESET",
                            "4-WIRE OHMS PRESET","EXTENDED OHMS PRESET","DC CURRENT PRESET","AC CURRENT PRESET"]
        self.real_function=""
        self.dvm_ranges=[]
        self.real_range=""
        self.dvm_digits=["3 1/2 Digits","4 1/2 Digits","5 1/2 Digits"]
        self.real_digit=""
        self.dvm_triggers=["Internal Trigger","External Trigger","Single Trigger","Trigger Hold","Fast Trigger"]
        self.real_trigger=""
        self.dvm_autozero=["Autozero On","Autozero Off"]
        self.real_autozero=""
    def controller_init(self):
        if not Controller_Initialized.get():
            try:
                msg1='GPIB Bus Could Take Up To 30 Seconds To Initialize.\n'
                msg2='You Will Be Prompted When Completed!\n'
                msg3='Press OK Button To Start.'
                msg=msg1+msg2+msg3
                messagebox.showinfo('GPIB Bus Initialization', msg)
                self.controller_resources=self.rm.list_resources() # Time Consuming
                ctrlr=self.rm.open_resource('GPIB::INTFC')
                interface_num=ctrlr.interface_number
                self.controller_name=ctrlr.resource_manufacturer_name
                self.controller_port=str(ctrlr.resource_name)[ctrlr.resource_name.find("'")+1:ctrlr.resource_name.find(':')]
                interface_name.config(text=self.controller_name)
                interface_address.config(text='Controller'+str(interface_num)+ ' @ '+str(ctrlr.resource_name)[ctrlr.resource_name.find("'")+1:ctrlr.resource_name.find(':')])
                self.rtn_val='Initialization Passed'
                interface_stat.config(text=self.rtn_val)
                root.update()
                msg1=self.controller_name+' Controller Initialization Complete!\n'
                if len(self.controller_resources)!=0:msg2='Instruments Detected @ '+str(self.rm.list_resources())
                else:msg2='No Instruments Detected Connected To GPIB Bus!'
                messagebox.showinfo(self.controller_name+' GPIB Bus Initialization', msg1+msg2)
                Controller_Initialized.set(True)
                return(self.rtn_val)
            except Exception as e:
                self.rtn_val='Initialization Failed'
                interface_stat.config(text=self.rtn_val)
                root.update()
                msg1='Exception occurred while code execution:\n'
                msg2=repr(e)+'GPIB Bus Controller Initialization'
                messagebox.showerror('GPIB Controller Initialization',msg1+msg2)
                Controller_Initialized.set(False)
                return 'break'
    def set_local(self):
        self.dvm.control_ren(pyvisa.constants.VI_GPIB_REN_ADDRESS_GTL)# Local Mode
    def dvm_initialize(self):
        if not Controller_Initialized.get():self.controller_init()
        try:
            if len(self.controller_resources)>=1:
                self.dvm_port=str(self.controller_port+'::'+self.dvm_address+'::INSTR')
                if self.dvm_port in self.controller_resources:
                    self.dvm=self.rm.open_resource(self.dvm_port)
                    self.dvm.clear()
                    sleep(2.0)
                    result = self.dvm.query('TST?')
                    if result.strip() == '1':
                        dvm_stat.config(text="Self-Test OK")
                        dvm_address.config(text=f"HP 3478A @ GPIB{self.dvm_address}")
                        msg1=self.dvm_name+' Initialization Complete!\n'
                        msg2='HP 3478A Detected @ '+str(self.dvm_port)
                        messagebox.showinfo(self.dvm_name+' HP 3478A Initialization',msg1+msg2)
                        DVM_Initialized.set(True)
                    else:
                        dvm_stat.config(text="Self-Test Failed!")
                        dvm_address.config(text=f"HP 3478A @ GPIB{self.dvm_address}")
                        msg=f"HP 3478A Self-Test Failed With Error Code: {result.strip()}"
                        messagebox.showerror(self.dvm_name+' HP 3478A Initialization',msg)
                        DVM_Initialized.set(False)
                        self.set_local()
                        return
                    root.update()
                else:
                    self.rtn_val='Initialization Failed'
                    dvm_stat.config(text=self.rtn_val)
                    root.update()
                    msg1='DVM Not Detected!\n'
                    msg2='Please Make Sure That The Instrument Is Powered On\n'
                    msg3='And The GPIB Cable Is Connected To The Instrument.\n'
                    msg4='Then Try Again!'
                    msg=msg1+msg2+msg3+msg4
                    messagebox.showerror('HP 3478A Initialization',msg)
                    return self.rtn_val
            else:    
                    msg1='HP 3478A Initialization Failed!\n'
                    msg2='No DVM Detected Connected To GPIB Bus!'
                    messagebox.showerror('Oscilloscope Initialization', msg1+msg2)
                    DVM_Initialized.set(False)
                    self.rtn_val='Failed'
                    self.set_local()
                    return self.rtn_val
        except Exception as e:
            self.rtn_val='Initialization Failed'
            msg1='Exception occurred while code execution:\n'
            msg2=repr(e)+'GPIB Bus "dvm_initialize()"'
            messagebox.showerror('HP 3478A Initialization',msg1+msg2)
            self.set_local
            return 'break'
    def send_to_dvm(self,command):
        if not Controller_Initialized.get():self.controller_init()
        if not DVM_Initialized.get():self.dvm_initialize()
        query.focus()
        query.config(text="")
        try:
            root.update()
            self.dvm.timeout = Timeout.get()*1000
            self.rtn_val=self.dvm.query(command)
            self.rtn_val=self.rtn_val.strip()
            query.focus()
            query.config(text=f'{Function.get()} = {self.rtn_val}')
            return(self.rtn_val)
        except Exception as e:
            query.focus()
            query.config(text="Measurement Failed")
            title='GPIB Bus Query Failed'
            msg1='Exception occurred while code execution:\n'
            msg2=f'Bus Command = {command}\n'
            msg3=repr(e)+'GPIB Bus'
            messagebox.showerror(title,msg1+msg2+msg3)
            return 'break'
def execute_cmd(event):
    if "PRESET" not in Function.get():
        if Range.get()=="30mV":GPIB.real_range="R-2"
        elif Range.get()=="300mV" or Range.get()=="300mA":GPIB.real_range="R-1"
        elif Range.get()=="3V" or Range.get()=="3A":GPIB.real_range="R0"
        elif Range.get()=="30V" or Range.get()=="3O Ohms":GPIB.real_range="R1"
        elif Range.get()=="300V" or Range.get()=="30O Ohms":GPIB.real_range="R2"
        elif Range.get()=="3K Ohms":GPIB.real_range="R3"
        elif Range.get()=="30K Ohms":GPIB.real_range="R4"
        elif Range.get()=="300K Ohms":GPIB.real_range="R5"
        elif Range.get()=="3M Ohms":GPIB.real_range="R6"
        elif Range.get()=="30M Ohms":GPIB.real_range="R7"
        elif Range.get()=="Autorange":GPIB.real_range="RA"
        if Digit.get()=="3 1/2 Digits":GPIB.real_digit="N3"
        elif Digit.get()=="4 1/2 Digits":GPIB.real_digit="N4"
        elif Digit.get()=="5 1/2 Digits":GPIB.real_digit="N5"
        if Trigger.get()=="Internal Trigger":GPIB.real_trigger="T1"
        elif Trigger.get()=="External Trigger":GPIB.real_trigger="T2"
        elif Trigger.get()=="Single Trigger":GPIB.real_trigger="T3"
        elif Trigger.get()=="Single Hold":GPIB.real_trigger="T4"
        elif Trigger.get()=="Fast Trigger":GPIB.real_trigger="T5"
        if Autozero.get()=="Autozero On":GPIB.real_autozero="Z1"
        elif Autozero.get()=="Autozero Off":GPIB.real_autozero="Z0"
        cmd=str(GPIB.real_function+GPIB.real_range+GPIB.real_autozero+GPIB.real_digit+GPIB.real_trigger)
    else:
        if Function.get()=="DC VOLTS PRESET 0":cmd="H0"
        elif Function.get()=="DC VOLTS PRESET 1":cmd="H1"   
        elif Function.get()=="AC VOLTS PRESET":cmd="H2"    
        elif Function.get()=="DC CURRENT PRESET":cmd="H5"    
        elif Function.get()=="AC CURRENT PRESET":cmd="H6"    
        elif Function.get()=="2-WIRE OHMS PRESET":cmd="H3"    
        elif Function.get()=="4-WIRE OHMS PRESET":cmd="H4"    
        elif Function.get()=="EXTENDED OHMS PRESET":cmd="H7"    
    actual_lbl.config(text=f"HP3478A.query({cmd})")
    val=GPIB.send_to_dvm(cmd)
    return 'break'
def set_funct_choices(event=None):
    GPIB.dvm_ranges.clear()
    if Function.get()=="DC VOLTS" or Function.get()=="AC VOLTS":
        GPIB.dvm_ranges=["Autorange","30mV","300mV","3V","30V","300V"]
        if Function.get()=="DC VOLTS":GPIB.real_function="F1"
        elif Function.get()=="AC VOLTS":GPIB.real_function="F2"
    elif Function.get()=="DC CURRENT" or Function.get()=="AC CURRENT":    
        GPIB.dvm_ranges=["Autorange","300mA","3A"]
        if Function.get()=="DC CURRENT":GPIB.real_function="F5"
        elif Function.get()=="AC CURRENT":GPIB.real_function="F6"
    elif Function.get()=="2-WIRE OHMS" or Function.get()=="4-WIRE OHMS"or Function.get()=="EXTENDED OHMS":
        GPIB.dvm_ranges=["Autorange","30 Ohms","300 Ohms","3K Ohms","30K Ohms","300K Ohms","3M Ohms","30M Ohms"]
        if Function.get()=="2-WIRE OHMS":GPIB.real_function="F3"
        elif Function.get()=="4-WIRE OHMS":GPIB.real_function="F4"
        elif Function.get()=="EXTENDED OHMS":GPIB.real_function="F7"
    elif Function.get()=="DC VOLTS PRESET 0":
        GPIB.dvm_ranges=["Autorange","30mV","300mV","3V","30V","300V"]
        Trigger.set("Trigger Hold")
        Range.set("30mV")
        Digit.set("4 1/2 Digits")
    elif Function.get()=="DC VOLTS PRESET 1" or Function.get()=="AC VOLTS PRESET":    
        GPIB.dvm_ranges=["Autorange","30mV","300mV","3V","30V","300V"]
        Trigger.set("Single Trigger")
        Range.set("30mV")
        Digit.set("4 1/2 Digits")
    elif Function.get()=="DC CURRENT PRESET" or Function.get()=="AC CURRENT PRESET":    
        GPIB.dvm_ranges=["Autorange","300mA","3A"]
        Trigger.set("Single Trigger")
        Range.set("30mA")
        Digit.set("4 1/2 Digits")
    elif Function.get()=="2-WIRE OHMS PRESET" or Function.get()=="4-WIRE OHMS PRESET" or Function.get()=="EXTENDED OHMS PRESET":    
        GPIB.dvm_ranges=["Autorange","30 Ohms","300 Ohms","3K Ohms","30K Ohms","300K Ohms","3M Ohms","30M Ohms"]
        Trigger.set("Single Trigger")
        Range.set("30 Ohms")
        Digit.set("4 1/2 Digits")
    ranges['values']=GPIB.dvm_ranges
    ranges.current(0)
    root.update()
def destroy():# X Icon Was Clicked Or File/Exit
    GPIB.set_local()# Set Instrument To Local Mode
    for widget in root.winfo_children():
        if isinstance(widget,tk.Canvas):widget.destroy()
        else: widget.destroy()
        os._exit(0)
def menu_popup(event):# display the popup menu
   try:
      popup.tk_popup(event.x_root, event.y_root)
   finally:
      popup.grab_release()#Release the grab
def about():
    title='About PyVISA HP 3478A Multimeter.'
    msg1='Creator: Ross Waters\nEmail: RossWatersjr@gmail.com.\n'
    msg2='Requires NI VISA For Windows 11 64 Bit, (2022 Q3) Or Later\n'
    msg3='And PyVISA 1.13.0 Or Latest Version.\n'
    msg4=f'Revision: 1.0 {version}.\n'
    msg5='Tested With (NI GPIB-USB-HS+) Controller.'
    msg=msg1+msg2+msg3+msg4+msg5
    messagebox.showinfo(title, msg)
if __name__ == "__main__":
    root=tk.Tk()
    dir=pathlib.Path(__file__).parent.absolute()
    filename='HP3478A.ico' # Program icon
    ico_path=os.path.join(dir, filename)
    root.iconbitmap(default=ico_path) # root and children
    root.font=font.Font(family='Lucida Sans',size=11,weight='normal',slant='italic')# Menu Font
    title_font=font.Font(family='Lucida Sans',size=12,weight='normal',slant='italic')# Menu Font
    root.title('PyVISA HP 3478A Multimeter GPIB Interface')
    monitor_info=GetMonitorInfo(MonitorFromPoint((0,0)))
    work_area=monitor_info.get("Work")
    monitor_area=monitor_info.get("Monitor")
    taskbar_hgt=(monitor_area[3]-work_area[3])
    root.configure(bg='#d3d3d3')
    root.option_add("*Font",root.font)
    root.protocol("WM_DELETE_WINDOW",destroy)
    root_hgt=((work_area[3]-taskbar_hgt)*0.5)
    root_wid=work_area[2]*0.34
    x=(work_area[2]/2)-(root_wid/2)
    y=(work_area[3]/2.5)-(root_hgt/2.5)
    root.geometry('%dx%d+%d+%d' % (root_wid,root_hgt,x,y,))
    root.update()
    popup=Menu(root, tearoff=0) # PopUp Menu
    popup.add_command(label="About PyVISA Tektronix TDS 210", background='aqua', command=lambda:about())
    root.bind("<Button-3>", menu_popup)
    Selected_Index=IntVar()
    cmd_lbl=tk.Label(root,foreground="navy",background='#d3d3d3',font=title_font,text='Select / Modify Command Arguments')
    cmd_lbl.place(relx=0.01,rely=0.03,relwidth=0.45,relheight=0.05)
    funct_lbl=tk.Label(root,background='#d3d3d3',font=title_font,text='Select Function',anchor="c",justify='c')
    funct_lbl.place(relx=0.01,rely=0.08,relwidth=0.45,relheight=0.05)
    style= ttk.Style()
    style.theme_use('default')
    style.configure("TCombobox",fieldbackground= '#f0ffff',background= '#f0ffff')
    style.configure( 'Horizontal.TScrollbar',width=22 )
    GPIB=HP3478A()
    Function=tk.StringVar()
    functions=ttk.Combobox(root,font=root.font,textvariable=Function)
    functions['state']='normal'
    functions.place(relx=0.014,rely=0.13,relwidth=0.45,relheight=0.05)
    functions.bind("<<ComboboxSelected>>", set_funct_choices)
    functions['values']=GPIB.dvm_functions
    functions.current(0)
    range_lbl=tk.Label(root,background='#d3d3d3',font=title_font,text='Select Range',anchor="c",justify='c')
    range_lbl.place(relx=0.01,rely=0.18,relwidth=0.45,relheight=0.05)
    Range=tk.StringVar()
    ranges=ttk.Combobox(root,font=root.font,textvariable=Range)
    ranges['state']='normal'
    ranges.place(relx=0.014,rely=0.23,relwidth=0.45,relheight=0.05)
    digit_lbl=tk.Label(root,background='#d3d3d3',font=title_font,text='Select Display Digits',anchor="c",justify='c')
    digit_lbl.place(relx=0.01,rely=0.28,relwidth=0.45,relheight=0.05)
    Digit=tk.StringVar()
    digits=ttk.Combobox(root,font=root.font,textvariable=Digit)
    digits['state']='normal'
    digits.place(relx=0.014,rely=0.33,relwidth=0.45,relheight=0.05)
    digits['values']=GPIB.dvm_digits
    digits.current(0)
    trigger_lbl=tk.Label(root,background='#d3d3d3',font=title_font,text='Select Trigger Type',anchor="c",justify='c')
    trigger_lbl.place(relx=0.014,rely=0.38,relwidth=0.45,relheight=0.05)
    Trigger=tk.StringVar()
    triggers=ttk.Combobox(root,font=root.font,textvariable=Trigger)
    triggers['state']='normal'
    triggers.place(relx=0.014,rely=0.43,relwidth=0.45,relheight=0.05)
    triggers['values']=GPIB.dvm_triggers
    triggers.current(0)
    autozero_lbl=tk.Label(root,background='#d3d3d3',font=title_font,text='Select Autozero',anchor="c",justify='c')
    autozero_lbl.place(relx=0.01,rely=0.48,relwidth=0.45,relheight=0.05)
    Autozero=tk.StringVar()
    autozero=ttk.Combobox(root,font=root.font,textvariable=Autozero)
    autozero['state']='normal'
    autozero.place(relx=0.014,rely=0.53,relwidth=0.45,relheight=0.05)
    autozero['values']=GPIB.dvm_autozero
    autozero.current(0)
    Timeout=tk.DoubleVar()
    Timeout.set(10)
    timeout_lbl=tk.Label(root,background='#d3d3d3',font=title_font,text='Time Out In Seconds')
    timeout_lbl.place(relx=0.014,rely=0.58,relwidth=0.45,relheight=0.05)
    time_out=tk.Entry(root,background='#f0ffff',font=title_font,textvariable=Timeout)
    time_out.place(relx=0.014,rely=0.63,relwidth=0.45,relheight=0.05)
    execute_btn=tk.Button(root,text='Execute Command',bg='navy',fg='#ffffff',
            activebackground='#ffffff',borderwidth=5,relief="raised",font=title_font)
    execute_btn.place(relx=0.014,rely=0.7,relwidth=0.45,relheight=0.05)
    execute_btn.bind("<Button-1>",execute_cmd)
    act_lbl=tk.Label(root,background='#d3d3d3',font=title_font,text='Actual Command Sent')
    act_lbl.place(relx=0.014,rely=0.75,relwidth=0.45,relheight=0.05)
    actual_lbl=tk.Label(root,bg='navy',fg='#ffffff',font=title_font,anchor="c",justify='left',text='',borderwidth=2,relief="groove")
    actual_lbl.place(relx=0.014,rely=0.8,relwidth=0.45,relheight=0.05)
    interface_stat_lbl=tk.Label(root,foreground="navy",background='#d3d3d3',font=title_font,text='GPIB Interface Status',anchor="c",justify='c')
    interface_stat_lbl.place(relx=0.48,rely=0.03,relwidth=0.45,relheight=0.05)
    interface_name=tk.Label(root,background='#E5E5E5',font=root.font,text='Interface Name',borderwidth=2,relief="groove")
    interface_name.place(relx=0.48,rely=0.08,relwidth=0.45,relheight=0.05)
    interface_stat=tk.Label(root,background='#E5E5E5',font=root.font,text='Not Initialized',borderwidth=2,relief="groove")
    interface_stat.place(relx=0.48,rely=0.15,relwidth=0.45,relheight=0.05)
    interface_address=tk.Label(root,background='#E5E5E5',font=root.font, text='Interface Address',borderwidth=2,relief="groove")
    interface_address.place(relx=0.48,rely=0.22,relwidth=0.45,relheight=0.05)
    dvm_stat_lbl=tk.Label(root,foreground="navy",background='#d3d3d3',font=title_font,text='HP 3478A Interface Status',anchor="c",justify='c')
    dvm_stat_lbl.place(relx=0.48,rely=0.29,relwidth=0.45,relheight=0.05)
    dvm_name=tk.Label(root,background='#E5E5E5',font=root.font,text='Hewlett Packard',borderwidth=2,relief="groove")
    dvm_name.place(relx=0.48,rely=0.34,relwidth=0.45,relheight=0.05)
    dvm_stat=tk.Label(root,background='#E5E5E5',font=root.font,text='Not Initialized',borderwidth=2,relief="groove")
    dvm_stat.place(relx=0.48,rely=0.41,relwidth=0.45,relheight=0.05)
    dvm_address=tk.Label(root,background='#E5E5E5',font=root.font,text='HP 3478A Address',borderwidth=2,relief="groove",anchor="c",justify='c')
    dvm_address.place(relx=0.48,rely=0.48,relwidth=0.45,relheight=0.05)
    query_lbl=tk.Label(root,background='#d3d3d3',font=title_font, text='HP 3478A Query Value')
    query_lbl.place(relx=0.48,rely=0.55,relwidth=0.5,relheight=0.05)
    query=tk.Label(root,bg='aqua',fg='#000000',font=root.font, text='',borderwidth=2,relief="groove")
    query.place(relx=0.48,rely=0.6,relwidth=0.5,relheight=0.05)
    Controller_Initialized=tk.BooleanVar()
    Controller_Initialized.set(False) # For TDS210 Class
    DVM_Initialized=tk.BooleanVar()
    set_funct_choices()
    DVM_Initialized.set(False) # For HP3478A Class
    GPIB.controller_init()# Initialize Controller
    if Controller_Initialized.get():GPIB.dvm_initialize()# Initialize Instrument
    root.mainloop()    

