#!/usr/bin/python
# -*- coding: utf-8 -*-
import sys
import time
import tkMessageBox
import wmi
from Tkinter import *

MODE_NORMAL = "normal"
MODE_DEBUG = "debug"
mod = MODE_NORMAL

# service/package status
STOP = 'Stopped'
RUN = 'Running'
PAUSE = 'Paused'
PENDING = 'Pending'
COMPLEX = 'Complex'
LOADING = 'Loading'
CFG_ERR = 'Config Error'

# fb is foreground color, and bg means background color
STATUS_COLOR = {STOP: {'fg': 'red'}, RUN: {'fg': 'dark green'}, PAUSE: {'fg': 'orange'}, PENDING: {'fg': 'green'},
                COMPLEX: {'fg': 'brown'}, LOADING: {'fg': 'white', 'bg': 'black'},
                CFG_ERR: {'fg': 'white', 'bg': 'red'}}
# default color settings
DEFAULT_FG = 'black'
DEFAULT_BG = None

# system messages
SYS_INFO = 'System Info'
SYS_WARN = 'System Warning'
ERR_MSG = 'Ouch...there are some problems. Perhaps some services are pending, ' \
          'please reload the service status later and check.'
ALL_SRV_MSG = 'All services are %s now.'
SELECT_SRV_MSG = 'Selected services are %s now.'
NOT_SELECT_SRV_STOP_MSG = 'Not selected services are STOPPED now.'
PENDING_MSG = 'Cannot %s services when the status is ' + PENDING.upper() + ', please try again later.'
PENDING_MSG_START = PENDING_MSG % 'START'
PENDING_MSG_STOP = PENDING_MSG % 'STOP'
LOADING_MSG = 'Now loading... Please wait.'
CFG_NOT_FOUND_MSG = 'Cannot find "srv_pak.cfg" in this directory. \n' \
                    'Please set correct configuration and press "Reload Config" button.'
MUST_SELECT_MSG = 'Please select at least one item.'

AFTER_MS = 100  # wait milliseconds and hide loading panel
SLEEP_SEC = 3  # wait seconds and reload service status

COLUMN_WIDTH = 20  # column width of service list
BTN_HEIGHT = 1  # function button height

CFG_FILE_PATH = './srv_pak.cfg'  # config file path


class WinServiceController(Frame):
    pkg_cfg = dict()  # package display name and its service names
    srv_status = dict()  # name and status of all services
    srv_list = list()  # package display name and its status
    select_pkg = dict()  # selected service name
    c = None  # WMI
    lp = None  # loading panel

    def __init__(self, master=None):
        if mode == MODE_DEBUG:
            print '====== init ======'
        Frame.__init__(self, master)
        # create static gui start
        # set window to screen center
        self._root().geometry('+%d+%d' % (self.winfo_screenwidth() / 3, self.winfo_screenheight() / 3))
        # create function buttons
        Button(self, text='Reload Config', width=COLUMN_WIDTH, height=BTN_HEIGHT,
               command=lambda: self.execute(self.load_config)).grid(row=0, column=0)
        Button(self, text='Reload Status', width=COLUMN_WIDTH, height=BTN_HEIGHT,
               command=lambda: self.execute(self.load_service_status)).grid(row=1, column=0)
        Button(self, text='Start', width=COLUMN_WIDTH, height=BTN_HEIGHT,
               command=lambda: self.execute(self.start)).grid(row=0, column=1)
        Button(self, text='Advanced Start', width=COLUMN_WIDTH, height=BTN_HEIGHT,
               command=lambda: self.execute(self.advanced_start)).grid(row=1, column=1)
        Button(self, text='Total Start', width=COLUMN_WIDTH, height=BTN_HEIGHT,
               command=lambda: self.execute(self.total_start)).grid(row=0, column=2)
        Button(self, text='Total Stop', width=COLUMN_WIDTH, height=BTN_HEIGHT,
               command=lambda: self.execute(self.total_stop)).grid(row=1, column=2)
        # create service list headers
        Label(self, text='Display Name', fg='blue').grid(row=2, column=0)
        Label(self, text='Status', fg='blue').grid(row=2, column=1)
        Label(self, text='Select', fg='blue').grid(row=2, column=2)
        # create loading panel
        self.lp = Toplevel(self._root())
        self.lp.title(LOADING)
        self.lp.attributes("-topmost", True)
        self.lp.protocol("WM_DELETE_WINDOW", lambda: None)
        Label(self.lp, text=LOADING_MSG).pack(padx=40, pady=40)
        self.lp.withdraw()
        self.grid()
        # create static gui finish
        self.c = wmi.WMI()  # get windows management instrumentation
        self.load_config()
        if mode == MODE_DEBUG:
            print '====== init finished ======'

    def execute(self, function):
        self.show_lp()
        self.master.after(AFTER_MS, lambda: function(after_func=self.hide_lp))

    def show_lp(self):
        self.master.grab_set()
        size = self._root().geometry().split('+')[0].split('x')
        wh = self._root().geometry().split('+')[1:3]
        x = int(wh[0]) + int(size[0]) / 2 - self.lp.winfo_width() / 2
        y = int(wh[1]) + int(size[1]) / 2 - self.lp.winfo_height() / 2
        self.lp.geometry('+%d+%d' % (x, y))
        self.lp.deiconify()

    def hide_lp(self):
        self.lp.withdraw()
        self.master.grab_release()

    def load_config(self, after_func=None):
        self.pkg_cfg = dict()
        # load package config from cfg file
        try:
            cfg_file = open(CFG_FILE_PATH)
            for cfg_line in cfg_file.readlines():
                self.pkg_cfg[cfg_line.split('=')[0].strip()] = map(str.strip, cfg_line.split('=')[1].split(','))
        except IOError:
            tkMessageBox.showerror(SYS_WARN, CFG_NOT_FOUND_MSG, parent=self._root())
        if mode == MODE_DEBUG:
            print '------ reload config ------'
            print 'pkg_cfg:', self.pkg_cfg
            print 'srv_status clear'
        # regenerate service status
        self.srv_status.clear()
        for srv_names in self.pkg_cfg.itervalues():
            for srv_name in srv_names:
                self.srv_status[srv_name] = LOADING
        if mode == MODE_DEBUG:
            print 'srv_status init:', self.srv_status
            print 'srv_list clear'
        # clear service list
        if self.srv_list:
            for srv_dict in self.srv_list:
                for srv_elem in srv_dict.itervalues():
                    srv_elem.destroy()
            del self.srv_list[:]
        if mode == MODE_DEBUG:
            print 'srv_list:', self.srv_list
        self.load_service_status()
        if after_func is not None:
            self.master.after(AFTER_MS, after_func())

    def load_service_status(self, wait=False, after_func=None):
        # avoid PENDING status
        if wait:
            time.sleep(SLEEP_SEC)
        if mode == MODE_DEBUG:
            print '------ reload service status ------'
            print 'srv_status reload'
        for srv_name in self.srv_status.iterkeys():
            status = self.srv_status[srv_name]
            if not status == CFG_ERR:  # ignore config error services
                self.srv_status[srv_name] = self.get_wmi_srv_status(srv_name)
        if mode == MODE_DEBUG:
            print 'srv_status:', self.srv_status
        if self.srv_status:  # if service status is not empty, refresh the service list
            self.refresh_dynamic_widget()
        if after_func is not None:
            self.master.after(AFTER_MS, after_func())

    def get_wmi_srv_status(self, srv_name):
        srv = self.c.Win32_Service(name=srv_name)
        if srv:
            status = srv[0].state
            if PENDING in status:
                status = PENDING
        else:
            status = CFG_ERR
        return status

    def refresh_dynamic_widget(self):
        if mode == MODE_DEBUG:
            print '------ refresh widget ------'
        if self.srv_list:  # if already exist, update them
            for srv_elem in self.srv_list:
                status = self.infer_pkg_status(srv_elem['pkg_name']['text'])
                pkg_name = srv_elem['pkg_name']['text']
                if mode == MODE_DEBUG:
                    print pkg_name, ':', status
                self.print_srv_status_of_pkg(pkg_name)
                srv_elem['pkg_status']['text'] = status
                status_color = STATUS_COLOR.get(status, dict())
                srv_elem['pkg_status']['fg'] = status_color.get('fg', DEFAULT_FG)
                srv_elem['pkg_status']['bg'] = status_color.get('bg', DEFAULT_BG)
                srv_elem['pkg_status'].update()
        else:  # generate service list
            counter = 0
            for pkg_name in self.pkg_cfg.iterkeys():
                srv_elem = dict()
                pkg_name_label = Label(self, text=pkg_name)
                pkg_name_label.grid(row=counter + 3, column=0)
                status = self.infer_pkg_status(pkg_name)
                if mode == MODE_DEBUG:
                    print pkg_name, ':', status
                self.print_srv_status_of_pkg(pkg_name)
                status_color = STATUS_COLOR.get(status, dict())
                status_label = Label(self, text=status, fg=status_color.get('fg', DEFAULT_FG),
                                     bg=status_color.get('bg', DEFAULT_BG))
                status_label.grid(row=counter + 3, column=1)
                srv_elem['pkg_name'] = pkg_name_label
                srv_elem['pkg_status'] = status_label
                if status != CFG_ERR:  # can't select config error services
                    self.select_pkg[pkg_name] = IntVar()
                    select_chk_btn = Checkbutton(self, text='', variable=self.select_pkg[pkg_name])
                    select_chk_btn.grid(row=counter + 3, column=2)
                    srv_elem['select'] = select_chk_btn
                self.srv_list.append(srv_elem)
                counter += 1

    def print_srv_status_of_pkg(self, pkg_name):
        for srv_name in self.pkg_cfg[pkg_name]:
            if mode == MODE_DEBUG:
                print '\t', srv_name, ':', self.srv_status[srv_name]

    def infer_pkg_status(self, pkg_name):
        srv_names = self.pkg_cfg[pkg_name]
        srv_status_set = set()
        pkg_status = LOADING
        for srv_name in srv_names:
            status = self.srv_status[srv_name]
            if status in [CFG_ERR, PENDING]:
                return status
            srv_status_set.add(self.srv_status[srv_name])
        if len(srv_status_set) > 1:
            pkg_status = COMPLEX
        elif len(srv_status_set) == 1:
            pkg_status = srv_status_set.pop()
        return pkg_status

    def is_available_to_act(self, fail_msg, check_select=False, after_func=None):
        # check is it available to use functions
        if check_select and all(check.get() == 0 for check in self.select_pkg.itervalues()):
            if after_func is not None:
                self.master.after(AFTER_MS, after_func())
            tkMessageBox.showwarning(SYS_WARN, MUST_SELECT_MSG, parent=self.lp)
            return False
        self.load_service_status()
        if PENDING in self.srv_status.itervalues():
            if after_func is not None:
                self.master.after(AFTER_MS, after_func())
            tkMessageBox.showwarning(SYS_WARN, fail_msg, parent=self.lp)
            return False
        return True

    def do_start(self, srv_status_items):
        result = list()
        for srv_name, status in srv_status_items.iteritems():
            if status == CFG_ERR:
                continue
            elif status == RUN:
                result.append(0)
            else:
                srv = self.c.Win32_Service(name=srv_name)
                if status == PAUSE:
                    result.append(srv[0].ResumeService()[0])
                else:
                    result.append(srv[0].StartService()[0])
        return result

    def after_do_start(self, result, success_msg, after_func, fail_msg=ERR_MSG):
        self.load_service_status(wait=True)
        self.master.after(AFTER_MS, after_func())
        if not result or all(x == 0 for x in result):
            tkMessageBox.showinfo(SYS_INFO, success_msg, parent=self._root())
        else:
            tkMessageBox.showerror(SYS_WARN, fail_msg, parent=self._root())

    def start(self, after_func):
        # start selected services
        if self.is_available_to_act(fail_msg=PENDING_MSG_START, check_select=True, after_func=after_func):
            select_srv = dict()
            for pkg_name, select_val in self.select_pkg.iteritems():
                if mode == MODE_DEBUG:
                    print pkg_name, ':', bool(select_val.get())
                if bool(select_val.get()):
                    for srv_name in self.pkg_cfg[pkg_name]:
                        select_srv[srv_name] = self.srv_status[srv_name]
            result = self.do_start(select_srv)
            self.after_do_start(result=result, success_msg=SELECT_SRV_MSG % RUN.upper(), after_func=after_func)

    def total_start(self, after_func):
        # start all services
        if self.is_available_to_act(fail_msg=PENDING_MSG_START, after_func=after_func):
            result = self.do_start(self.srv_status)
            self.after_do_start(result=result, success_msg=ALL_SRV_MSG % RUN.upper(), after_func=after_func)

    def do_stop(self, srv_status_items):
        result = list()
        for srv_name, status in srv_status_items.iteritems():
            if status == CFG_ERR:
                continue
            elif status == STOP:
                result.append(0)
            else:
                srv = self.c.Win32_Service(name=srv_name)
                response = srv[0].StopService()[0]
                result.append(response)
        return result

    def advanced_start(self, after_func):
        # start selected services and stop the others
        if self.is_available_to_act(fail_msg=PENDING_MSG_START, check_select=True, after_func=after_func):
            select_srv = dict()
            for pkg_name, select_val in self.select_pkg.iteritems():
                if mode == MODE_DEBUG:
                    print pkg_name, ':', bool(select_val.get())
                if bool(select_val.get()):
                    for srv_name in self.pkg_cfg[pkg_name]:
                        select_srv[srv_name] = self.srv_status[srv_name]
            not_select_srv = {key: value for key, value in self.srv_status.iteritems() if
                              key not in select_srv and value != CFG_ERR}
            if mode == MODE_DEBUG:
                print 'start:', select_srv
                print 'stop:', not_select_srv
            stop_result = self.do_stop(not_select_srv)
            start_result = self.do_start(select_srv)
            result = stop_result + start_result
            self.load_service_status(wait=True)
            self.master.after(AFTER_MS, after_func())
            if not result or all(x == 0 for x in result):
                tkMessageBox.showinfo(SYS_INFO, SELECT_SRV_MSG % RUN.upper() + '\n' + NOT_SELECT_SRV_STOP_MSG,
                                      parent=self._root())
            else:
                tkMessageBox.showerror(SYS_WARN, ERR_MSG, parent=self._root())

    def total_stop(self, after_func):
        # stop all services
        if self.is_available_to_act(fail_msg=PENDING_MSG_STOP, after_func=after_func):
            result = self.do_stop(self.srv_status)
            self.load_service_status(wait=True)
            self.master.after(AFTER_MS, after_func())
            if not result or all(x == 0 for x in result):
                tkMessageBox.showinfo(SYS_INFO, ALL_SRV_MSG % STOP.upper(), parent=self._root())
            else:
                tkMessageBox.showerror(SYS_WARN, ERR_MSG, parent=self._root())


if __name__ == '__main__':
    if sys.argv[1] == MODE_DEBUG:    
        global mode
        mode = MODE_DEBUG
    root = Tk()
    root.title('Windows Services Controller')
    app = WinServiceController(master=root)
    app.mainloop()
