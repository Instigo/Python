import win32service
import win32serviceutil
import win32event
import win32evtlogutil
import win32api
import wmi
import psutil
import datetime
import servicemanager
import logging
import time
import os

# Description: The purpose for this script was to monitor the Tomcat7.exe process within Windows and
# to keep track of the CPU utilization.  This script could be further expanded to provide email alerts,
# automatic killing of the Tomcat7.exe process (if high CPU lingers over a period of time).  This can be tweaked
# to monitor some other process entirely.
#  
#
# Main Requirements: Python 2.7, Modules: {psutil,pywin32, python WMI}
#

# Setup the basic configuration for logging state information
#
logging.basicConfig(
    filename = 'C:\\Program Files (x86)\\Atlassian\\JIRA\\logs\\Tomcat-Cpu-Service.log',
    level = logging.DEBUG,
    format = '[tomcat-cpu-service] %(levelname)-7.7s %(message)s',) 

class TomcatSvc(win32serviceutil.ServiceFramework):
    # you can NET START/STOP the service by the following name
    _svc_name_ = "TomcatSvc"
    # this text shows up as the service name in the Service Control Manager (SCM)
    _svc_display_name_ = "Tomcat CPU Monitor Service"
    # this text shows up as the description in the SCM
    _svc_description_ = "This services monitors the CPU utilization for Tomcat7.exe"
    
    # path to the JIRA logs 
    _svc_log_folder = "C:\\Program Files (x86)\\Atlassian\\JIRA\\logs\\"

    # name of the file to log the CPU percentage 
    _svc_log_filename = "tomcat-cpu-log.txt"
    
    numTimesExecuted = 0
    now = datetime.datetime.now()
    start_time = time.clock()

    # windows management instrumentation interface
    wm = wmi.WMI()

    processNameToFind = u'tomcat7.exe'
    memory_percent = None
    executionFrequency = 1800 # number of seconds in an 30 minute span

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self,args)
        # create an event to listen for stop requests on
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)

    # core logic of the service   
    def SvcDoRun(self):
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, '')
        )
        create_process_log_header(self._svc_log_folder)
        # This is how long the service will wait to run / refresh itself
        self.timeout = 1000 # (value is in milliseconds)
        while 1:
            # Wait for service stop signal, if I timeout, loop again
            rc = win32event.WaitForSingleObject(self.hWaitStop, self.timeout)
            # Check to see if self.hWaitStop happened
            if rc == win32event.WAIT_OBJECT_0:
                # Stop signal encountered
                servicemanager.LogInfoMsg("Tomcat CPU Monitor Service - STOPPED!")  #For Event Log
                break
            else:
                try:
                    self.now = datetime.datetime.now()
                    end_time = (time.clock() - self.start_time)
                    tomcatProcess = find_process(self.processNameToFind)
                    memoryPercent = tomcatProcess.memory_percent()
                    
                    if (memoryPercent > 75.0): 
                        filename = self._svc_log_folder + self._svc_log_filename
                        f = open(filename,'a')
                        f.write('*Tomcat7.exe Running At High CPU Percentage: %6.2f%% - Date: %s*\n' % (self.process.memory_percent(), self.now.strftime("%Y-%m-%d %H:%M:%S")))
                        f.close()
                    else:
                        self.numTimesExecuted += 1
                        create_process_log(self._svc_log_folder, self.processNameToFind)
                        # if the service has run 
                        if (self.numTimesExecuted >= self.executionFrequency or end_time >= self.executionFrequency ):
                            #create_process_log(self._svc_log_folder, self.processNameToFind)
                            self.numTimesExecuted = 0
                            self.start_time = time.clock()
                except Exception as err:
                    logging.exception("error thrown: %s, %s", err.__class__.__name__, err)
                    pass
    

    # called when we're being shut down    
    def SvcStop(self):
        # tell the SCM we're shutting down
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        # fire the stop event
        win32event.SetEvent(self.hWaitStop)

def ctrlHandler(ctrlType):
    return True

def find_process(process_name):

    wm = wmi.WMI()
    for currentProcess in wm.Win32_Process():
        if process_name in currentProcess.Name:
            return psutil.Process(currentProcess.ProcessId)

def create_process_log_header(log_dir):
    if not os.path.exists(log_dir):
        try:
            os.mkdir(log_dir)
        except:
            pass

    separator = "-" * 90
    columnFormat = "%10s %10s %10s %12s %12s %30s"
    
    logPath = os.path.join(log_dir, "tomcat7Perf.log")
    if not os.path.exists(logPath):
        f = open(logPath, "a+")
        f.write(separator + "\n")
        f.write(columnFormat % ("DATE", "%CPU", "%MEMORY", "VMS", "RSS", "NAME"))
        f.write("\n")

def pretty_size(nbytes):
    i = 0
    suffixes = ['B','KB','MB','GB','TB','PB']
    if nbytes == 0: return '0 B'
    while nbytes >= 1024 and i < len(suffixes)-1:
        nbytes /= 1024
        i += 1
    f = ('%.2f' % nbytes).rstrip('0').rstrip('.')
    return '%s %s' % (f, suffixes[i])

def create_process_log(log_dir, process_name):

    if not os.path.exists(log_dir):
        try:
            os.mkdir(log_dir)
        except:
            pass
    
    process = find_process(process_name)
    dataFormat = "%10s %7.4f %7.2f %12s %12s %30s"
    logPath = os.path.join(log_dir, "tomcat7Perf.log")
    f = open(logPath, "a+")

    cpu_percent = process.get_cpu_percent()
    mem_percent = process.get_memory_percent()
    rss, vms = process.get_memory_info()
    rss = str(rss)
    vms = str(vms)
    name = process_name
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %I:%M:%S")
    f.write(dataFormat % (current_time, cpu_percent, mem_percent, pretty_size(float(vms)), pretty_size(float(rss)), name))
    f.write("\n")
    f.close()

if __name__ == '__main__':
    win32api.SetConsoleCtrlHandler(ctrlHandler, True)
    win32serviceutil.HandleCommandLine(TomcatSvc)
