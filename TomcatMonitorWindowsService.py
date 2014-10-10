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

# Description: The purpose for this script was to monitor the Tomcat7.exe process within Windows and
# to keep track of the CPU utilization.  This script can be further expanded to provide email alerts,
# automatic killing of the Tomcat7.exe process (if high CPU lingers over a period of time).  
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

    process = None
    processNameToFind = u'tomcat7'
    memory_percent = None
   
    # loop over all processes looking for the process
    for currentProcess in wm.Win32_Process ():
        if processNameToFind in currentProcess.Name:
            process = psutil.Process(currentProcess.ProcessId)
            break

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
                    memoryPercent = self.process.memory_percent()

                    if (memoryPercent > 50.0): 
                        filename = self._svc_log_folder + self._svc_log_filename
                        f = open(filename,'a')
                        f.write('*Tomcat7.exe Running At High CPU Percentage: %6.2f%% - Date: %s*\n' % (self.process.memory_percent(), self.now.strftime("%Y-%m-%d %H:%M:%S")))
                        f.close()
                    else:
                        self.numTimesExecuted += 1
                        if (self.numTimesExecuted >= 20 and end_time >= 20 ):
                            logging.info('Tomcat7.exe CPU is operating at normal levels ( %6.2f%% ) - Date: %s' 
                                % (self.process.memory_percent(), self.now.strftime("%Y-%m-%d %H:%M:%S")))
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

if __name__ == '__main__':
    win32api.SetConsoleCtrlHandler(ctrlHandler, True)
    win32serviceutil.HandleCommandLine(TomcatSvc)

