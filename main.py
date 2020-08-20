import win32serviceutil
import win32service
import win32event
import subprocess
import servicemanager
import socket
import signal
import sys
import os


class ServiceWrapper(win32serviceutil.ServiceFramework):
    _svc_name_ = "TestService"
    _svc_display_name_ = "Test Service"
    _svc_description_ = "My service description"

    def __init__(self, args):
        self.stream_server = None
        self.pipe = open(os.path.dirname(sys.executable) + '\\Service.log', 'a')
        try:
            self.stream_server = subprocess.Popen(["python", "main.py"], shell=True, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, cwd=os.path.dirname(sys.executable))
            try:
                self.pipe.write(str(self.stream_server.stdout.readline().decode().encode("UTF-8")) + '\n')
            except TimeoutError as e:
                self.pipe.write('timeout\n')
            self.pipe.flush()
        except Exception as e:
            self.pipe.write("Exception happened: \n")
            self.pipe.write(str(e) + "\n")
            self.pipe.flush()
        finally:
            self.pipe.write("self:")
            self.pipe.write(str(self.stream_server) + '\n')
            self.pipe.flush()

        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)

    def SvcStop(self):
        self.pipe.close()
        os.kill(signal.CTRL_C_EVENT, 0)
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        rc = None
        while rc != win32event.WAIT_OBJECT_0:
            try:
                self.pipe.write(str(self.stream_server.stdout.readline().strip()) + '\n')
            except TimeoutError as e:
                self.pipe.write("Exception happened: \n")
                self.pipe.write(str(e) + "\n")
            self.pipe.flush()

            rc = win32event.WaitForSingleObject(self.hWaitStop, 500)  # service commands acquire


if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(ServiceWrapper)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(ServiceWrapper)