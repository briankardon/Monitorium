import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import sys
import time
from pathlib import Path

# Thanks for the guidance from:
#   https://dev.to/demola12/building-a-robust-windows-service-in-python-with-win32serviceutil-part-13-1k6k
#   https://thepythoncorner.com/posts/2018-08-01-how-to-create-a-windows-service-in-python/

class MonitoriumService(win32serviceutil.ServiceFramework):
    _svc_name_ = 'Monitorium'
    _svc_display_name_ = 'Monitorium'
    _svc_description_ = 'Server infrastructure monitoring service'

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(120)
        self.is_alive = True

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.is_alive = False

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()

    def main(self):
        # Main service logic goes here
        root = Path(r'C:\Users\briankardon\Downloads')
        name = 'test.txt'
        newname = 'test_CHANGED.txt'
        path = root / name
        newpath = root / newname
        while self.is_alive:
            # Perform your service tasks here
            time.sleep(2)  # Example: Sleep for 5 seconds
            # Check if file exists:
            if path.exists() and not newpath.exists():
                try:
                    path.rename(newpath)
                except:
                    self.SvcStop()
                    break

if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(MonitoriumService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(MonitoriumService)

# Build:
#   pyinstaller --hidden-import=win32timezone Source\monitorium.py
# Deploy:
#   dist/monitorium/monitorium.exe install
