import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import sys
import time
from pathlib import Path
import yaml

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
        self.monitorium_params = None

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)
        self.is_alive = False

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()

    def load_config(self):
        # Load config file
        if getattr(sys, 'frozen', False):
            # If the application is run as a bundle, the PyInstaller bootloader
            # extends the sys module by a flag frozen=True and sets the app
            # path into variable _MEIPASS'.
            application_path = Path(sys._MEIPASS).resolve()
        else:
            application_path = Path.absolute(__file__).resolve().parent

        config_path = application_path / 'config.yml'

        success = True

        try:
            servicemanager.LogErrorMsg("Looking for config file: {f}".format(f=str(config_path)))
            with open(config_path, 'r') as f:
                self.monitorium_params = yaml.safe_load(f)
        except FileNotFoundError:
            servicemanager.LogErrorMsg("Couldn't find config file!")
            success = False
        except yaml.YAMLError as exc:
            if hasattr(exc, 'problem'):
                error_msg = exc.problem
            else:
                error_msg = 'unknown problem'
            servicemanager.LogErrorMsg("Error parsing config file: {problem}".format(problem=error_msg))
            success = False

        # Check that required config params exist:
        required_params = [
            'newname',
        ]
        for required_param in required_params:
            if required_param not in self.monitorium_params:
                servicemanager.LogErrorMsg("Error in config file - missing required parameter \"{param}\"".format(param=required_param))
                success = False

        if not success:
            raise Exception('Configuration error!')

    def main(self):
        # Main service logic goes here
        self.load_config()

        root = Path(r'C:\Users\briankardon\Downloads')
        name = 'test.txt'
        newname = self.monitorium_params['newname']
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
