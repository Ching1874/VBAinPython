import abc
import win32com.client as win32
import sys
import shutil
from win32com.client import gencache
import functools
import threading
import pythoncom

def dispatch(clsid, new_instance=True):
    """Create a new COM instance and ensure cache is built,
       unset read-only gencache flag"""
    if new_instance:
        clsid = pythoncom.CoCreateInstanceEx(clsid, None, pythoncom.CLSCTX_SERVER,
                                             None, (pythoncom.IID_IDispatch,))[0]
    if gencache.is_readonly:
        # fix for "freezed" app: py2exe.org/index.cgi/UsingEnsureDispatch
        gencache.is_readonly = False
        gencache.Rebuild()
    try:
        return gencache.EnsureDispatch(clsid)
    except (KeyError, AttributeError):  # no attribute 'CLSIDToClassMap'
        # something went wrong, reset cache
        shutil.rmtree(gencache.GetGeneratePath())
        for i in [i for i in sys.modules if i.startswith("win32com.gen_py.")]:
            del sys.modules[i]
        return gencache.EnsureDispatch(clsid)

class DispatchInterface(metaclass=abc.ABCMeta):
    @classmethod
    def __subclasshook__(cls, subclass):
        return (hasattr(subclass, 'open') and
                callable(subclass.open) and
                hasattr(subclass, 'close') and
                callable(subclass.close) and
                NotImplemented)

    @abc.abstractmethod
    def open(self, file: str):
        raise NotImplementedError

    # @abc.abstractmethod
    # def close(self):
    #     raise NotImplementedError


class Application(DispatchInterface):
    instances = []

    def __init__(self, clsid):
        Application.instances.append(self)
        self.opened = None
        self.path = None
        self.clsid = clsid # Word.Application or Excel.Application
        self.threading_thread = None
        self.threading_event = threading.Event()
        self.task = []
        self.interface = None
        self.app = None
        self.quitApp = None

    def worker(self):
        pythoncom.CoInitialize()
        # if self.clsid == "Word.Application":
        #     self.opened = False
        #     self.app = dispatch(self.clsid)
        #     doc = self.app.Documents.Add()
        #     doc.Application.Visible = False
        #     doc.Application.ScreenUpdating = False
        #     # Set the border spacing
        #     # Get all sections in the document
        #     sections = doc.Sections
        #     # Loop through each section and set the border spacing
        #     for sec in sections:
        #         sec.Range.PageSetup.TopMargin = 5
        #         sec.Range.PageSetup.BottomMargin = 5
        #         sec.Range.PageSetup.LeftMargin = 5
        #         sec.Range.PageSetup.RightMargin = 5
        #
        #     self.interface = doc

        if self.clsid == "Excel.Application":
            self.app = win32.Dispatch(self.clsid)
            wbs = [wb.Path.replace("\\", "/") + "/" + wb.Name for wb in self.app.Workbooks]
            if self.path in wbs:
                # print('excel opened.')
                if len(wbs) == 1:
                    self.quitApp = True
                else:
                    self.quitApp = False
                self.opened = True
                wb = win32.GetObject(self.path)
                wb.Application.ScreenUpdating = False
            else:
                # print('excel not opened.')
                if len(wbs) == 0:
                    self.quitApp = True
                else:
                    self.quitApp = False
                self.opened = False
                if self.quitApp:
                    self.app.Quit()
                self.app = dispatch(self.clsid)
                wb = self.app.Workbooks.Open(self.path)
                wb.Application.Visible = False
                wb.Application.ScreenUpdating = False
            self.interface = wb
        else:
            raise Exception("- Error: only accept Excel.Application!")

        # task
        for job in self.task:
            self.threading_event.clear()
            job()
            self.threading_event.wait()


        self.interface.Application.ScreenUpdating = True
        if not self.opened:
            self.interface.Close(True)
            if self.quitApp:
                self.app.Quit()

        Application.instances.remove(self)

        if len(Application.instances) == 0:
            pythoncom.CoUninitialize()

    def open(self, file):
        # self.path = os.getcwd() + "/" + file
        self.path = file.replace("\\", "/")

    def start(self):
        self.threading_thread = threading.Thread(target=self.worker) # args=(end_event,)
        self.threading_thread.start()

    def add_task(self, job, *args, **kwargs):
        if args or kwargs:
            # Use functools.partial to create a new function with the arguments
            job = functools.partial(job, *args, **kwargs)
        self.task.append(job)