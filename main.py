#!/usr/bin/env python

import uno
import unohelper
from com.sun.star.task import XJobExecutor

ID_EXTENSION = "org.ossii.writerbatchimport.do"
SERVICE = ("com.sun.star.task.Job",)

CTX = uno.getComponentContext()
SM = CTX.getServiceManager()

# The MainJob is a UNO component derived from unohelper.Base class
# and also the XJobExecutor, the implemented interface
class MainJob(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx

    def trigger(self, args):
        from batchimport import image_batch_import

        image_batch_import()
        return


# pythonloader loads a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation(
    MainJob,  # UNO object class
    ID_EXTENSION,  # implementation name (customize for yourself)
    SERVICE)  # implemented services (only 1)
