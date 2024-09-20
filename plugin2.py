import uno
import unohelper
from com.sun.star.awt import XDialog

class SimpleDialog(unohelper.Base):
    def __init__(self, ctx):
        self.ctx = ctx
        self.sm = ctx.getServiceManager()

    def show_message_box(self):
        toolkit = self.sm.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)
        message_box = toolkit.createMessageBox(None, 0, 0, "Test Message", "This is a test message box.")
        message_box.execute()

def plugin():
    ctx = uno.getComponentContext()
    dialog = SimpleDialog(ctx)
    dialog.show_message_box()

if __name__ == "__main__":
    plugin()

