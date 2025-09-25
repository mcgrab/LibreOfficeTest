import sys
import uno
from com.sun.star.connection import NoConnectException

try:
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )

    # Try connecting to soffice on default UNO port
    ctx = resolver.resolve(
        "uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext"
    )

    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    # If we can get an XDesktop, UNO is alive
    sys.exit(0)
except NoConnectException:
    print("LibreOffice not responding on UNO socket")
    sys.exit(1)
except Exception as e:
    print("LibreOffice UNO error:", e)
    sys.exit(1)
