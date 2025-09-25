import sys
import uno
import os
from com.sun.star.beans import PropertyValue

def connect_uno():
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )
    ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    return ctx

def convert(input_path, output_path, fmt):
    ctx = connect_uno()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    hidden = [PropertyValue("Hidden", 0, True, 0)]
    doc = desktop.loadComponentFromURL(
        uno.systemPathToFileUrl(os.path.abspath(input_path)),
        "_blank",
        0,
        tuple(hidden)
    )

    filter_map = {
        "odp": "impress8",
        "xodp": "impress8",
        "pptx": "Impress MS PowerPoint 2007 XML",
        "pdf": "impress_pdf_Export",
    }

    if fmt not in filter_map:
        raise ValueError(f"Unsupported format: {fmt}")

    export_props = [PropertyValue("FilterName", 0, filter_map[fmt], 0)]
    out_url = uno.systemPathToFileUrl(os.path.abspath(output_path))
    doc.storeToURL(out_url, tuple(export_props))
    doc.close(True)

def main():
    if len(sys.argv) < 4:
        print("Usage: convert.py <input> <output> <format>")
        sys.exit(1)

    input_file, output_file, fmt = sys.argv[1], sys.argv[2], sys.argv[3]
    convert(input_file, output_file, fmt)

if __name__ == "__main__":
    main()
