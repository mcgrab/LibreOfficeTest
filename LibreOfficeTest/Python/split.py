import os
import sys
import argparse
import uno
from com.sun.star.beans import PropertyValue


def connect_uno():
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )
    ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
    return ctx


def split_odp(input_odp, output_dir, fmt="odp"):
    ctx = connect_uno()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    props = [PropertyValue("Hidden", 0, True, 0)]
    doc = desktop.loadComponentFromURL(
        uno.systemPathToFileUrl(os.path.abspath(input_odp)), "_blank", 0, tuple(props)
    )

    draw_pages = doc.getDrawPages()
    num_pages = draw_pages.getCount()
    doc.close(True)

    # LibreOffice filter names
    filter_map = {
        "odp": "impress8",
        "xodp": "impress8",
        "pdf": "impress_pdf_Export",
        "pptx": "Impress MS PowerPoint 2007 XML",
        "png": "impress_png_Export",
        "jpeg": "impress_jpg_Export",
        "jpg": "impress_jpg_Export",
        "webp": "impress_webp_Export",
    }
    if fmt not in filter_map:
        raise ValueError(f"Unsupported format: {fmt}")

    for i in range(num_pages):
        subdoc = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(os.path.abspath(input_odp)), "_blank", 0, tuple(props)
        )
        subpages = subdoc.getDrawPages()

        # delete all slides except i
        for j in range(num_pages - 1, -1, -1):
            if j != i:
                subpages.remove(subpages.getByIndex(j))

        export_props = [PropertyValue("FilterName", 0, filter_map[fmt], 0)]

        out_name = f"{os.path.splitext(os.path.basename(input_odp))[0]}_slide{i+1}.{fmt}"
        out_path = os.path.join(output_dir, out_name)
        out_url = uno.systemPathToFileUrl(os.path.abspath(out_path))

        subdoc.storeToURL(out_url, tuple(export_props))
        subdoc.close(True)

        print(f"Exported slide {i+1} -> {out_path}")


def main():
    parser = argparse.ArgumentParser(description="Split ODP into multiple files.")
    parser.add_argument("input", help="Input ODP file")
    parser.add_argument("output_dir", help="Directory to save slides")
    parser.add_argument("--format", choices=["odp", "pdf", "xodp", "pptx", "png", "jpeg", "jpg", "webp"], default="odp", help="Output format")
    args = parser.parse_args()

    if not os.path.exists(args.input):
        sys.exit(f"Input file not found: {args.input}")

    os.makedirs(args.output_dir, exist_ok=True)

    split_odp(args.input, args.output_dir, args.format)


if __name__ == "__main__":
    main()
