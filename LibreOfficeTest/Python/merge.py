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


def ensure_master_in_target(src_master, target_doc):
    """Clone master page from source into target if not already present."""
    target_masters = target_doc.getMasterPages()
    try:
        name = src_master.getName()
    except Exception:
        name = "ImportedMaster"

    # Look for existing master with same name
    for i in range(target_masters.getCount()):
        existing = target_masters.getByIndex(i)
        try:
            if existing.getName() == name:
                return existing
        except Exception:
            pass

    # If not found, insert a new master and copy shapes
    new_master = target_masters.insertNewByIndex(target_masters.getCount())
    try:
        new_master.setName(name)
    except Exception:
        pass

    for i in range(src_master.getCount()):
        shape = src_master.getByIndex(i)
        try:
            clone = shape.Duplicate()
            new_master.add(clone)
        except Exception:
            pass

    return new_master


def copy_slide_simple(source_doc, source_index, target_doc, ctx):
    """Copy one slide from source_doc to target_doc using dispatcher Copy/Paste."""
    try:
        source_controller = source_doc.getCurrentController()
        target_controller = target_doc.getCurrentController()

        source_pages = source_doc.getDrawPages()
        target_pages = target_doc.getDrawPages()

        # Select source slide
        source_slide = source_pages.getByIndex(source_index)
        source_controller.setCurrentPage(source_slide)

        # Select all shapes on the slide (text, images, etc.)
        selection_supplier = source_controller
        selection_supplier.select(source_slide)

        # Dispatcher for copy/paste
        dispatcher = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.DispatchHelper", ctx
        )

        # Copy
        dispatcher.executeDispatch(source_controller.getFrame(), ".uno:Copy", "", 0, ())

        # Create new slide in target
        target_pages.insertNewByIndex(target_pages.getCount())
        target_slide = target_pages.getByIndex(target_pages.getCount() - 1)
        target_controller.setCurrentPage(target_slide)

        # Paste
        dispatcher.executeDispatch(target_controller.getFrame(), ".uno:Paste", "", 0, ())

        # Clone and apply master page from source slide
        try:
            src_master = source_slide.getMasterPage()
            tgt_master = ensure_master_in_target(src_master, target_doc)
            target_slide.setMasterPage(tgt_master)
            print(f"✅ Copied slide {source_index + 1} with cloned master page")
        except Exception as e:
            print(f"⚠️ Could not clone master page for slide {source_index + 1}: {e}")
            try:
                base_master = target_doc.getMasterPages().getByIndex(0)
                target_slide.setMasterPage(base_master)
                print(f"➡️ Applied base master page instead for slide {source_index + 1}")
            except Exception as e2:
                print(f"❌ Could not set any master page for slide {source_index + 1}: {e2}")

        return True
    except Exception as e:
        print(f"❌ Copy-paste failed for slide {source_index + 1}: {e}")
        return False


def merge_presentations(inputs, output_file, fmt="odp"):
    ctx = connect_uno()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    hidden = [PropertyValue("Hidden", 0, True, 0)]

    # Open first file as base
    base_doc = desktop.loadComponentFromURL(
        uno.systemPathToFileUrl(os.path.abspath(inputs[0])), "_blank", 0, tuple(hidden)
    )
    base_pages = base_doc.getDrawPages()

    # Append slides from the rest
    for path in inputs[1:]:
        doc = desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(os.path.abspath(path)), "_blank", 0, tuple(hidden)
        )
        pages = doc.getDrawPages()
        for i in range(pages.getCount()):
            copy_slide_simple(doc, i, base_doc, ctx)
        doc.close(True)
        print(f"Merged: {path}")

    # Save in requested format
    filter_map = {
        "odp": "impress8",
        "xodp": "impress8",
        "pptx": "Impress MS PowerPoint 2007 XML",
        "pdf": "impress_pdf_Export",
    }
    if fmt not in filter_map:
        raise ValueError(f"Unsupported format: {fmt}")

    out_url = uno.systemPathToFileUrl(os.path.abspath(output_file))
    export_props = [PropertyValue("FilterName", 0, filter_map[fmt], 0)]
    base_doc.storeToURL(out_url, tuple(export_props))
    base_doc.close(True)
    print(f"✅ Merged into {output_file}")


def main():
    parser = argparse.ArgumentParser(description="Merge multiple presentations using UNO only.")
    parser.add_argument("inputs", nargs="+", help="Input presentations (odp/pptx/...)")
    parser.add_argument("-o", "--output", required=True, help="Output file")
    parser.add_argument("--format", choices=["odp", "pptx", "xodp", "pdf"], default="odp", help="Output format")
    args = parser.parse_args()

    if len(args.inputs) < 2:
        sys.exit("Need at least 2 inputs.")

    merge_presentations(args.inputs, args.output, args.format)


if __name__ == "__main__":
    main()
