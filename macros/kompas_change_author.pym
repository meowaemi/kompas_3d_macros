# -*- coding: utf-8 -*-
#|kompas_folder_author
"""Batch author change for KOMPAS-3D v25 documents.

Opens every KOMPAS file in the script's folder, sets the author and saves
the changed files:

* drawings, specifications, text documents: title block column 110;
* parts and assemblies: the "Разработал" / "Автор" property;
* any document type: the file-level Author field.

Fragments are skipped.
"""

import os
import sys
import ctypes
import traceback
import pythoncom
from win32com.client import Dispatch, gencache

STAMP_COLUMN = 110                      # "Разраб." column of the title block
PROP_NAMES = ("Разработал", "Автор")    # must match KOMPAS property names exactly
RECURSIVE = False                       # also scan subfolders

KOMPAS_EXT = (".cdw", ".spw", ".kdw", ".frw", ".m3d", ".a3d", ".t3d")

# ksDocumentTypeEnum
DRAWING, FRAGMENT, SPEC, PART, ASSEMBLY, TEXT, TECH_ASSEMBLY = 1, 2, 3, 4, 5, 6, 7
TYPE_NAMES = {DRAWING: "drawing", FRAGMENT: "fragment", SPEC: "specification",
              PART: "part", ASSEMBLY: "assembly", TEXT: "text document",
              TECH_ASSEMBLY: "technological assembly"}

TITLE = "Change author"


# "Run macro" starts the script without a console, so print() goes nowhere.
# Output is collected here and shown in a message box at the end.
LOG = []


def log(*args):
    line = " ".join(str(a) for a in args)
    LOG.append(line)
    try:
        print(line)
    except Exception:
        pass


def message_box(text, icon=0x40):
    ctypes.windll.user32.MessageBoxW(0, text, TITLE, icon | 0x40000 | 0x10000)


def ask_string(prompt, width=32):
    """Ask for a string in a dialog, falling back to the console."""
    try:
        import tkinter as tk
        from tkinter import ttk
    except Exception:
        # A newline inside the input() prompt makes the KOMPAS console
        # mask the typed text, so it is printed separately.
        print()
        return input(prompt + " ").strip()

    # Without DPI awareness Windows scales the window up and blurs the text.
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    result = {"value": ""}

    root = tk.Tk()
    root.withdraw()                       # hidden until centered, no jump
    root.title(TITLE)
    root.resizable(False, False)

    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    style.configure("TLabel", font=("Segoe UI", 9))
    style.configure("TButton", font=("Segoe UI", 9))

    frame = ttk.Frame(root, padding=(16, 14, 16, 12))
    frame.pack(fill="both", expand=True)

    ttk.Label(frame, text=prompt).pack(anchor="w", pady=(0, 4))
    entry = ttk.Entry(frame, width=width, font=("Segoe UI", 9))
    entry.pack(fill="x")

    def ok(event=None):
        result["value"] = entry.get()
        root.destroy()

    def cancel(event=None):
        root.destroy()

    buttons = ttk.Frame(frame)
    buttons.pack(pady=(12, 0))            # centered under the entry
    ttk.Button(buttons, text="OK", width=9, command=ok).pack(side="left", padx=4)
    ttk.Button(buttons, text="Cancel", width=9, command=cancel).pack(side="left", padx=4)

    root.bind("<Return>", ok)
    root.bind("<Escape>", cancel)
    root.protocol("WM_DELETE_WINDOW", cancel)

    # Center using the requested size: the real size is unknown until shown.
    root.update_idletasks()
    w, h = root.winfo_reqwidth(), root.winfo_reqheight()
    x = (root.winfo_screenwidth() - w) // 2
    y = (root.winfo_screenheight() - h) // 2
    root.geometry("{0}x{1}+{2}+{3}".format(w, h, x, y))
    root.deiconify()

    # Topmost is re-applied after mapping: Windows may ignore it before.
    def bring_to_front():
        root.attributes("-topmost", True)
        root.lift()
        root.focus_force()
        entry.focus_set()

    bring_to_front()
    root.after(100, bring_to_front)

    root.mainloop()
    return result["value"].strip()


def ask_folder(initial):
    """Let the user pick a folder. Returns None if cancelled or unavailable."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askdirectory(title="Folder with KOMPAS files",
                                       initialdir=initial, parent=root)
        root.destroy()
        return path or None
    except Exception:
        return None


def iterate(collection):
    """Return the items of a KOMPAS collection as a list.

    Count is not always exposed through pywin32, so items are read
    by index until Item() fails or returns None.
    """
    result, i = [], 0
    while collection is not None:
        try:
            item = collection.Item(i)
        except Exception:
            break
        if item is None:
            break
        result.append(item)
        i += 1
    return result


def norm(path):
    return os.path.normcase(os.path.abspath(path))


def script_folder():
    # __file__ may be undefined when the script is started via "Run macro".
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        pass
    if sys.argv and sys.argv[0]:
        return os.path.dirname(os.path.abspath(sys.argv[0]))
    return os.getcwd()


def find_files(folder):
    found = []
    if not folder or not os.path.isdir(folder):
        return found
    if RECURSIVE:
        for root, _, names in os.walk(folder):
            for n in names:
                if os.path.splitext(n)[1].lower() in KOMPAS_EXT:
                    found.append(os.path.join(root, n))
    else:
        for n in os.listdir(folder):
            path = os.path.join(folder, n)
            if os.path.isfile(path) and os.path.splitext(n)[1].lower() in KOMPAS_EXT:
                found.append(path)
    return sorted(found)


def open_files(application, files):
    """Open the given files, reusing documents that are already open."""
    docs = application.Documents
    open_docs = {}
    for d in iterate(docs):
        try:
            open_docs[norm(d.PathName)] = d
        except Exception:
            pass

    result = []
    for path in files:
        name = os.path.basename(path)
        key = norm(path)
        if key in open_docs:
            log("  already open:", name)
            result.append(open_docs[key])
            continue
        try:
            doc = docs.Open(path, True, False)   # path, visible, read-only
            if doc is None:
                raise RuntimeError("Open returned nothing")
            log("  opened:", name)
            result.append(doc)
        except Exception as e:
            log("  FAILED to open {0}: {1}".format(name, e))
    return result


def change_stamp(doc, author):
    sheets = iterate(doc.LayoutSheets)
    if not sheets:
        log("    title block not found")
        return False
    for s in sheets:
        stamp = s.Stamp
        stamp.Text(STAMP_COLUMN).Str = author
        stamp.Update()
    log("    'Developed by' updated, sheets:", len(sheets))
    return True


def change_properties(api7, application, doc, author):
    prop_mng = api7.IPropertyMng(application)
    part = api7.IKompasDocument3D(doc).TopPart
    keeper = api7.IPropertyKeeper(part)

    props = prop_mng.GetProperties(doc) or ()
    target = [p for p in props if p.Name in PROP_NAMES]
    if not target:
        log("    property not found, available:",
            ", ".join(p.Name for p in props))
        return False

    for p in target:
        ok = keeper.SetPropertyValue(p, author, False)
        log("    property '{0}': {1}".format(p.Name, "ok" if ok else "FAILED"))
    part.Update()
    return True


def change_doc_author(api7, doc, author):
    try:
        doc1 = api7.IKompasDocument1(doc)
        doc1.Author = author
        log("    document author: ok")
        return True
    except Exception as e:
        log("    document author: error -", e)
        return False


def process(api7, application, doc, author):
    """Set the author in one document. Returns True if anything changed."""
    doc_type = doc.DocumentType
    log("")
    log("{0}  [{1}]".format(doc.Name, TYPE_NAMES.get(doc_type, doc_type)))

    if doc_type not in TYPE_NAMES:
        log("    skipped")
        return False

    # Fragments have no title block, only the file-level author.
    if doc_type == FRAGMENT:
        return change_doc_author(api7, doc, author)

    try:
        if doc_type in (DRAWING, SPEC, TEXT):
            ok = change_stamp(doc, author)
        else:
            ok = change_properties(api7, application, doc, author)
    except Exception as e:
        log("    error:", e)
        ok = False

    ok = change_doc_author(api7, doc, author) or ok
    return ok


def main():
    api7 = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    application = api7.IApplication(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(
            api7.IApplication.CLSID, pythoncom.IID_IDispatch))

    folder = script_folder()
    files = find_files(folder)
    if not files:
        picked = ask_folder(folder)
        if picked:
            folder = picked
            files = find_files(folder)

    log("Folder:", folder)
    if not files:
        log("No KOMPAS files found")
        return

    log("Found files:", len(files))
    for f in files:
        log("  -", os.path.basename(f))

    # Ask before opening anything, so an empty answer costs nothing.
    author = ask_string("New author:")
    if not author:
        log("Empty name - nothing changed")
        return
    log("New author:", author)

    log("")
    log("Opening files...")
    docs = open_files(application, files)

    changed = [d for d in docs if process(api7, application, d, author)]

    log("")
    log("Documents changed: {0} of {1}".format(len(changed), len(docs)))

    for d in changed:
        try:
            d.Save()
            log("  saved:", d.Name)
        except Exception as e:
            log("  NOT saved {0}: {1}".format(d.Name, e))


try:
    main()
    report = "\n".join(LOG[-40:])
    message_box(report or "Done")
except Exception:
    message_box("ERROR\n\n" + traceback.format_exc() +
                "\n\nLog:\n" + "\n".join(LOG[-15:]), 0x10)