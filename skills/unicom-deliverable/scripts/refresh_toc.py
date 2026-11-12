#!/usr/bin/env python3
"""用 LibreOffice 刷新目录页码，另存 docx 并导出 PDF 供版面检查。

必须用自带 UNO 的解释器运行（通常是 /usr/bin/python3），并在外层加 timeout。
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import time


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source")
    parser.add_argument("docx")
    parser.add_argument("pdf")
    args = parser.parse_args()
    try:
        import uno
        from com.sun.star.beans import PropertyValue
    except ImportError:
        sys.exit("FAIL: 当前解释器没有 UNO；请改用 LibreOffice 自带或系统的 python3（如 /usr/bin/python3）")
    soffice = shutil.which("soffice")
    if not soffice:
        sys.exit("FAIL: 未找到 soffice")
    source, docx, pdf = (os.path.abspath(path) for path in (args.source, args.docx, args.pdf))
    profile = tempfile.mkdtemp(prefix="lo-profile-")
    port = 20000 + os.getpid() % 20000
    office = subprocess.Popen(
        [soffice, "--headless", "--norestore", "--nologo", f"-env:UserInstallation=file://{profile}",
         f"--accept=socket,host=127.0.0.1,port={port};urp;"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    def value(name, content):
        item = PropertyValue()
        item.Name, item.Value = name, content
        return item

    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        context = None
        for _ in range(120):
            try:
                context = resolver.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext")
                break
            except Exception:  # 连接未就绪时 UNO 抛的异常类型不固定，只用于重试
                time.sleep(0.5)
        if context is None:
            sys.exit("FAIL: 连不上 soffice")
        desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
        document = desktop.loadComponentFromURL("file://" + source, "_blank", 0, (value("Hidden", True),))
        indexes = document.getDocumentIndexes()
        if indexes.getCount() < 1:
            sys.exit("FAIL: 文档没有目录")
        for _ in range(2):  # 第二遍让页码在目录自身变长后稳定
            for index in range(indexes.getCount()):
                indexes.getByIndex(index).update()
        document.storeToURL("file://" + docx, (value("FilterName", "MS Word 2007 XML"),))
        document.storeToURL("file://" + pdf, (value("FilterName", "writer_pdf_Export"),))
        print("pages", document.getCurrentController().PageCount)
        document.close(True)
    finally:
        office.terminate()
        shutil.rmtree(profile, ignore_errors=True)


if __name__ == "__main__":
    main()
