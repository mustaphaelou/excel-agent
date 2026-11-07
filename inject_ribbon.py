#!/usr/bin/env python3
"""
Inject the ribbon (customUI14) part into the .xlam that Excel produced.

Excel's object model cannot create ribbon XML, so it is grafted into the OPC
package afterwards: one part, one relationship in _rels/.rels.

The content type is the subtle part. Per [MS-OI29500] 3.4.1.3 the CustomUI part
is `application/xml` - i.e. it is covered by the package's existing
`<Default Extension="xml" ContentType="application/xml"/>`. Giving it the
vendor type `application/vnd.ms-office.customUI+xml` makes Excel silently ignore
the part: no error, no tab. So: no Override, and strip one if a previous run
added it.

Usage: inject_ribbon.py <raw.xlam> <out.xlam> <customUI14.xml>
"""
import os
import re
import sys
import zipfile

RIB_PART = "customUI/customUI14.xml"
REL_ID = "rIdExcelAgentUI"
REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"


def main() -> int:
    if len(sys.argv) != 4:
        print(__doc__)
        return 2
    src, dst, ui_path = sys.argv[1:4]
    ui_xml = open(ui_path, encoding="utf-8").read()

    with zipfile.ZipFile(src) as zin:
        entries = [(i, zin.read(i.filename)) for i in zin.infolist()]

    names = [i.filename for i, _ in entries]
    if "[Content_Types].xml" not in names or "_rels/.rels" not in names:
        print("ERROR: package is missing OPC plumbing")
        return 1

    # never carry a stale copy of the part from the source file
    entries = [(i, d) for i, d in entries if i.filename != RIB_PART]

    out: list[tuple[zipfile.ZipInfo, bytes]] = []
    for info, data in entries:
        if info.filename in ("[Content_Types].xml", "_rels/.rels"):
            text = data.decode("utf-8")

            if info.filename == "[Content_Types].xml":
                # drop any Override for this part (see module docstring)
                text = re.sub(
                    r'\s*<Override[^>]*PartName="/?' + re.escape(RIB_PART) + r'"[^>]*/>',
                    "", text)
                # make sure .xml parts default to application/xml, which is the
                # content type the CustomUI part is expected to have
                if not re.search(r'<Default[^>]*Extension="xml"', text):
                    text = text.replace(
                        "</Types>",
                        '<Default Extension="xml" ContentType="application/xml"/>\n</Types>')

            if info.filename == "_rels/.rels":
                # drop a previous relationship, then add ours once
                text = re.sub(
                    r'\s*<Relationship[^>]*Target="/?' + re.escape(RIB_PART) + r'"[^>]*/>',
                    "", text)
                text = text.replace(
                    "</Relationships>",
                    f'  <Relationship Id="{REL_ID}" Type="{REL_TYPE}" '
                    f'Target="{RIB_PART}"/>\n</Relationships>')

            data = text.encode("utf-8")
        out.append((info, data))

    out.append((zipfile.ZipInfo(RIB_PART, date_time=(2024, 1, 1, 0, 0, 0)),
                ui_xml.encode("utf-8")))

    # Content_Types has to be the first entry in a well-formed OPC package
    out.sort(key=lambda e: (e[0].filename != "[Content_Types].xml",
                            e[0].filename != "_rels/.rels"))

    tmp = dst + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in out:
            zi = zipfile.ZipInfo(info.filename,
                                 date_time=info.date_time or (2024, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o600 << 16
            zout.writestr(zi, data)
    os.replace(tmp, dst)

    print(f"ribbon injected: {RIB_PART} (content type application/xml)")
    print(f"  entries: {len(out)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
