"""Build a .ods saved by LibreOffice itself that uses what Cellestra keeps but does not edit.

    python3 scripts/make-libreoffice-outline.py [out.ods]

The file has workbook structure protection, print ranges, repeated header rows and columns
(print titles), nested row groups (one collapsed), a column group, manual row and column page
breaks, a Basic library, a document event (Open Document) and a sheet event (Activate Document),
each bound to a macro. The round-trip test saves it with Cellestra and compares the element and
attribute names before and after, so content dropped on save shows up as a failing test.

A headless LibreOffice converts a hand-written flat file (.fods) to .ods. LibreOffice drops
event bindings it reads from a flat file, so a second run opens the .ods with a Basic macro in
the throwaway profile, binds the two events and saves it again. Like
make-libreoffice-sample.py, it never runs LibreOffice's bundled Python.
The default output is packages/core/test/serialization/fixtures/libreoffice-outline.ods.
"""

import os
import shutil
import subprocess
import sys
import tempfile
from xml.sax.saxutils import escape

SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(ROOT, "packages/core/test/serialization/fixtures/libreoffice-outline.ods")

FLAT = """<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.spreadsheet">
 <office:scripts>
  <office:script script:language="ooo:Basic">
   <ooo:libraries>
    <ooo:library-embedded ooo:name="Standard">
     <ooo:module ooo:name="Module1" ooo:language="StarBasic">
      <ooo:source-code>Sub OnOpen
End Sub
Sub OnFocus
End Sub
</ooo:source-code>
     </ooo:module>
    </ooo:library-embedded>
   </ooo:libraries>
  </office:script>
 </office:scripts>
 <office:automatic-styles>
  <style:style style:name="co1" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="2.258cm"/></style:style>
  <style:style style:name="co2" style:family="table-column"><style:table-column-properties fo:break-before="page" style:column-width="2.258cm"/></style:style>
  <style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="0.452cm" fo:break-before="auto" style:use-optimal-row-height="true"/></style:style>
  <style:style style:name="ro2" style:family="table-row"><style:table-row-properties style:row-height="0.452cm" fo:break-before="page" style:use-optimal-row-height="true"/></style:style>
 </office:automatic-styles>
 <office:body>
  <office:spreadsheet table:structure-protected="true">
   <table:table table:name="Data" table:print-ranges="Data.A1:Data.E12">
    <table:table-header-columns>
     <table:table-column table:style-name="co1"/>
    </table:table-header-columns>
    <table:table-column-group>
     <table:table-column table:style-name="co1" table:number-columns-repeated="2"/>
    </table:table-column-group>
    <table:table-column table:style-name="co2"/>
    <table:table-column table:style-name="co1"/>
    <table:table-header-rows>
     %(header)s
    </table:table-header-rows>
    <table:table-row-group>
     %(north)s
     <table:table-row-group table:display="false">
      %(north_east)s
      %(north_west)s
     </table:table-row-group>
    </table:table-row-group>
    %(south)s
    %(east)s
    %(west)s
    <table:table-row table:style-name="ro1"><table:table-cell table:formula="of:=SUBTOTAL(9;[.B2:.B7])" office:value-type="float"/></table:table-row>
   </table:table>
  </office:spreadsheet>
 </office:body>
</office:document>
"""


def text(value):
    return '<table:table-cell office:value-type="string"><text:p>%s</text:p></table:table-cell>' % escape(value)


def num(value):
    return '<table:table-cell office:value-type="float" office:value="%s"/>' % value


def row(cells, style="ro1", extra=""):
    return '<table:table-row table:style-name="%s"%s>%s</table:table-row>' % (style, extra, "".join(cells))


def flat_document():
    hidden = ' table:visibility="collapse"'
    return FLAT % {
        "header": row([text(t) for t in ("Region", "Q1", "Q2", "Q3", "Q4")]),
        "north": row([text("North"), num(1), num(2)]),
        "north_east": row([text("North-East"), num(3)], extra=hidden),
        "north_west": row([text("North-West"), num(4)], extra=hidden),
        "south": row([text("South"), num(5)]),
        "east": row([text("East"), num(6)], style="ro2"),
        "west": row([text("West"), num(7)]),
    }


def bind_events_macro(source, target):
    """A Basic module that opens `source`, binds the two events and saves it as `target`."""
    url = lambda path: "file://" + path  # noqa: E731
    lines = [
        "Sub BindEvents",
        "  Dim load(0) As New com.sun.star.beans.PropertyValue",
        '  load(0).Name = "Hidden"',
        "  load(0).Value = True",
        '  doc = StarDesktop.loadComponentFromURL("%s", "_blank", 0, load())' % url(source),
        '  doc.Events.replaceByName("OnLoad", Binding("OnOpen"))',
        '  doc.Sheets.getByIndex(0).Events.replaceByName("OnFocus", Binding("OnFocus"))',
        "  Dim store(0) As New com.sun.star.beans.PropertyValue",
        '  store(0).Name = "FilterName"',
        '  store(0).Value = "calc8"',
        '  doc.storeToURL("%s", store())' % url(target),
        "  doc.close(True)",
        "End Sub",
        "",
        "Function Binding(macro As String)",
        "  Dim event(1) As New com.sun.star.beans.PropertyValue",
        '  event(0).Name = "EventType"',
        '  event(0).Value = "Script"',
        '  event(1).Name = "Script"',
        '  event(1).Value = "vnd.sun.star.script:Standard.Module1." & macro & "?language=Basic&location=document"',
        "  Binding = event()",
        "End Function",
    ]
    return (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
        '<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" '
        'script:language="StarBasic">%s</script:module>\n' % escape("\n".join(lines), {'"': "&quot;"})
    )


def soffice(profile, *args):
    subprocess.run(
        [SOFFICE, "--headless", "--nologo", "--norestore", "-env:UserInstallation=file://" + profile, *args],
        check=True,
        timeout=300,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def make(out):
    work = tempfile.mkdtemp(prefix="cellestra-lo-")
    try:
        profile = os.path.join(work, "profile")
        source = os.path.join(work, "outline.fods")
        with open(source, "w", encoding="utf-8") as f:
            f.write(flat_document())
        soffice(profile, "--convert-to", "ods", "--outdir", work, source)
        converted = os.path.join(work, "outline.ods")
        if not os.path.exists(converted):
            sys.exit("LibreOffice wrote no .ods")
        # The first run created the profile; its Standard library now gets the binding macro.
        module = os.path.join(profile, "user/basic/Standard/Module1.xba")
        with open(module, "w", encoding="utf-8") as f:
            f.write(bind_events_macro(converted, os.path.join(work, "bound.ods")))
        soffice(profile, "macro:///Standard.Module1.BindEvents")
        made = os.path.join(work, "bound.ods")
        if not os.path.exists(made):
            sys.exit("LibreOffice did not save the file with its events")
        os.makedirs(os.path.dirname(out), exist_ok=True)
        shutil.copy(made, out)
        print("wrote", out)
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    args = sys.argv[1:]
    make(os.path.abspath(args[0]) if args else DEFAULT_OUT)
