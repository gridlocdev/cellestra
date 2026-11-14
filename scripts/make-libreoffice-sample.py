"""Build a sample .ods saved by LibreOffice itself, for the reader's real-file tests.

    python3 scripts/make-libreoffice-sample.py [out.ods]

This writes a flat OpenDocument spreadsheet (.fods) by hand, with formulas but no results:
several sheets, every value type, number formats, styles, merges, formulas across sheets,
named ranges, a note, a validation, an autofilter, hidden rows and columns. (No frozen panes:
a headless conversion saves no view settings.)
A headless LibreOffice then opens it, calculates every formula and saves it as .ods, so the
fixture is LibreOffice's own output: its styles, settings, cached values and package layout.
The default output is packages/core/test/serialization/fixtures/libreoffice-sample.ods.

    python3 scripts/make-libreoffice-sample.py --check <file.ods>

opens any .ods (one Cellestra saved, say) in LibreOffice and prints every sheet as CSV of the
shown text, so LibreOffice's view of two files can be compared with `diff`.

It only runs LibreOffice's `soffice --convert-to`, never LibreOffice's bundled Python, which
crashes on start on recent macOS. Each run uses a throwaway profile.
"""

import glob
import os
import shutil
import subprocess
import sys
import tempfile
from xml.sax.saxutils import escape

SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_OUT = os.path.join(ROOT, "packages/core/test/serialization/fixtures/libreoffice-sample.ods")

NS = " ".join(
    'xmlns:%s="%s"' % pair
    for pair in [
        ("office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"),
        ("style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0"),
        ("text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0"),
        ("table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0"),
        ("fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"),
        ("number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"),
        ("svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"),
        ("dc", "http://purl.org/dc/elements/1.1/"),
        ("of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2"),
    ]
)

STYLES = """
<office:automatic-styles>
  <number:currency-style style:name="N-usd-pos" style:volatile="true">
    <number:currency-symbol number:language="en" number:country="US">$</number:currency-symbol>
    <number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/>
  </number:currency-style>
  <number:currency-style style:name="N-usd">
    <style:text-properties fo:color="#ff0000"/>
    <number:text>-</number:text>
    <number:currency-symbol number:language="en" number:country="US">$</number:currency-symbol>
    <number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/>
    <style:map style:condition="value()&gt;=0" style:apply-style-name="N-usd-pos"/>
  </number:currency-style>
  <number:date-style style:name="N-date">
    <number:year number:style="long"/><number:text>-</number:text>
    <number:month number:style="long"/><number:text>-</number:text>
    <number:day number:style="long"/>
  </number:date-style>
  <number:percentage-style style:name="N-pct">
    <number:number number:decimal-places="1" number:min-decimal-places="1" number:min-integer-digits="1"/>
    <number:text>%</number:text>
  </number:percentage-style>
  <number:boolean-style style:name="N-bool"><number:boolean/></number:boolean-style>

  <style:style style:name="co-wide" style:family="table-column"><style:table-column-properties style:column-width="1.65in"/></style:style>
  <style:style style:name="co-notes" style:family="table-column"><style:table-column-properties style:column-width="1.97in"/></style:style>
  <style:style style:name="co" style:family="table-column"><style:table-column-properties style:column-width="0.89in"/></style:style>
  <style:style style:name="ro-tall" style:family="table-row"><style:table-row-properties style:row-height="0.35in" style:use-optimal-row-height="false"/></style:style>
  <style:style style:name="ro" style:family="table-row"><style:table-row-properties style:use-optimal-row-height="true"/></style:style>

  <style:style style:name="title" style:family="table-cell" style:parent-style-name="Default">
    <style:table-cell-properties fo:background-color="#1f4e79" style:vertical-align="middle"/>
    <style:paragraph-properties fo:text-align="center"/>
    <style:text-properties fo:font-size="16pt" fo:font-weight="bold" fo:color="#ffffff"/>
  </style:style>
  <style:style style:name="head" style:family="table-cell" style:parent-style-name="Default">
    <style:table-cell-properties fo:background-color="#ddebf7" fo:border-bottom="0.74pt solid #000000"/>
    <style:text-properties fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="bold" style:family="table-cell" style:parent-style-name="Default">
    <style:text-properties fo:font-weight="bold"/>
  </style:style>
  <style:style style:name="money" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N-usd"/>
  <style:style style:name="date" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N-date"/>
  <style:style style:name="pct" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N-pct"/>
  <style:style style:name="bool" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N-bool"/>
  <style:style style:name="wrap" style:family="table-cell" style:parent-style-name="Default">
    <style:table-cell-properties fo:wrap-option="wrap"/>
  </style:style>
  <style:style style:name="italic" style:family="table-cell" style:parent-style-name="Default">
    <style:text-properties fo:font-style="italic" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/>
  </style:style>
</office:automatic-styles>
"""


def text(value, style=None, extra=""):
    attrs = ' table:style-name="%s"' % style if style else ""
    lines = "".join("<text:p>%s</text:p>" % escape(line) for line in value.split("\n"))
    return '<table:table-cell%s%s office:value-type="string">%s</table:table-cell>' % (attrs, extra, lines)


def num(value, style=None, kind="float"):
    attrs = ' table:style-name="%s"' % style if style else ""
    if kind == "date":
        return '<table:table-cell%s office:value-type="date" office:date-value="%s"/>' % (attrs, value)
    if kind == "boolean":
        return '<table:table-cell%s office:value-type="boolean" office:boolean-value="%s"/>' % (
            attrs,
            "true" if value else "false",
        )
    if kind == "percentage":
        return '<table:table-cell%s office:value-type="percentage" office:value="%s"/>' % (attrs, value)
    if kind == "currency":
        return '<table:table-cell%s office:value-type="currency" office:currency="USD" office:value="%s"/>' % (
            attrs,
            value,
        )
    return '<table:table-cell%s office:value-type="float" office:value="%s"/>' % (attrs, value)


def formula(f, style=None):
    """A formula with no result: LibreOffice works it out when it opens the file."""
    attrs = ' table:style-name="%s"' % style if style else ""
    return '<table:table-cell%s table:formula="of:=%s"/>' % (attrs, escape(f, {'"': "&quot;"}))


EMPTY = "<table:table-cell/>"


def row(cells, style="ro", extra=""):
    return '<table:table-row table:style-name="%s"%s>%s</table:table-row>' % (style, extra, "".join(cells))


ITEMS = [
    ("Widget", 12, 3.5, "2026-03-01", 0.1, True, "Ships in pairs"),
    ("Gadget", 3, 19.99, "2026-03-04", 0, False, ""),
    ("Zoë's sprocket", 40, 0.25, "2026-03-09", 0.05, True, "Line one\nLine two"),
    ("Bolt, M6", 250, 0.04, "2026-02-27", 0, True, 'Says "hi"'),
    ("Ünïcødé ☃", 1, 1234.5, "2026-01-15", 0.25, False, "<b>not html</b> & co"),
    ("=not a formula", 7, 2, "2026-03-31", 0.125, True, "   leading spaces"),
]
FIRST = 3  # first data row, 1-based
LAST = FIRST + len(ITEMS) - 1  # last data row
TOTAL = LAST + 1


def data_sheet():
    out = ['<table:table table:name="Data">']
    out.append('<table:table-column table:style-name="co-wide" table:default-cell-style-name="Default"/>')
    out.append('<table:table-column table:style-name="co" table:number-columns-repeated="6"/>')
    out.append('<table:table-column table:style-name="co-notes"/>')
    out.append('<table:table-column table:style-name="co" table:visibility="collapse"/>')
    title = '<table:table-cell table:style-name="title" table:number-columns-spanned="8" office:value-type="string"><text:p>Inventory — März 2026 😀</text:p></table:table-cell>'
    out.append(row([title] + ["<table:covered-table-cell/>"] * 7, "ro-tall"))
    heads = ["Item", "Qty", "Price", "Received", "Discount", "In stock", "Total", "Notes"]
    out.append(row([text(h, "head") for h in heads]))
    for i, (name, qty, price, received, discount, stock, note) in enumerate(ITEMS):
        r = FIRST + i
        cells = [
            text(name, extra=' table:content-validation-name="val1"'),
            num(qty),
            num(price, "money", "currency"),
            num(received, "date", "date"),
            num(discount, "pct", "percentage"),
            num(stock, "bool", "boolean"),
            formula("ROUND([.B%d]*[.C%d]*(1-[.E%d]);2)" % (r, r, r), "money"),
        ]
        if r == FIRST + 1:
            # A note on the Qty cell instead of the plain number.
            cells[1] = (
                '<table:table-cell office:value-type="float" office:value="%s">'
                "<office:annotation><dc:creator>Sample</dc:creator><text:p>Back-ordered until April</text:p></office:annotation>"
                "</table:table-cell>" % qty
            )
        cells.append(text(note, "wrap") if "\n" in note else (text(note) if note else EMPTY))
        cells.append(text("hidden helper") if i == 0 else EMPTY)
        out.append(row(cells))
    out.append(
        row(
            [
                text("Total", "bold"),
                formula("SUM([.B%d:.B%d])" % (FIRST, LAST)),
                EMPTY,
                EMPTY,
                formula("AVERAGE([.E%d:.E%d])" % (FIRST, LAST), "pct"),
                formula("COUNTIF([.F%d:.F%d];TRUE())" % (FIRST, LAST)),
                formula("SUBTOTAL(9;[.G%d:.G%d])" % (FIRST, LAST), "money"),
                text("italic, underlined", "italic"),
            ]
        )
    )
    out.append(row([EMPTY]))
    out.append(row([text("hidden row")], extra=' table:visibility="collapse"'))
    out.append("</table:table>")
    return "".join(out)


LOOKUPS = [
    ("VLOOKUP", 'VLOOKUP("Gadget";[$Data.A%d:.G%d];3;0)'),
    ("INDEX/MATCH", 'INDEX([$Data.G%d:.G%d];MATCH("Widget";[$Data.A%d:.A%d];0))'),
    ("Named range", "SUM(Prices)"),
    ("IF", 'IF([$Data.B4]>10;"plenty";"few")'),
    ("Concatenate", 'CONCATENATE([$Data.A3];" × ";[$Data.B3])'),
    ("TEXT", 'TEXT([$Data.D3];"DD MMM YYYY")'),
    ("Date math", "[$Data.D5]-[$Data.D3]"),
    ("YEAR", "YEAR([$Data.D3])"),
    ("Quoted sheet", "SUM([$'Q2 Sales'.B2:.B4])"),
    ("Div by zero", "1/0"),
    ("IFERROR", 'IFERROR(1/0;"none")'),
    ("Not found", 'VLOOKUP("Nope";[$Data.A%d:.B%d];2;0)'),
    ("ROUND", "ROUND(2/3;4)"),
    ("Power", "2^10"),
    ("Comparison", "[$Data.C3]>[$Data.C4]"),
    ("Text as number", '"12"+1'),
    ("SUMPRODUCT", "SUMPRODUCT([$Data.B%d:.B%d];[$Data.C%d:.C%d])"),
    ("COUNTA", "COUNTA([$Data.A%d:.A%d])"),
    ("MAX", "MAX([$Data.C%d:.C%d])"),
    ("LEN", "LEN([$Data.A5])"),
    ("UPPER", "UPPER([$Data.A3])"),
    ("Absolute mix", "[$Data.$B3]+[$Data.B$3]"),
    ("Percent of total", "[$Data.G3]/[$Data.G%d]"),
    ("3D sum", "SUM([$Data.B3:$'Q2 Sales'.B4])"),
    ("3D count", "COUNT([$Data.A1:$'Q2 Sales'.H9])"),
    ("3D max", "MAX([$Data.C3:$Lookup.C8])"),
    ("INDEX area", "INDEX(([$Data.A3:.B4]~[$Data.G3:.H4]);2;1;2)"),
    ("Intersection", "SUM([$Data.A3:.C5]![$Data.B4:.D8])"),
    ("No overlap", "SUM([.A1]![.B2])"),
]


def lookup_sheet():
    out = ['<table:table table:name="Lookup">']
    out.append('<table:table-column table:style-name="co-wide" table:number-columns-repeated="2"/>')
    out.append(row([text("What", "head"), text("Result", "head")]))
    for label, f in LOOKUPS:
        n = f.count("%d")
        if label == "Percent of total":
            f = f % TOTAL
        elif n:
            f = f % tuple([FIRST, LAST] * (n // 2))
        out.append(row([text(label), formula(f, "pct" if label == "Percent of total" else None)]))
    out.append("</table:table>")
    return "".join(out)


def q2_sheet():
    out = ['<table:table table:name="Q2 Sales">']
    out.append(row([text("Month", "head"), text("Sales", "head")]))
    for month, sales in [("April", 1200), ("May", 950.5), ("June", 1410.25)]:
        out.append(row([text(month), num(sales, "money", "currency")]))
    out.append(row([EMPTY]))
    out.append(row([text("Share of Data total"), formula("[.B2]/[$Data.G%d]" % TOTAL, "pct")]))
    out.append("</table:table>")
    return "".join(out)


def extras():
    return (
        "<table:content-validations>"
        '<table:content-validation table:name="val1" table:condition="of:cell-content-is-in-list(&quot;Widget&quot;;&quot;Gadget&quot;;&quot;Bolt, M6&quot;)" table:allow-empty-cell="true" table:base-cell-address="$Data.A3" table:display-list="unsorted"/>'
        "</table:content-validations>"
    )


def named():
    # No formula uses TaxRate: Cellestra doesn't evaluate named expressions yet, only keeps them on save.
    return (
        "<table:named-expressions>"
        '<table:named-range table:name="Prices" table:base-cell-address="$Data.$A$1" table:cell-range-address="$Data.$C$%d:.$C$%d"/>'
        '<table:named-expression table:name="TaxRate" table:base-cell-address="$Data.$A$1" table:expression="of:=0.2"/>'
        "</table:named-expressions>"
        "<table:database-ranges>"
        '<table:database-range table:name="Inventory" table:target-range-address="Data.A2:Data.H%d" table:display-filter-buttons="true"/>'
        "</table:database-ranges>"
    ) % (FIRST, LAST, LAST)


def flat_document():
    return (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<office:document %s office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.spreadsheet">'
        "%s<office:body><office:spreadsheet>%s%s%s%s"
        '<table:table table:name="Empty"><table:table-column/><table:table-row><table:table-cell/></table:table-row></table:table>'
        "%s</office:spreadsheet></office:body></office:document>\n"
    ) % (NS, STYLES, extras(), data_sheet(), lookup_sheet(), q2_sheet(), named())


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
        source = os.path.join(work, "libreoffice-sample.fods")
        with open(source, "w", encoding="utf-8") as f:
            f.write(flat_document())
        soffice(os.path.join(work, "profile"), "--convert-to", "ods", "--outdir", work, source)
        made = os.path.join(work, "libreoffice-sample.ods")
        if not os.path.exists(made):
            sys.exit("LibreOffice wrote no .ods")
        os.makedirs(os.path.dirname(out), exist_ok=True)
        shutil.copy(made, out)
        print("wrote", out)
    finally:
        shutil.rmtree(work, ignore_errors=True)


def check(path):
    """Every sheet as CSV of the shown text, as LibreOffice sees the file after recalculating."""
    work = tempfile.mkdtemp(prefix="cellestra-lo-")
    try:
        source = os.path.join(work, "in.ods")
        shutil.copy(path, source)
        # 44,34,76: comma, double quote, UTF-8; 1: first line; then formatted text as shown,
        # and -1 exports every sheet to its own file.
        csv = "csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,true,false,false,-1"
        soffice(os.path.join(work, "profile"), "--convert-to", csv, "--outdir", work, source)
        for name in sorted(glob.glob(os.path.join(work, "in-*.csv"))):
            sheet = os.path.basename(name)[3:-4]
            print("== %s" % sheet)
            sys.stdout.write(open(name, encoding="utf-8").read())
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    args = sys.argv[1:]
    if args[:1] == ["--check"]:
        check(os.path.abspath(args[1]))
    else:
        make(os.path.abspath(args[0] if args else DEFAULT_OUT))
