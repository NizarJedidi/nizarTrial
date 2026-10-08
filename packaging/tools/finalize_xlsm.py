"""Assemble le .xlsm : classeur .xlsx + vbaProject.bin + boutons de formulaire « Remplir ».

Usage : python3 finalize_xlsm.py in.xlsx vbaProject.bin out.xlsm
"""
import re, sys, zipfile

xlsx_in, vba_bin, xlsm_out = sys.argv[1], sys.argv[2], sys.argv[3]

# boutons : (nom de feuille, macro, colonne (0-based), largeur px, marge gauche pt)
BUTTONS = [
    ("Packaging Calculator", "RemplirPackagingCalculator", 3, 110, 384.0),
    ("Mixed Pack Helper", "RemplirMixedPackHelper", 9, 110, 1044.0),
]
EMU_PX = 9525

VML = '''<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
 <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="{idmap}"/></o:shapelayout>
 <v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe">
  <v:stroke joinstyle="miter"/>
  <v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>
  <o:lock v:ext="edit" shapetype="t"/>
 </v:shapetype>
 <v:shape id="_x0000_s{shapeid}" type="#_x0000_t201" style='position:absolute;margin-left:{left}pt;margin-top:0;width:{wpt}pt;height:30pt;z-index:1;mso-wrap-style:tight' o:button="t" fillcolor="buttonFace [67]" strokecolor="windowText [64]" o:insetmode="auto">
  <v:fill color2="buttonFace [67]" o:detectmouseclick="t"/>
  <v:textbox style='mso-direction-alt:auto' o:singleclick="f">
   <div style='text-align:center'><font face="Calibri" size="220" color="#000000">Remplir</font></div>
  </v:textbox>
  <x:ClientData ObjectType="Button">
   <x:Anchor>{col}, 0, 0, 0, {col}, {wpx}, 2, 0</x:Anchor>
   <x:PrintObject>False</x:PrintObject>
   <x:AutoFill>False</x:AutoFill>
   <x:FmlaMacro>[0]!{macro}</x:FmlaMacro>
   <x:TextHAlign>Center</x:TextHAlign>
   <x:TextVAlign>Center</x:TextVAlign>
  </x:ClientData>
 </v:shape>
</xml>'''

CTRL = ('<legacyDrawing r:id="rIdVml"/>'
        '<controls><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
        '<mc:Choice Requires="x14"><control shapeId="{shapeid}" r:id="rIdCtrl" name="Button 1">'
        '<controlPr defaultSize="0" print="0" autoFill="0" autoPict="0" macro="[0]!{macro}">'
        '<anchor moveWithCells="1" sizeWithCells="1">'
        '<from><xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></from>'
        '<to><xdr:col>{col}</xdr:col><xdr:colOff>{wemu}</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></to>'
        '</anchor></controlPr></control></mc:Choice></mc:AlternateContent></controls>')

CTRLPROP = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="Button" lockText="1"/>')

SHEET_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              '<Relationship Id="rIdVml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing{n}.vml"/>'
              '<Relationship Id="rIdCtrl" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp" Target="../ctrlProps/ctrlProp{n}.xml"/>'
              '</Relationships>')

zin = zipfile.ZipFile(xlsx_in)
parts = {n: zin.read(n) for n in zin.namelist()}

# --- feuille -> fichier
wb = parts["xl/workbook.xml"].decode("utf-8")
rels = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
rid_target = dict(re.findall(r'<Relationship[^>]*Target="([^"]+)"[^>]*Id="([^"]+)"', rels))
rid_target = {v: k for k, v in rid_target.items()}
if not rid_target:
    rid_target = {m[0]: m[1] for m in re.findall(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels)}
sheet_file = {}
for name, rid in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb):
    t = rid_target[rid]
    sheet_file[name] = t.lstrip("/") if t.startswith("/") else "xl/" + t

# --- VBA : content types + relation
ct = parts["[Content_Types].xml"].decode("utf-8")
ct = ct.replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
ct = ct.replace("</Types>", '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>')
if 'Extension="vml"' not in ct:
    ct = ct.replace("<Default Extension=", '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Default Extension=', 1)
rels = rels.replace("</Relationships>", '<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>')
parts["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
parts["xl/vbaProject.bin"] = open(vba_bin, "rb").read()

# --- boutons
for n, (sheet, macro, col, wpx, left) in enumerate(BUTTONS, start=1):
    f = sheet_file[sheet]
    xml = parts[f].decode("utf-8")
    root_m = re.search(r"<worksheet[^>]*>", xml)
    root = root_m.group(0)
    new_root = root
    for prefix, uri in (("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
                        ("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")):
        if f"xmlns:{prefix}=" not in root:
            new_root = new_root[:-1] + f' xmlns:{prefix}="{uri}">'
    xml = xml.replace(root, new_root, 1)
    assert "<legacyDrawing" not in xml and "<tableParts" not in xml and "<extLst" not in xml
    shapeid = 1024 * n + 1
    xml = xml.replace("</worksheet>", CTRL.format(shapeid=shapeid, macro=macro, col=col, wemu=wpx * EMU_PX) + "</worksheet>")
    parts[f] = xml.encode("utf-8")
    base = f.rsplit("/", 1)[1]
    parts[f"xl/worksheets/_rels/{base}.rels"] = SHEET_RELS.format(n=n).encode("utf-8")
    parts[f"xl/drawings/vmlDrawing{n}.vml"] = VML.format(idmap=n, shapeid=shapeid, left=left, wpt=wpx * 0.75, col=col, wpx=wpx, macro=macro).encode("utf-8")
    parts[f"xl/ctrlProps/ctrlProp{n}.xml"] = CTRLPROP.encode("utf-8")
    ct = ct.replace("</Types>", f'<Override PartName="/xl/ctrlProps/ctrlProp{n}.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/></Types>')
parts["[Content_Types].xml"] = ct.encode("utf-8")

order = ["[Content_Types].xml", "_rels/.rels"] + [n for n in parts if n not in ("[Content_Types].xml", "_rels/.rels")]
with zipfile.ZipFile(xlsm_out, "w", zipfile.ZIP_DEFLATED) as zout:
    for n in order:
        zout.writestr(n, parts[n])
print("ecrit", xlsm_out, "feuilles:", sheet_file)
