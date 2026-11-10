def generate_report(PPTX_FILE, JSON_FILE, REPORT_MONTH, PHOTO_ZIP=None, SAFETY_ZERO_MEANS_MISSING=True):
    import os, io, re, json, zipfile, copy, calendar, unicodedata
    from pathlib import Path
    from collections import defaultdict
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
    from PIL import Image

    with open(JSON_FILE,encoding='utf-8-sig') as fp:
        raw=json.load(fp)
    assert isinstance(raw,list),'Expected a JSON list of Power BI rows.'
    month_rows = [d for d in raw if str(d.get('MH_BI[Month]','')).startswith(REPORT_MONTH)]
    if not month_rows:
        raise ValueError(f'No data for {REPORT_MONTH}. Available: {sorted({str(d.get("MH_BI[Month]",""))[:7] for d in raw})}')

    # Exported overall field C_V 2 reflects mean of S1–S5 in this supplied dataset.
    # S6 entries are all zero; by default interpret as unreported rather than verified failed audits.
    SKEYS=[f'MH_BI[Value S{i}]' for i in range(1,7)]
    OVERALL='MH_BI[C_V 2]'
    SEGMENT='MH_BI[Segment ]'

    def norm(s):
        s=unicodedata.normalize('NFKD',str(s)).encode('ascii','ignore').decode('ascii').lower()
        return re.sub(r'[^a-z0-9]','',s)

    def pct(v):
        return 'N/A' if v is None else f'{float(v)*100:.2f} %'

    def safety_value(d):
        val=d.get(SKEYS[5]);return None if (SAFETY_ZERO_MEANS_MISSING and (val is None or float(val)==0)) else val

    def overall(d):
        v=d.get(OVERALL)
        return float(v) if v is not None else sum(float(d[k]) for k in SKEYS[:5])/5

    def avg(xs):
        vals=[float(v) for v in xs if v is not None]
        return sum(vals)/len(vals) if vals else None

    pr=Presentation(PPTX_FILE)
    assert len(pr.slides)>=5, 'Template must have at least five slides.'
    # Use actual source slide objects as master templates for duplication.
    cover, overview, ranking, detail, closing=list(pr.slides)[:5]
    month_label=calendar.month_name[int(REPORT_MONTH[-2:])].upper()+' '+REPORT_MONTH[:4]

    def set_text(shape,new):
        # Updating the first existing run retains the template styling.
        if not shape.has_text_frame:return
        paras=shape.text_frame.paragraphs
        if paras and paras[0].runs:
            paras[0].runs[0].text=str(new)
            for p in paras:
                for r in (p.runs[1:] if p==paras[0] else p.runs):r.text=''
        else:shape.text=str(new)

    def find_text_shape(slide, exact):
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text.strip()==exact:return shape
        return None

    def duplicate_slide(deck, source):
        # Clone all shapes incl. template's colored background and charts.
        dest=deck.slides.add_slide(source.slide_layout)
        for sh in source.shapes:
            clone=copy.deepcopy(sh.element)
            dest.shapes._spTree.insert_element_before(clone,'p:extLst')
        # Rebind image relationships (each cloned picture references source slide's rId)
        for rel in source.part.rels.values():
            if rel.reltype.endswith('/image') or rel.reltype.endswith('/hyperlink'):
                dest.part.relate_to(rel.target_part if not rel.is_external else rel.target_ref,rel.reltype,is_external=rel.is_external)
        return dest

    def move_slide(deck, slide, index):
        lst=deck.slides._sldIdLst
        sid=next(e for e in lst if e.rId==slide.part._partname) if False else None
        for e in list(lst):
            if deck.part.related_part(e.rId)==slide.part:
                sid=e;break
        if sid is None:raise RuntimeError('Slide not found')
        lst.remove(sid);lst.insert(index,sid)

    def remove_shape(sh):
        el=sh._element;el.getparent().remove(el)

    def set_month(slide):
        for sh in slide.shapes:
            if sh.has_text_frame and '[MONTH] [YEAR]' in sh.text:set_text(sh,sh.text.replace('[MONTH] [YEAR]',month_label))

    def create_text(slide,x,y,w,h,value,size=15,bold=False,color=(255,255,255),align=PP_ALIGN.LEFT):
        tb=slide.shapes.add_textbox(x,y,w,h);tf=tb.text_frame;tf.clear();tf.word_wrap=True
        tf.margin_left=tf.margin_right=Inches(.07);tf.margin_top=tf.margin_bottom=0
        tf.vertical_anchor=MSO_ANCHOR.MIDDLE
        p=tf.paragraphs[0];p.alignment=align
        r=p.add_run();r.text=str(value);r.font.size=Pt(size);r.font.bold=bold;r.font.color.rgb=RGBColor(*color)
        return tb

    # Optional photo ZIP: only filenames and image bytes; do not access OneDrive automatically.
    zip_path=PHOTO_ZIP
    photo_by_seg=defaultdict(list)
    if zip_path:
        z=zipfile.ZipFile(zip_path)
        for info in z.infolist():
            if info.is_dir() or info.file_size>25_000_000:continue
            path=Path(info.filename)
            if path.suffix.lower() not in ('.jpg','.jpeg','.png','.webp'):continue
            parts=list(path.parts)
            # Skip photos under a different YYYY-MM folder
            datefolders=[p for p in parts if re.fullmatch(r'20\d\d-\d\d',p)]
            if datefolders and REPORT_MONTH not in datefolders:continue
            matching=[str(d[SEGMENT]) for d in month_rows if norm(d[SEGMENT]) in [norm(p) for p in parts[:-1]]]
            if matching:photo_by_seg[matching[0]].append((info.filename,z.read(info)))
        for k in photo_by_seg:photo_by_seg[k].sort()

    # Fill OVERVIEW — mean per available segment (not weighted by audit count).
    set_month(overview)
    for i,v in enumerate([avg([d[SKEYS[k]] for d in month_rows]) for k in range(5)]):
        pass
    shapes=[sh for sh in overview.shapes if sh.has_text_frame and sh.text.strip()=='[XX.XX %]']
    # Template order: S1, S2, S3, S4, S6, overall, S5.
    if len(shapes)!=7:raise ValueError(f'Expected seven score boxes on slide 2, found {len(shapes)}')
    values=[avg([d[SKEYS[k]] for d in month_rows]) for k in range(5)]
    s6=avg([safety_value(d) for d in month_rows])
    for sh,val in zip(shapes,[values[0],values[1],values[2],values[3],s6,avg([overall(d) for d in month_rows]),values[4]]):set_text(sh,pct(val))

    # Ranking slide: native editable table instead of placeholder.
    set_month(ranking)
    ph=find_text_shape(ranking,'[POWER BI SEGMENT RESULTS]')
    if ph:remove_shape(ph)
    ordered=sorted(month_rows,key=overall,reverse=True)
    nrows=len(ordered)+1
    left,top,width=Inches(.78),Inches(2.28),Inches(11.7)
    rowheight=min(.39,4.18/nrows)
    height=Inches(rowheight*nrows)
    tb=ranking.shapes.add_table(nrows,3,left,top,width,height).table
    tb.columns[0].width=Inches(1.3);tb.columns[1].width=Inches(6.3);tb.columns[2].width=Inches(4.1)
    for i,record in enumerate([('RANK','SEGMENT','OVERALL (S1–S5)')]+[(str(k),str(d[SEGMENT]),pct(overall(d))) for k,d in enumerate(ordered,1)]):
        for col,value in enumerate(record):
            cell=tb.cell(i,col)
            cell.text=str(value)
            cell.fill.solid()
            cell.fill.fore_color.rgb=RGBColor(13,90,140) if i==0 else RGBColor(10,44,83) if i%2 else RGBColor(13,60,103)
            cell.margin_left=Inches(.12);cell.margin_top=Inches(.015);cell.margin_bottom=Inches(.015)
            cell.vertical_anchor=MSO_ANCHOR.MIDDLE
            for p in cell.text_frame.paragraphs:
                p.alignment=PP_ALIGN.CENTER if col in (0,2) else PP_ALIGN.LEFT
                for r in p.runs:
                    r.font.size=Pt(14 if nrows<11 else 12);r.font.bold=(i==0);r.font.color.rgb=RGBColor(255,255,255)

    # Populate detail template for the first segment, clone it for further segments.
    def fill_detail(slide,d,photos,offset=0):
        set_month(slide)
        title=find_text_shape(slide,'[SEGMENT] — RESULTS')
        if title:set_text(title,f'{d[SEGMENT]} — RESULTS'+(f' (PHOTOS {offset+1}–{min(offset+6,len(photos))})' if offset else ''))
        scoreboxes=[sh for sh in slide.shapes if sh.has_text_frame and sh.text.strip()=='[XX.XX %]']
        if len(scoreboxes)!=7:raise ValueError(f'Expected 7 segment score boxes, found {len(scoreboxes)}')
        for sh,val in zip(scoreboxes,[overall(d)]+[d[k] for k in SKEYS[:5]]+[safety_value(d)]):set_text(sh,pct(val))
        slots=[sh for sh in slide.shapes if sh.has_text_frame and sh.text.startswith('PHOTO ')]
        for pos,slot in enumerate(slots):
            x,y,w,h=slot.left,slot.top,slot.width,slot.height
            remove_shape(slot)
            idx=offset+pos
            if idx>=len(photos):
                r=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,x,y,w,h)
                r.fill.solid();r.fill.fore_color.rgb=RGBColor(12,52,87)
                r.line.color.rgb=RGBColor(51,122,168)
                create_text(slide,x,y,w,h,'NO PHOTO',12,False,(176,196,212),PP_ALIGN.CENTER)
                continue
            fname,b=photos[idx]
            try:
                im=Image.open(io.BytesIO(b)).convert('RGB')
                bio=io.BytesIO();im.save(bio,'JPEG',quality=88);bio.seek(0)
                pic=slide.shapes.add_picture(bio,x,y,width=w,height=h)
                ratio_img=im.width/im.height;ratio_box=w/h
                if ratio_img>ratio_box:
                    portion=ratio_box/ratio_img;pic.crop_left=(1-portion)/2;pic.crop_right=(1-portion)/2
                else:
                    portion=ratio_img/ratio_box;pic.crop_top=(1-portion)/2;pic.crop_bottom=(1-portion)/2
            except Exception as err:
                create_text(slide,x,y,w,h,'INVALID PHOTO',12,color=(255,220,220),align=PP_ALIGN.CENTER)

    # Rebuild slide sequence explicitly: cover, overview, ranking, all segment slides, closing.
    # Remove original detail slide; clones from the unchanged source copy.
    for d in ordered:
        photos=photo_by_seg.get(d[SEGMENT],[])
        for off in range(0,max(1,len(photos)),6):
            s=duplicate_slide(pr,detail)
            fill_detail(s,d,photos,off)
            # Clone copies source image relationships; duplicate is appended, so move before closing.
            move_slide(pr,s,len(pr.slides)-2)
    # Remove the old unfilled slide 4.
    old_id=pr.slides._sldIdLst[3]
    pr.part.drop_rel(old_id.rId);pr.slides._sldIdLst.remove(old_id)
    # Confirm the closing slide is last; adjust if template had more slides.
    assert len(pr.slides)==4+sum(max(1,(len(photo_by_seg.get(d[SEGMENT],[]))+5)//6) for d in ordered)
    out=io.BytesIO()
    pr.save(out)
    return out.getvalue(), len(ordered), sum(map(len,photo_by_seg.values())), len(pr.slides)
