from pathlib import Path
import json, argparse
from copy import deepcopy
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN, MSO_AUTO_SIZE
from pptx.shapes.placeholder import PicturePlaceholder
from pptx.parts.slide import SlideLayoutPart
from pptx.opc.packuri import PackURI
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

parser=argparse.ArgumentParser(description='Build the editable 56-page A4 travel magazine.')
parser.add_argument('--output',type=Path,default=Path(__file__).resolve().parent.parent)
args=parser.parse_args()
ROOT = args.output.resolve()
ROOT.mkdir(parents=True, exist_ok=True)
(ROOT/'Source').mkdir(exist_ok=True)
W, H = 595.276, 841.89
PAPER, INK, BLUE, RUST = 'F7F2E9', '182E37', '243F85', 'BC4F35'
SAND, MIST, MUTED, LINE, WHITE = 'E6D8C3', 'E0E7EF', '647278', 'C8C6BB', 'FFFFFF'
FONT_FILES = {
    'serif': '/usr/share/fonts/truetype/noto/NotoSerifDisplay-Regular.ttf',
    'italic': '/usr/share/fonts/truetype/noto/NotoSerifDisplay-Italic.ttf',
    'sans': '/usr/share/fonts/truetype/open-sans/OpenSans-Regular.ttf',
    'bold': '/usr/share/fonts/truetype/open-sans/OpenSans-Bold.ttf',
}
for key, file in FONT_FILES.items():
    bundled=ROOT/'Fonts'/Path(file).name
    if bundled.exists(): FONT_FILES[key]=str(bundled)
for key, file in FONT_FILES.items():
    pdfmetrics.registerFont(TTFont(key, file))
prs = Presentation()
prs.slide_width, prs.slide_height = Pt(W), Pt(H)
prs.core_properties.title = 'The Days Between — A Personal Travel Magazine'
prs.core_properties.subject = 'Editable A4 travel journal / 52-page magazine + 4 alternate covers'
prs.core_properties.author = 'Independent editorial template'
prs.core_properties.keywords = 'travel, journal, magazine, editable, A4, Italy'
records, metrics, slides = [], [], []

def rgb(value): return RGBColor.from_string(value)

def rect(s, x, y, w, h, fill, stroke=None, sw=.65):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = rgb(fill)
    if stroke:
        sh.line.color.rgb = rgb(stroke); sh.line.width = Pt(sw)
    else: sh.line.fill.background()
    for effect in sh.element.xpath('p:style/a:effectRef'): effect.set('idx','0')
    return sh

def oval(s, x, y, w, h, fill, stroke=None):
    sh = s.shapes.add_shape(MSO_SHAPE.OVAL, Pt(x), Pt(y), Pt(w), Pt(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = rgb(fill)
    if stroke: sh.line.color.rgb=rgb(stroke);sh.line.width=Pt(.8)
    else: sh.line.fill.background()
    for effect in sh.element.xpath('p:style/a:effectRef'): effect.set('idx','0')
    return sh

def rule(s, x, y, w, color=LINE, thickness=.6):
    return rect(s,x,y,w,thickness,color)

def wrap(text, font, size, width):
    result = []
    for para in text.split('\n'):
        words=para.split()
        if not words:
            result.append('');continue
        line=words[0]
        for word in words[1:]:
            candidate=line+' '+word
            if pdfmetrics.stringWidth(candidate,font,size) <= width:
                line=candidate
            else: result.append(line);line=word
        result.append(line)
    return result

def text(s, value, x, y, w, size=11, font='sans', color=INK,
         h=None, leading=None, align='left'):
    leading = leading or size*1.36
    lines=wrap(value,font,size,w-2)
    h=h or max(leading*len(lines)+5, size+5)
    sh=s.shapes.add_textbox(Pt(x),Pt(y),Pt(w),Pt(h))
    tf=sh.text_frame;tf.clear()
    tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=0
    tf.word_wrap=False;tf.auto_size=MSO_AUTO_SIZE.NONE
    tf.vertical_anchor=MSO_ANCHOR.TOP
    for i,line in enumerate(lines):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.text=line;p.space_before=Pt(0);p.space_after=Pt(0);p.line_spacing=Pt(leading)
        p.alignment={'left':PP_ALIGN.LEFT,'right':PP_ALIGN.RIGHT,'center':PP_ALIGN.CENTER}[align]
        p.font.name='Noto Serif Display' if font in ('serif','italic') else 'Open Sans'
        p.font.size=Pt(size);p.font.color.rgb=rgb(color)
        p.font.italic=font=='italic';p.font.bold=font=='bold'
    records[-1]['texts'].append(value)
    metrics.append(dict(page=len(slides),text=value[:65],x=x,y=y,w=w,h=h,
                        needed=leading*len(lines),font=font,size=size))
    sh.name='TEXT / '+value.replace('\n',' ')[:70]
    return sh

def label(s,value,x=48,y=40,w=499,color=MUTED):
    return text(s,value.upper(),x,y,w,8.1,'bold',color,leading=11)

def page(section, title, bg=PAPER, number=True):
    # Each page has its own layout. Picture replacement inherits geometry from
    # this layout, as required by PowerPoint's picture-placeholder mechanism.
    seed=prs.slide_layouts[6]
    element=deepcopy(seed._element)
    element.set('preserve','1')
    element.cSld.set('name',f'{len(slides)+1:02} / {title}')
    for child in list(element.spTree):
        if child.tag not in (qn('p:nvGrpSpPr'),qn('p:grpSpPr')):
            element.spTree.remove(child)
    part=SlideLayoutPart(PackURI(f'/ppt/slideLayouts/slideLayout{len(slides)+12}.xml'),
                         seed.part.content_type,prs.part.package,element)
    part.relate_to(prs.slide_master.part,RT.SLIDE_MASTER)
    rid=prs.slide_master.part.relate_to(part,RT.SLIDE_LAYOUT)
    ids=prs.slide_master._element.get_or_add_sldLayoutIdLst()
    entry=OxmlElement('p:sldLayoutId')
    entry.set('id',str(max(int(e.get('id')) for e in ids)+1))
    entry.set(qn('r:id'),rid);ids.append(entry)
    s=prs.slides.add_slide(part.slide_layout);slides.append(s)
    records.append(dict(page=len(slides),section=section,title=title,texts=[],photos=[]))
    s.background.fill.solid();s.background.fill.fore_color.rgb=rgb(bg)
    if number:
        light=bg in (BLUE,INK,RUST)
        c=PAPER if light else MUTED
        label(s,'THE DAYS BETWEEN   /   '+section,color=c)
        rule(s,48,786,499, c)
        label(s,'A PERSONAL TRAVEL JOURNAL',48,799,340,c)
        text(s,f'{len(slides):02}',510,796,37,10,'bold',c,align='right')
    s.notes_slide.notes_text_frame.text = (
        f'{len(slides):02} / {title}\n'
        'Original editable layout. Replace sample copy and photo placeholders with your own material. '
        'Picture placeholders support insert_picture() and PowerPoint picture insertion. '
        'Use Open Sans and Noto Serif Display. A4 portrait, 210 x 297 mm. '
        'See the included guide for editing, alternative covers, export and rights notes.'
    )
    return s

def photo(s,x,y,w,h,caption='YOUR PHOTO',tone=MIST):
    idx=100+len(records[-1]['photos'])
    sp=s.shapes._spTree.add_placeholder(s.shapes._next_shape_id,
        f'PHOTO / {caption}',PP_PLACEHOLDER.PICTURE,'horz','full',idx)
    sh=PicturePlaceholder(sp,s.shapes)
    sh.left,sh.top,sh.width,sh.height=Pt(x),Pt(y),Pt(w),Pt(h)
    sh.fill.solid();sh.fill.fore_color.rgb=rgb(tone)
    sh.line.fill.background()
    # Text belongs to the placeholder itself and disappears on picture insertion.
    tf=sh.text_frame;tf.clear()
    tf.margin_left=tf.margin_right=Pt(10)
    tf.margin_top=tf.margin_bottom=Pt(5)
    tf.vertical_anchor=MSO_ANCHOR.MIDDLE
    p=tf.paragraphs[0];p.text=caption.upper()
    p.alignment=PP_ALIGN.CENTER
    p.font.name='Open Sans';p.font.size=Pt(8);p.font.color.rgb=rgb(MUTED)
    # Preserve the same geometry in the layout so inserted pictures stay put.
    layout_sp=deepcopy(sh.element)
    layout_sp.nvSpPr.cNvPr.id=s.slide_layout.shapes._next_shape_id
    s.slide_layout.shapes._spTree.insert_element_before(layout_sp,'p:extLst')
    records[-1]['photos'].append(dict(idx=idx,x=x,y=y,w=w,h=h,caption=caption))
    return sh

def caption(s,value,x,y,w):
    text(s,value,x,y,w,8.5,'sans',MUTED,leading=12)

def heading(s,kicker,title,subtitle=None):
    label(s,kicker,48,89,color=RUST)
    text(s,title,48,110,499,34,'serif',leading=39)
    if subtitle: text(s,subtitle,48,200,465,11,color=MUTED)

def field(s,title,prompt,x,y,w=230):
    label(s,title,x,y,w,RUST)
    text(s,prompt,x,y+21,w,11,leading=16)

def artwork(s,x,y,w,h,style=0):
    # Original editable geometric compositions, not traced from the reference.
    rect(s,x,y,w,h,SAND)
    if style%3==0:
        rect(s,x+w*.55,y,w*.45,h,MIST)
        oval(s,x+w*.70,y+h*.09,w*.16,w*.16,RUST)
        rect(s,x,y+h*.70,w,h*.30,BLUE)
        rect(s,x+w*.10,y+h*.32,w*.28,h*.48,PAPER)
        oval(s,x+w*.10,y+h*.22,w*.28,w*.28,PAPER)
        rect(s,x+w*.17,y+h*.40,w*.14,h*.40,INK)
        oval(s,x+w*.17,y+h*.33,w*.14,w*.14,INK)
        for i in range(4):
            rect(s,x+w*(.41+i*.1),y+h*(.65-i*.08),w*.11,h*(.15+i*.08),RUST)
        for i in range(3):rule(s,x+w*.08,y+h*(.87+i*.04),w*.84,PAPER,.9)
    elif style%3==1:
        rect(s,x,y,w*.35,h,RUST)
        oval(s,x+w*.48,y+h*.12,w*.25,w*.25,PAPER)
        rect(s,x+w*.35,y+h*.58,w*.65,h*.42,BLUE)
        for i in range(5):
            rect(s,x+w*(.12+i*.12),y+h*(.60-i*.08),w*.10,h*(.40+i*.08),PAPER)
        rule(s,x+w*.39,y+h*.85,w*.55,SAND,1)
        rule(s,x+w*.39,y+h*.90,w*.55,SAND,1)
    else:
        rect(s,x,y,w,h,BLUE)
        oval(s,x+w*.60,y+h*.12,w*.23,w*.23,SAND)
        for i in range(6):
            rect(s,x+w*(.08+i*.14),y+h*(.64-.035*i),w*.09,h*(.36+.035*i),PAPER)
        rect(s,x,y+h*.83,w,h*.17,RUST)
        rule(s,x+w*.06,y+h*.90,w*.88,SAND,1)

def cover(variant=0):
    s=page('COVER','Cover' if not variant else f'Alternative cover {variant}',number=False)
    if variant==0:
        label(s,'ITALY   /   A PERSONAL TRAVEL MAGAZINE',40,30,515,BLUE)
        text(s,'The days\nbetween',40,72,515,62,'serif',leading=64)
        artwork(s,40,241,515,405,0)
        text(s,'A place. A pace.\nA story to keep.',40,674,310,24,'italic',leading=29)
        label(s,'YOUR NAMES',395,680,160,BLUE)
        text(s,'MONTH / YEAR\nISSUE 01',395,704,155,10,color=MUTED,leading=16)
        rule(s,40,791,515,INK)
        label(s,'SIX DAYS OF DISCOVERY   /   ONE SHARED JOURNEY',40,804,515,INK)
    elif variant==1:
        rect(s,0,0,W,H,BLUE)
        artwork(s,304,0,W-304,570,1)
        label(s,'ITALY / ISSUE 01',35,35,210,PAPER)
        text(s,'The\ndays\nbetween',35,108,255,53,'serif',PAPER,leading=63)
        rule(s,35,621,525,PAPER)
        text(s,'Our own kind\nof away.',35,649,370,32,'italic',PAPER,leading=38)
        label(s,'YOUR NAMES / MONTH YEAR',35,780,510,PAPER)
    else:
        label(s,'A PERSONAL TRAVEL MAGAZINE',48,38,499,BLUE)
        text(s,'The days between',48,96,499,39,'serif')
        rule(s,48,162,499,BLUE)
        photo(s,48,198,499,414,'COVER IMAGE',SAND)
        rect(s,48,641,110,110,RUST)
        text(s,'01',59,654,90,48,'serif',PAPER,align='center')
        text(s,'ITALY',183,649,355,37,'serif')
        text(s,'YOUR NAMES\nMONTH / YEAR',185,710,350,10,color=MUTED,leading=16)
        label(s,'PLACES, PEOPLE & THE DETAILS IN BETWEEN',48,796,499,BLUE)

cover()
s=page('OPENING','A journey, kept in small details',BLUE)
text(s,'A journey,\nkept in\nsmall details.',48,112,485,48,'serif',PAPER,leading=57)
artwork(s,48,342,499,282,2)
text(s,'For the streets we wandered, the tables we shared,\nand the moments that made us pause.',48,655,470,13,'sans',PAPER,leading=20)
label(s,'THIS COPY BELONGS TO / YOUR NAMES',48,738,499,PAPER)

s=page('OPENING','A note before we begin')
heading(s,'01 / OPENING NOTES','A note before\nwe begin.')
photo(s,360,226,187,278,'PORTRAIT OR DETAIL',SAND)
text(s,'Dear future us,',48,231,280,21,'italic',BLUE)
text(s,
     'This is where our journey becomes a story. Not a list of every place we visited, '
     'but a collection of the things we want to remember.\n\n'
     'Replace this letter with your own introduction. Write about why you chose this trip, '
     'what you hoped to find and who was there beside you. Leave room for the surprises.\n\n'
     'What did this journey change about the way you see a place, each other or yourself?',
     48,278,276,11,leading=17)
label(s,'YOUR SIGNATURE',48,650,270,RUST)
text(s,'Names / Date',48,675,320,21,'italic')
caption(s,'A face, a doorway, a first impression.',360,516,187)

s=page('OPENING','The journey at a glance')
heading(s,'THE DETAILS','The journey\nat a glance.')
for title,prompt,x,y in [
    ('DESTINATION','Italy / add your region',48,233),
    ('TRAVELLERS','Your names',310,233),
    ('DATES','Arrival — Departure',48,325),
    ('OUR BASE','Town / neighbourhood / stay',310,325),
    ('HOW WE TRAVELLED','Train, road, sea or sky',48,417),
    ('THE REASON','What brought us here',310,417)]:
    field(s,title,prompt,x,y,237);rule(s,x,y+69,237)
rect(s,48,547,499,179,SAND)
label(s,'THREE WORDS FOR THIS TRIP',68,569,459,RUST)
text(s,'Unhurried.\nCurious. Together.',68,598,450,30,'serif',leading=36)

s=page('CONTENTS','Inside this issue')
heading(s,'THE READING ROUTE','Inside this issue.')
sections=[
 ('Opening notes','Letters, details & a first impression',3),
 ('The way there','Route, preparation & a moodboard',6),
 ('Six days, six stories','A daily journal in words and pictures',10),
 ('At the table','Flavours, favourite tables & addresses',28),
 ('What we wore','A small wardrobe, many memories',32),
 ('Ways of seeing','Architecture, traditions & collected words',36),
 ('Objects with a story','Keepsakes, tickets & little treasures',40),
 ('The photo edit','A gallery of light, colour & connection',44),
 ('From here, with love','Postcards and a closing reflection',48)]
for i,(title,sub,n) in enumerate(sections):
    y=195+i*59
    text(s,f'{n:02}',48,y,52,20,'serif',RUST)
    text(s,title,115,y,425,17,'serif')
    caption(s,sub,115,y+27,425);rule(s,48,y+51,499)
caption(s,'Main magazine: 01–52. Alternative covers and backs: 53–56.',48,746,499)

s=page('THE WAY THERE','A route to remember')
heading(s,'02 / THE WAY THERE','A route to\nremember.')
rect(s,69,237,1,446,BLUE)
for i,(name,detail) in enumerate([
 ('Arrival','Where the journey began'),
 ('First stop','A town we wanted to get lost in'),
 ('A change of pace','A quieter place along the way'),
 ('By the water','A view worth staying for'),
 ('The way home','The last place we paused')]):
    y=236+i*94
    oval(s,62,y+4,15,15,RUST)
    label(s,f'{i+1:02} / ADD PLACE',101,y,420,RUST)
    text(s,name,101,y+21,420,21,'serif')
    caption(s,detail+' / date / transport',101,y+53,420)
label(s,'REORDER THE STOPS TO MATCH YOUR OWN JOURNEY',48,743,499,BLUE)

s=page('THE WAY THERE','Three places, one journey')
heading(s,'A SENSE OF PLACE','Three places,\none journey.')
for i,(t,desc) in enumerate([
 ('The city','A street, a square, an unexpected turn.'),
 ('The coast','The sound, the light, the colour of the water.'),
 ('The quieter place','Somewhere we could slow down.')]):
    y=229+i*171
    photo(s,48,y,185,145,t,[MIST,SAND,MIST][i])
    label(s,f'0{i+1} / PLACE NAME',255,y+7,292,RUST)
    text(s,t,255,y+31,290,23,'serif')
    text(s,desc,255,y+72,282,11,leading=16)

s=page('BEFORE WE GO','Room for the essentials')
heading(s,'BEFORE WE GO','Room for\nthe essentials.')
for col,(title,items) in enumerate([
 ('The practical things',['Travel documents','Bookings and tickets','Charging essentials','Everyday medication','A reusable bottle','Comfortable shoes']),
 ('The things that matter',['A book for the journey','Something to write in','A camera or phone','One favourite outfit','Room for a small souvenir','Time with no fixed plan'])]):
    x=48+col*262
    text(s,title,x,234,237,22,'serif',BLUE)
    for j,item in enumerate(items):
        y=294+j*51
        rect(s,x,y+3,10,10,PAPER,LINE)
        text(s,item,x+24,y,210,10.5)
        rule(s,x,y+35,237)
rect(s,48,644,499,91,SAND)
label(s,'OUR OWN ADDITION',65,661,460,RUST)
text(s,'One thing we never travel without: __________________',65,686,460,11)

s=page('BEFORE WE GO','The feeling we followed')
heading(s,'THE MOODBOARD','The feeling\nwe followed.')
photo(s,48,225,310,325,'THE PLACE',SAND)
photo(s,377,225,170,153,'A COLOUR')
photo(s,377,397,170,153,'A DETAIL')
for i,c in enumerate([BLUE,RUST,SAND,INK]):
    rect(s,48+i*129,577,112,39,c)
text(s,'A palette, a texture, a feeling.',48,643,499,26,'italic',BLUE)
text(s,'Collect the images that capture how you imagined this journey.\nAfter the trip, return to this page and notice what changed.',
     48,693,499,11,leading=17)

days=[
 ('Arrival & first impressions','The first pause','The moment the journey became real.'),
 ('Finding our rhythm','A slower morning','The small routine we would happily keep.'),
 ('Taking the long way','An unexpected turn','The unplanned stop that became a favourite.'),
 ('A place to linger','Staying a little longer','The view, table or street we did not want to leave.'),
 ('Following our curiosity','Something new','A taste, a conversation or a different way of seeing.'),
 ('One more look','Before we leave','What we noticed when we knew we were going home.')]

for i,(title,story,prompt) in enumerate(days):
    day=i+1
    s=page(f'DAY {day:02}',title)
    label(s,f'03 / FIELD NOTES   —   DAY {day:02}',48,89,color=RUST)
    text(s,f'{day:02}',43,112,140,80,'serif',BLUE,leading=88)
    text(s,title,215,119,332,30,'serif',leading=36)
    label(s,'PLACE / DATE / WEATHER',215,207,332)
    if i%2==0:
        photo(s,48,253,499,342,'THE DAY IN ONE IMAGE',SAND if i%3==0 else MIST)
        text(s,prompt,48,625,499,23,'italic',BLUE,leading=29)
        text(s,'Write a few lines about where the day began, what drew you in,\nand the detail you would like to remember.',48,703,499,11,leading=17)
    else:
        photo(s,48,253,288,455,'THE DAY IN ONE IMAGE',MIST)
        text(s,prompt,361,262,186,23,'italic',BLUE,leading=29)
        label(s,'A MOMENT TO KEEP',361,451,186,RUST)
        text(s,'Add a short note about the sound, colour or feeling that belongs to this day.',361,481,186,11,leading=17)
        caption(s,'Place / time / who was there',48,722,499)
    s=page(f'DAY {day:02}',story)
    label(s,f'DAY {day:02} / THE STORY',48,89,color=RUST)
    text(s,story,48,112,499,34,'serif',leading=40)
    if i%3==0:
        photo(s,48,197,200,318,'ONE SMALL DETAIL',SAND)
        text(s,'Start with a scene.',278,201,269,22,'italic',BLUE)
        text(s,'Describe the place before you explain what happened. What could you hear? '
             'Where did the light fall? Who was with you?\n\n'
             'Let one ordinary detail lead into the story. Add a conversation, a surprise '
             'or a moment that shifted the mood of the day.',278,249,269,11,leading=17)
        rule(s,48,558,499)
        field(s,'WHAT SURPRISED US','Replace this with your favourite unexpected moment.',48,583,230)
        field(s,'WHAT WE WOULD REPEAT','A place, an activity or simply a slower pace.',310,583,237)
        text(s,'“Add a sentence you want to carry home.”',48,705,499,20,'italic',RUST)
    elif i%3==1:
        photo(s,48,194,499,237,'A VIEW FROM THE DAY',SAND)
        text(s,'The part we keep returning to',48,464,499,22,'italic',BLUE)
        text(s,'Write about a moment that stayed with you after the day was over. '
             'Give it a setting and a small detail that makes it yours.\n\n'
             'It might be a table in the shade, a wrong turn or a conversation that lasted longer than expected.',
             48,511,235,11,leading=17)
        text(s,'Then tell the rest of the story in your own words. What did you learn? '
             'What changed your plans? What made you laugh?\n\n'
             'Keep this space personal. A short, honest memory is enough.',
             312,511,235,11,leading=17)
    else:
        text(s,'“Give this day\nits own words.”',48,196,250,31,'italic',BLUE,leading=39)
        photo(s,327,194,220,270,'A PORTRAIT OR DETAIL',SAND)
        text(s,'Use this page for the story behind the pictures. Describe how the day unfolded, '
             'from the first plan to the moment that changed it.\n\n'
             'Include a detail that photographs cannot hold: the temperature, a scent, '
             'a fragment of conversation or how you felt.',48,505,236,11,leading=17)
        text(s,'Finish with one thing you would tell your future self about this day. '
             'It does not need to be a big lesson; it can be as simple as remembering to look up.\n\n'
             'Replace these prompts with your own words.',312,505,235,11,leading=17)
    s=page(f'DAY {day:02}','Small scenes')
    label(s,f'DAY {day:02} / THE PHOTO NOTES',48,89,color=RUST)
    text(s,'Small scenes.',48,111,499,35,'serif')
    if i%2==0:
        for j,(x,y,w,h) in enumerate([(48,198,240,219),(307,198,240,219),(48,463,161,232),(228,463,319,232)]):
            photo(s,x,y,w,h,['A PLACE','A PERSON','A DETAIL','A MOMENT'][j],SAND if j==2 else MIST)
            caption(s,f'{j+1:02} / Add a short caption.',x,y+h+10,w)
    else:
        photo(s,48,197,270,518,'THE SCENE',SAND)
        for j in range(3):
            photo(s,337,197+j*181,210,133,f'DETAIL {j+1:02}')
            caption(s,'Add a short caption.',337,338+j*181,210)
    label(s,'PLACE / DATE / ONE THING THE PHOTO CANNOT SHOW',48,748,499,BLUE)

s=page('AT THE TABLE','At the table',BLUE)
label(s,'04 / TASTES & TABLES',48,91,color=PAPER)
text(s,'At the\ntable.',48,119,499,62,'serif',PAPER,leading=70)
photo(s,245,337,302,350,'A FAVOURITE TABLE',SAND)
text(s,'Some days\nare remembered\nby their flavour.',48,383,177,25,'italic',PAPER,leading=32)
label(s,'TASTE / PLACE / COMPANY',48,735,499,PAPER)

s=page('AT THE TABLE','Flavours to remember')
heading(s,'THE TASTING NOTES','Flavours to\nremember.')
for i,(x,y) in enumerate([(48,228),(310,228),(48,492),(310,492)]):
    photo(s,x,y,237,153,f'DISH {i+1:02}',SAND if i%2 else MIST)
    text(s,['A first sip','Something shared','The local favourite','A sweet ending'][i],x,y+168,237,19,'serif')
    caption(s,'Name / place / what made it memorable',x,y+203,237)

s=page('AT THE TABLE','A table worth keeping')
heading(s,'ONE GOOD MEAL','A table worth\nkeeping.')
photo(s,48,228,275,402,'THE MEAL OR THE SETTING',SAND)
field(s,'WHERE WE ATE','Restaurant / town',349,235,198)
field(s,'WHAT WE ORDERED','Dish / drink / something shared',349,332,198)
field(s,'WHY IT STAYED WITH US','Write about the company, the atmosphere or an unexpected flavour.',349,444,198)
text(s,'More than what was on the plate.',48,667,499,25,'italic',BLUE)
caption(s,'Use the space above for a personal memory, not a restaurant review.',48,725,499)

s=page('AT THE TABLE','For another visit')
heading(s,'THE LITTLE ADDRESS BOOK','For another visit.')
text(s,'Keep the places you would recommend to a friend.',48,171,499,11,color=MUTED)
for i,t in enumerate(['Coffee','Lunch','Dinner','Something sweet','A market','Hidden gem']):
    y=232+i*80
    text(s,f'{i+1:02}',48,y,45,20,'serif',RUST)
    text(s,t,110,y,150,17,'serif')
    text(s,'Name / neighbourhood',285,y,262,11)
    caption(s,'Address or note / add a booking tip if useful',110,y+31,437)
    rule(s,48,y+62,499)

s=page('WHAT WE WORE','What we wore')
label(s,'05 / THE TRAVEL WARDROBE',48,90,color=RUST)
text(s,'What\nwe wore.',48,117,499,58,'serif',leading=65)
photo(s,48,293,290,419,'A FAVOURITE LOOK',SAND)
photo(s,360,293,187,210,'A SMALL DETAIL')
text(s,'Clothes that\nwent places.',360,538,187,25,'italic',BLUE,leading=32)
caption(s,'The pieces we reached for again and again.',360,650,187)

s=page('WHAT WE WORE','A small, useful wardrobe')
heading(s,'THE CAPSULE EDIT','A small, useful\nwardrobe.')
for i in range(6):
    x=48+(i%3)*173;y=233+(i//3)*258
    photo(s,x,y,153,189,f'PIECE {i+1:02}',SAND if i%2 else MIST)
    caption(s,['The everyday layer','The easy top','The walk-all-day pair','The evening piece','The useful bag','The finishing touch'][i],x,y+202,153)

s=page('WHAT WE WORE','Two moments, two looks')
heading(s,'THE OUTFIT NOTES','Two moments,\ntwo looks.')
for i in range(2):
    x=48+i*262
    photo(s,x,231,237,354,['DAYTIME','AFTER DARK'][i],SAND if i else MIST)
    text(s,['Out in the light','After the sun'][i],x,608,237,22,'serif')
    text(s,'Where / what we wore / why it worked',x,653,237,10.5,color=MUTED,leading=16)

s=page('WHAT WE WORE','The colours of our days')
heading(s,'COLOUR & TEXTURE','The colours\nof our days.')
for i,(c,name) in enumerate([(BLUE,'Deep blue'),(RUST,'Sun-warmed clay'),(SAND,'Soft stone'),(INK,'Evening shade')]):
    y=229+i*120
    rect(s,48,y,94,94,c)
    text(s,name,166,y+6,381,23,'serif')
    text(s,'Replace with a colour you noticed and a memory it brings back.',166,y+47,381,11,color=MUTED,leading=17)

s=page('WAYS OF SEEING','Ways of seeing')
label(s,'06 / CULTURE & CURIOSITY',48,90,color=RUST)
text(s,'Ways of\nseeing.',48,117,499,56,'serif',leading=64)
artwork(s,48,289,499,322,1)
text(s,'Look beyond the landmark.',48,645,499,28,'italic',BLUE)
text(s,'Keep the everyday details: a doorway, a ritual,\na local phrase, a different rhythm of life.',48,701,499,11,leading=17)

s=page('WAYS OF SEEING','Look a little closer')
heading(s,'ARCHITECTURE & DETAILS','Look a little\ncloser.')
photo(s,48,230,245,433,'THE WHOLE SCENE',SAND)
photo(s,313,230,234,196,'DETAIL 01')
photo(s,313,467,234,196,'DETAIL 02')
caption(s,'A doorway, a facade or a quiet corner.',48,678,245)
caption(s,'What drew your eye? Add the details here.',313,678,234)
label(s,'PLACE / MATERIAL / COLOUR / LIGHT',48,741,499,BLUE)

s=page('WAYS OF SEEING','A different daily rhythm')
heading(s,'EVERYDAY RITUALS','A different\ndaily rhythm.')
photo(s,48,229,499,251,'A LOCAL MOMENT',SAND)
text(s,'What we noticed',48,513,235,23,'serif',BLUE)
text(s,'Describe something ordinary that felt new to you: a morning routine, '
     'a gathering place or the pace of an evening.\n\n'
     'Write from your own experience, with curiosity and care.',48,558,235,11,leading=17)
text(s,'What we took home',312,513,235,23,'serif',BLUE)
text(s,'Was there a habit you wanted to bring into your everyday life?\n\n'
     'Use this space for a small personal reflection rather than a general claim about a culture.',
     312,558,235,11,leading=17)

s=page('WAYS OF SEEING','Words we collected')
heading(s,'HEARD ALONG THE WAY','Words we\ncollected.')
for i in range(4):
    x=48+(i%2)*262;y=235+(i//2)*247
    rect(s,x,y,237,220,SAND if i in (0,3) else MIST)
    label(s,f'{i+1:02} / WORD OR PHRASE',x+17,y+18,203,RUST)
    text(s,'Your words\nhere.',x+17,y+54,203,28,'serif',BLUE,leading=34)
    text(s,'Meaning / who said it /\nwhere we heard it',x+17,y+150,203,10,color=MUTED,leading=16)

s=page('OBJECTS WITH A STORY','Objects with a story',RUST)
label(s,'07 / THE THINGS WE KEPT',48,91,color=PAPER)
text(s,'Objects\nwith a story.',48,122,499,52,'serif',PAPER,leading=61)
photo(s,48,308,499,303,'A KEEPSAKE',SAND)
text(s,'Small enough to bring home.\nBig enough to bring it all back.',48,656,499,25,'italic',PAPER,leading=32)

s=page('OBJECTS WITH A STORY','The little collection')
heading(s,'OUR PERSONAL ARCHIVE','The little\ncollection.')
for i in range(6):
    x=48+(i%3)*173;y=232+(i//3)*257
    photo(s,x,y,153,182,f'OBJECT {i+1:02}',SAND if i%2 else MIST)
    caption(s,'What / where / why we kept it',x,y+195,153)

s=page('OBJECTS WITH A STORY','One object, a whole story')
heading(s,'A CLOSER LOOK','One object,\na whole story.')
photo(s,48,229,499,289,'THE OBJECT',SAND)
field(s,'FOUND IN','Place / date / how it came to us',48,553,236)
field(s,'WHAT IT HOLDS','The memory attached to this object',312,553,235)
text(s,'“Add the sentence that explains why it matters.”',48,686,499,22,'italic',BLUE)

s=page('OBJECTS WITH A STORY','Paper traces')
heading(s,'TICKETS, NOTES & LITTLE FINDS','Paper traces.')
text(s,'A place for the things we almost threw away.',48,169,499,11,color=MUTED)
photo(s,48,216,499,142,'TICKET / RECEIPT / NOTE',SAND)
photo(s,48,400,240,226,'A PAPER FIND')
photo(s,307,400,240,226,'ANOTHER PAPER FIND',SAND)
text(s,'The ordinary evidence of a good day.',48,665,499,25,'italic',BLUE)
caption(s,'Photograph or scan your keepsakes; remove personal booking details before sharing.',48,723,499)

s=page('THE PHOTO EDIT','Light, held still',BLUE)
label(s,'08 / THE PHOTO EDIT',48,89,color=PAPER)
text(s,'Light,\nheld still.',48,120,499,55,'serif',PAPER,leading=64)
photo(s,48,288,499,390,'YOUR FAVOURITE PHOTOGRAPH',SAND)
text(s,'A few frames. A whole feeling.',48,717,499,22,'italic',PAPER)

s=page('THE PHOTO EDIT','Nine frames')
heading(s,'THE CONTACT SHEET','Nine frames.')
for i in range(9):
    x=48+(i%3)*173;y=197+(i//3)*177
    photo(s,x,y,153,137,f'{i+1:02}',SAND if i%3==1 else MIST)
    caption(s,'Place / moment',x,y+145,153)
label(s,'CHOOSE A SEQUENCE: COLOUR, TIME OR FEELING',48,748,499,BLUE)

s=page('THE PHOTO EDIT','Two sides of the same day')
heading(s,'A VISUAL CONVERSATION','Two sides of\nthe same day.')
photo(s,48,231,237,420,'FIRST IMAGE',SAND)
photo(s,310,231,237,420,'SECOND IMAGE')
text(s,'Near / far. Quiet / busy. Before / after.',48,685,499,21,'italic',BLUE)
caption(s,'Choose a pair of photographs that say more together.',48,736,499)

s=page('THE PHOTO EDIT','The frame we would print')
label(s,'THE ONE WE WOULD FRAME',48,89,color=RUST)
photo(s,48,132,499,486,'ONE IMAGE TO KEEP',SAND)
text(s,'The frame we\nwould print.',48,651,320,31,'serif',leading=37)
text(s,'Place / date\nA memory in one sentence.',390,663,157,10,color=MUTED,leading=17)

s=page('FROM HERE, WITH LOVE','A postcard from here')
heading(s,'09 / FROM HERE, WITH LOVE','A postcard\nfrom here.')
photo(s,48,228,499,299,'THE POSTCARD VIEW',SAND)
label(s,'TO SOMEONE WE WISH WERE HERE',48,562,499,RUST)
text(s,'Dear __________,',48,589,499,23,'italic',BLUE)
text(s,'Describe this place in a way only you could. Include a small detail,\na wish or a moment you would have loved to share.',48,643,499,11,leading=17)
text(s,'With love, __________',48,716,499,21,'italic')

s=page('FROM HERE, WITH LOVE','A note to our future selves')
heading(s,'FOR ANOTHER DAY','A note to our\nfuture selves.')
text(s,'When we read this again…',48,238,499,28,'italic',BLUE)
for i,(label_,prompt) in enumerate([
    ('REMEMBER THIS','A scene we never want to forget.'),
    ('KEEP DOING THIS','Something this journey taught us to make time for.'),
    ('GO BACK FOR THIS','A reason to return, or a place still waiting.')]):
    y=325+i*126
    field(s,label_,prompt,48,y,499)
    rule(s,48,y+86,499)
label(s,'WRITTEN IN / PLACE     •     ON / DATE',48,742,499,BLUE)

s=page('CLOSING NOTES','What stays with us')
heading(s,'THE LAST REFLECTION','What stays\nwith us.')
photo(s,48,230,205,347,'A LAST DETAIL',SAND)
text(s,'The best surprise',280,235,267,23,'serif',BLUE)
text(s,'Write about something you could not have planned.',280,278,267,11,leading=17)
text(s,'The thing we learned',280,371,267,23,'serif',BLUE)
text(s,'A discovery about a place, each other or ourselves.',280,414,267,11,leading=17)
text(s,'If we could repeat one hour…',48,622,499,27,'italic',RUST)
text(s,'Choose a moment and tell the story of why you would return to it.',48,680,499,11,leading=17)

s=page('CLOSING NOTES','Until the next chapter')
text(s,'Until the\nnext chapter.',48,109,499,49,'serif',BLUE,leading=59)
photo(s,48,276,499,303,'THE LAST LOOK',SAND)
text(s,'There is more to see.\nFor now, there is this to remember.',48,622,499,26,'italic',leading=33)
label(s,'YOUR NAMES / PLACE / DATE',48,738,499,RUST)

def back(variant=0):
    s=page('BACK COVER','Back cover' if not variant else f'Alternative back cover {variant}',number=False)
    if variant==0:
        rect(s,0,0,W,H,BLUE)
        artwork(s,48,60,499,378,2)
        text(s,'We brought home\nmore than we packed.',48,490,499,36,'serif',PAPER,leading=44)
        text(s,'A PERSONAL TRAVEL MAGAZINE\nYOUR NAMES / ITALY / MONTH YEAR',48,674,499,10,'sans',PAPER,leading=17)
        rule(s,48,779,499,PAPER)
        label(s,'THE DAYS BETWEEN / ISSUE 01',48,798,499,PAPER)
    elif variant==1:
        artwork(s,48,62,499,440,0)
        text(s,'Somewhere new.\nSomething to keep.',48,553,499,35,'italic',BLUE,leading=43)
        label(s,'YOUR NAMES / MONTH YEAR',48,720,499,RUST)
        rule(s,48,779,499,INK)
        label(s,'THE DAYS BETWEEN / ISSUE 01',48,798,499,INK)
    else:
        photo(s,48,62,499,495,'A FINAL IMAGE',SAND)
        text(s,'And this is how\nwe will remember it.',48,598,499,32,'serif',BLUE,leading=40)
        label(s,'YOUR NAMES / PLACE / DATE',48,739,499,RUST)
        rule(s,48,779,499,INK)
        label(s,'THE DAYS BETWEEN / ISSUE 01',48,798,499,INK)

back()
cover(1);cover(2);back(1);back(2)
assert len(slides)==56, len(slides)
overflow=[m for m in metrics if m['needed']>m['h']+.1 or m['y']+m['h']>H+1 or m['x']+m['w']>W+1]
assert not overflow, overflow
prs.save(ROOT/'The-Days-Between-A4-Editable.pptx')
(ROOT/'Source'/'page-manifest.json').write_text(json.dumps(records,indent=2),encoding='utf-8')
(ROOT/'Source'/'text-metrics.json').write_text(json.dumps(metrics,indent=2))
print(f'Created {len(slides)} slides; {sum(len(r["photos"]) for r in records)} native picture placeholders; {len(metrics)} text boxes.')
