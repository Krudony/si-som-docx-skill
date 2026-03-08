import sys; import os; from docx import Document; from docx.shared import Twips, Pt; from docx.enum.text import WD_ALIGN_PARAGRAPH; from docx.enum.section import WD_SECTION; from docx.oxml.ns import qn; from docx.oxml import OxmlElement; 
class SiSomDocxEngine:
    def __init__(self, filename=None, font_size=16):
        if filename and os.path.exists(filename): self.doc = Document(filename)
        else: self.doc = Document(); self._set_page_setup(); self.set_font('TH SarabunPSK', font_size)
    def _set_page_setup(self):
        section = self.doc.sections[0]; section.page_height = Twips(16838); section.page_width = Twips(11906); section.top_margin = Twips(1418); section.bottom_margin = Twips(1418); section.left_margin = Twips(1418); section.right_margin = Twips(1418)
    def set_font(self, name='TH SarabunPSK', size=16):
        style = self.doc.styles['Normal']; font = style.font; font.name = name; font.size = Pt(size); rPr = style._element.get_or_add_rPr(); rFonts = OxmlElement('w:rFonts'); rFonts.set(qn('w:ascii'), name); rFonts.set(qn('w:hAnsi'), name); rFonts.set(qn('w:eastAsia'), name); rFonts.set(qn('w:cs'), name); rPr.append(rFonts)
    def add_paragraph(self, text, bold=False, align=None, size=None):
        p = self.doc.add_paragraph(); run = p.add_run(text)
        if bold: run.bold = True
        if size: run.font.size = Pt(size)
        if align == 'center': p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == 'right': p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        return p
    def save(self, filename): self.doc.save(filename); print(f'Document saved: {filename}')

if __name__ == '__main__':
    engine = SiSomDocxEngine()
    engine.add_paragraph('โครงการ                 การพัฒนาอาคารสถานที่และสิ่งแวดล้อมอย่างมีคุณภาพ', bold=True)
    engine.add_paragraph('แผนงาน                  งานบริหารทั่วไป')
    engine.add_paragraph('สนองกลยุทธ์สถานศึกษา           กลยุทธ์ที่ 4')
    engine.add_paragraph('มาตรฐานดำเนินการ                มฐ.ที่ 2.5')
    engine.add_paragraph('ลักษณะโครงการ           โครงการต่อเนื่อง')
    engine.add_paragraph('งบประมาณ                         10,000  บาท')
    engine.add_paragraph('ผู้รับผิดชอบโครงการ             นายอุดร  มะโนเนือง')
    engine.add_paragraph('ระยะเวลาดำเนินการ               ปีการศึกษา 2569')
    engine.add_paragraph('********************************************')
    engine.add_paragraph('1.  หลักการและเหตุผล', bold=True)
    engine.add_paragraph('             อาคารสถานที่และสิ่งแวดล้อมในสถานศึกษา เป็นปัจจัยสำคัญที่ส่งผลต่อการจัดการเรียนรู้อย่างมีคุณภาพ สถานศึกษาที่สะอาด ร่มรื่น ปลอดภัย และมีอาคารสถานที่ที่มั่นคงแข็งแรง จะช่วยสร้างบรรยากาศที่เอื้อต่อการเรียนรู้ของนักเรียนและส่งเสริมสุขภาพจิตที่ดีของคณะครูและบุคลากรทางการศึกษา')
    engine.add_paragraph('             โรงเรียนบ้านแม่ทรายจึงเห็นความสำคัญในการปรับปรุงและพัฒนาอาคารสถานที่ ห้องเรียน และสภาพแวดล้อมโดยรอบให้มีความพร้อมใช้งาน มีความปลอดภัย และสอดคล้องกับมาตรฐานการศึกษา เพื่อยกระดับคุณภาพชีวิตและคุณภาพการศึกษาให้ยั่งยืน')
    engine.add_paragraph('2.  วัตถุประสงค์', bold=True)
    engine.add_paragraph('            2.1  เพื่อปรับปรุงและพัฒนาอาคารสถานที่และสิ่งแวดล้อมให้มีความมั่นคง แข็งแรง และปลอดภัย')
    engine.add_paragraph('            2.2  เพื่อสร้างบรรยากาศและสิ่งแวดล้อมที่เอื้อต่อการจัดการเรียนรู้ของนักเรียน')
    engine.add_paragraph('            2.3  เพื่อให้สถานศึกษามีความพร้อมในการให้บริการแก่ชุมชนและหน่วยงานภายนอก')
    engine.add_paragraph('3.   เป้าหมาย', bold=True)
    engine.add_paragraph('         3.1  เชิงปริมาณ')
    engine.add_paragraph('                   3.1.1  อาคารสถานที่และห้องเรียนได้รับการปรับปรุงและซ่อมแซมให้พร้อมใช้งาน ร้อยละ 95')
    engine.add_paragraph('                   3.1.2  สภาพแวดล้อมโดยรอบมีความสะอาด ร่มรื่น และสวยงาม ร้อยละ 95')
    engine.add_paragraph('         3.2  เชิงคุณภาพ')
    engine.add_paragraph('                    3.2.1 โรงเรียนมีสภาพแวดล้อมที่เอื้อต่อการเรียนรู้ มีความปลอดภัย และเป็นที่พึงพอใจของผู้รับบริการ')
    engine.add_paragraph('4. กิจกรรมและขั้นตอนการดำเนินโครงการ', bold=True)
    engine.add_paragraph('         4.1 กิจกรรมบิ๊กคลีนนิ่งเดย์ พัฒนาสภาพแวดล้อม')
    engine.add_paragraph('         4.2 กิจกรรมปรับปรุงซ่อมแซมอาคารเรียนและห้องน้ำ')
    engine.add_paragraph('         4.3 กิจกรรมจัดทำสวนหย่อมและพื้นที่สีเขียว')
    engine.add_paragraph('5.  งบประมาณ        งบอุดหนุนรายหัว 10,000 บาท', bold=True)
    engine.add_paragraph('\n\n')
    engine.add_paragraph('                     ผู้เสนอโครงการ                                                ผู้เห็นชอบโครงการ      ', align='center')
    engine.add_paragraph('\n')
    engine.add_paragraph('                ( นายอุดร  มะโนเนือง)                                        (นายสมศักดิ์  อิปิน)       ', align='center')
    engine.add_paragraph('                 ตำแหน่ง ครูชำนาญการ                         ประธานคณะกรรมการสถานศึกษาขั้นพื้นฐาน', align='center')
    engine.add_paragraph('\n')
    engine.add_paragraph('           ผู้อนุมัติโครงการ', align='center', bold=True)
    engine.add_paragraph('\n')
    engine.add_paragraph('                                               (นายสงวน  จันทอน)', align='center')
    engine.add_paragraph('                             ผู้อำนวยการโรงเรียนบ้านบ้านแม่ทราย(คุรุราษฎร์เจริญวิทย์)', align='center')
    engine.save(r'C:\Users\User\Desktop\New folder (3)\โครงการพัฒนาอาคารสถานที่_นายอุดร.docx')
