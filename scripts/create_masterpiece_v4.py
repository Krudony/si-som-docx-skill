import os; from docx import Document; from docx.shared import Pt, Twips; from docx.enum.text import WD_ALIGN_PARAGRAPH;
def create_v4_fix(output_path):
    doc = Document(); section = doc.sections[0]; section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Twips(1440)
    style = doc.styles['Normal']; font = style.font; font.name = 'TH SarabunPSK'; font.size = Pt(16)
    pf = style.paragraph_format; pf.line_spacing = 1.0; pf.space_after = pf.space_before = Pt(0); pf.widow_control = False
    def add_p(text, bold=False, align=None, indent=False):
        p = doc.add_paragraph()
        if align == 'center': p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == 'left': p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        else: p.alignment = WD_ALIGN_PARAGRAPH.THAI_JUSTIFY if hasattr(WD_ALIGN_PARAGRAPH, 'THAI_JUSTIFY') else WD_ALIGN_PARAGRAPH.JUSTIFY
        if indent: p.paragraph_format.first_line_indent = Twips(720)
        run = p.add_run(text); run.bold = bold
        return p
    add_p('โครงการ                     การพัฒนาอาคารสถานที่และสิ่งแวดล้อมอย่างมีคุณภาพ', bold=True, align='left')
    add_p('ผู้รับผิดชอบโครงการ          นายอุดร  มะโนเนือง', bold=True, align='left')
    add_p('********************************************', align='center')
    add_p('1.  หลักการและเหตุผล', bold=True, align='left')
    p1 = 'อาคารสถานที่และสิ่งแวดล้อมในสถานศึกษาเป็นปัจจัยพื้นฐานที่สำคัญอย่างยิ่งในการจัดการศึกษาที่มีคุณภาพ เนื่องจากสภาพแวดล้อมที่สะอาดร่มรื่นมั่นคงแข็งแรงและปลอดภัยจะช่วยสร้างบรรยากาศที่เอื้อต่อการเรียนรู้ของนักเรียน และส่งเสริมประสิทธิภาพในการทำงานของคณะครูและบุคลากรทางการศึกษา ในแต่ละปีการศึกษา สถานศึกษาจะได้รับการสนับสนุนงบประมาณ เพื่อให้บริหารจัดการกลุ่มงานทั้ง 4 ด้าน ได้แก่ งานวิชาการ งานงบประมาณงานบริหารงานบุคคลและงานบริหารทั่วไปซึ่งงานพัฒนาอาคารสถานที่และสิ่งแวดล้อมถือเป็นภารกิจหลักในส่วนงานบริหารทั่วไปที่มีผลกระทบโดยตรงต่อความปลอดภัยและสุขภาวะของนักเรียน'
    add_p(p1, indent=True)
    p2 = 'โรงเรียนบ้านแม่ทรายจึงเล็งเห็นความจำเป็นในการยกระดับสภาพแวดล้อมและอาคารสถานที่ ห้องเรียนรวมถึงห้องน้ำสถานศึกษาให้มีความพร้อมใช้งานมีความเป็นระเบียบเรียบร้อยและสอดคล้องกับมาตรฐานการจัดการศึกษาที่กำหนดไว้โดยเน้นการมีส่วนร่วมของคณะครู คณะกรรมการสถานศึกษา และชุมชน เพื่อให้โรงเรียนเป็นสถานที่ที่ปลอดภัยน่าอยู่น่าเรียนและเป็นแหล่งเรียนรู้ที่มีคุณภาพสำหรับเยาวชนในพื้นที่อย่างแท้จริง จึงได้จัดทำโครงการนี้ขึ้น'
    add_p(p2, indent=True)
    doc.save(output_path); print(f'V4 Created: {output_path}')
if __name__ == '__main__':
    create_v4_fix(r'C:\Users\User\Desktop\New folder (3)\สรุปโครงการ_อาคารสถานที่_ตัดคำเป๊ะ_v4.docx')
