const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  PageNumber, Header, Footer
} = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const borders = { top: border, bottom: border, left: border, right: border };

function cell(text, opts = {}) {
  return new TableCell({
    borders,
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    columnSpan: opts.span,
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: [new TextRun({ text, font: "TH Sarabun New", size: opts.size || 28, bold: opts.bold || false })]
    })]
  });
}

function mcq(num, question, choices) {
  const labels = ['ก', 'ข', 'ค', 'ง'];
  const rows = [
    new Paragraph({
      spacing: { before: 80, after: 40 },
      children: [
        new TextRun({ text: `${num}. `, font: "TH Sarabun New", size: 28, bold: true }),
        new TextRun({ text: question, font: "TH Sarabun New", size: 28 })
      ]
    })
  ];
  choices.forEach((c, i) => {
    rows.push(new Paragraph({
      spacing: { before: 20, after: 20 },
      indent: { left: 720 },
      children: [
        new TextRun({ text: `${labels[i]}. `, font: "TH Sarabun New", size: 28, bold: true }),
        new TextRun({ text: c, font: "TH Sarabun New", size: 28 })
      ]
    }));
  });
  return rows;
}

function sectionHeader(part, detail, color = "1565C0") {
  return new Paragraph({
    spacing: { before: 0, after: 80 },
    children: [
      new TextRun({ text: `${part}  `, font: "TH Sarabun New", size: 30, bold: true, color }),
      new TextRun({ text: detail, font: "TH Sarabun New", size: 28, bold: true }),
    ]
  });
}

function subHeader(text) {
  return new Paragraph({
    spacing: { before: 80, after: 40 },
    children: [new TextRun({ text, font: "TH Sarabun New", size: 26, bold: true, color: "6A1B9A" })]
  });
}

// ข้อมูล 60 ข้อ
const questions = [
  // หมวด 1
  { q: "ข้อใดคือความหมายของ 'ตรรกะ' ในการแก้ปัญหา?", c: ["การเดาคำตอบล่วงหน้า", "การใช้เหตุผลและเงื่อนไขมาพิจารณาปัญหา", "การทำงานตามความรู้สึก", "การสุ่มวิธีแก้ปัญหาจนกว่าจะสำเร็จ"] },
  { q: "ลำดับขั้นตอนการต้มบะหมี่กึ่งสำเร็จรูปข้อใดถูกต้องที่สุด?", c: ["ใส่น้ำร้อน -> ใส่บะหมี่ -> ฉีกซอง", "ฉีกซอง -> ใส่บะหมี่ -> ใส่น้ำร้อน -> รอ 3 นาที", "รอ 3 นาที -> ใส่บะหมี่ -> ใส่น้ำร้อน -> ฉีกซอง", "ใส่บะหมี่ -> รอ 3 นาที -> ใส่น้ำร้อน -> ฉีกซอง"] },
  { q: "ข้อใดคือสัญลักษณ์เริ่มต้นและสิ้นสุดของผังงาน (Flowchart)?", c: ["สี่เหลี่ยมผืนผ้า (Process)", "สี่เหลี่ยมข้าวหลามตัด (Decision)", "วงรี/แคปซูล (Terminator)", "สี่เหลี่ยมด้านขนาน (Input/Output)"] },
  { q: "การเขียนผังงานเพื่อตัดสินใจว่า 'ฝนตกหรือไม่' ต้องใช้สัญลักษณ์ใด?", c: ["สี่เหลี่ยมผืนผ้า", "วงกลม", "แคปซูล", "สี่เหลี่ยมข้าวหลามตัด"] },
  { q: "'การแบ่งปัญหาใหญ่ออกเป็นปัญหาย่อย' สอดคล้องกับหลักการใด?", c: ["Decomposition", "Pattern Recognition", "Abstraction", "Algorithm Design"] },
  { q: "ในชีวิตประจำวัน การจัดตารางสอนของนักเรียนใช้หลักการคิดแบบใดมากที่สุด?", c: ["การคิดแบบมีเงื่อนไข (เพื่อไม่ให้วิชาซ้ำกันในเวลาเดียวกัน)", "การคิดแบบวนซ้ำ", "การคิดแบบสุ่ม", "การทดลองแบบผิดพลาด"] },
  { q: "การทายตัวเลข 1-100 โดยใช้วิธีถามว่า 'มากกว่า 50 หรือไม่' เป็นการใช้อัลกอริทึมแบบใด?", c: ["การค้นหาแบบสุ่ม (Random Search)", "การค้นหาแบบเส้นตรง (Linear Search)", "การค้นหาแบบแบ่งครึ่ง (Binary Search)", "การค้นหาแบบวนรอบ (Circular Search)"] },
  { q: "คอมพิวเตอร์ไม่สามารถทำงานใดต่อไปนี้ได้ด้วยตัวเองโดยปราศจากมนุษย์สั่งการ?", c: ["ประมวลผลตัวเลข", "รับข้อมูลเข้า", "คิดค้นตั้งโจทย์ปัญหาใหม่", "แสดงผลทางหน้าจอ"] },
  { q: "ข้อใดไม่ใช่ประโยชน์ของการเขียนอัลกอริทึมก่อนเริ่มเขียนโปรแกรม?", c: ["ทำให้ตรวจพบข้อผิดพลาดได้ง่าย", "ทำให้โปรแกรมมีสีสันสวยงามขึ้น", "ทำให้เข้าใจลำดับการทำงาน", "ทำให้ผู้พัฒนาร่วมกันทำงานได้ง่าย"] },
  { q: "การกำหนดว่า 'ถ้าอายุมากกว่า 12 ปี ให้จ่ายราคาผู้ใหญ่ ถ้าไม่ใช่ให้จ่ายราคาเด็ก' ถือเป็นโครงสร้างแบบใด?", c: ["ทำซ้ำ (Loop)", "ลำดับ (Sequence)", "ทางเลือก (Condition/Decision)", "เงื่อนไขซ้อนเงื่อนไข"] },
  { q: "หากต้องการให้หุ่นยนต์รดน้ำต้นไม้ทุกวัน เวลา 07.00 น. ควรใช้คำสั่งโครงสร้างแบบใด?", c: ["หน่วงเวลา 1 สัปดาห์", "วนซ้ำและตรวจสอบเงื่อนไขเวลา", "รอรับข้อมูลจากคีย์บอร์ดทุกวัน", "เลือกทำแบบสุ่มเวลา"] },
  { q: "สมชายไปตลาดซื้อส้ม 5 กิโลกรัม กิโลกรัมละ 35 บาท ถ้าสมชายจ่ายแบงก์ 500 บาทคอมพิวเตอร์ต้องใช้ขั้นตอนใดบ้าง?", c: ["รับข้อมูลจำนวนและราคา -> คูณเงินทั้งหมด -> รับข้อมูลเงินที่จ่าย -> ลบเงินที่จ่ายด้วยยอดรวม -> แสดงเงินทอน", "รับทอนเงิน -> จ่ายเงิน -> ซื้อส้ม", "คูณเงินทั้งหมด -> จ่ายเงิน 500", "รับรู้ราคาส้มและเงินทอนเท่านั้น"] },

  // หมวด 2
  { q: "บล็อกคำสั่ง 'Move 10 steps' จะทำให้ตัวละครทำสิ่งใด?", c: ["ขยับไปข้างหน้า 10 หน่วย", "หมุนตัว 10 องศา", "กระโดดขึ้น 10 หน่วย", "ร้องเสียงสัตว์ 10 ครั้ง"] },
  { q: "คำสั่งใดที่ทำให้ตัวละครทำท่าทางเดิมซ้ำๆ ไปเรื่อยๆ ไม่มีวันสิ้นสุด?", c: ["Repeat 10", "Forever", "Wait 1 seconds", "If ... Then"] },
  { q: "ในโปรแกรม Scratch 'ฉากหลัง' เรียกว่าอะไร?", c: ["Sprite", "Backdrop หรือ Stage", "Script", "Costume"] },
  { q: "หากตัวละครเดินชนขอบจอแล้วต้องการให้เด้งกลับ ต้องใช้บล็อกใด?", c: ["Point in direction", "Go to x: ... y: ...", "If on edge, bounce", "Turn 15 degrees"] },
  { q: "การจะเปลี่ยนรูปร่างท่าทางของตัวละครเดียว ต้องเข้าเมนูใด?", c: ["Sounds", "Code", "Costumes", "Backdrop"] },
  { q: "ตัวแปร (Variable) เปรียบเสมือนอะไรในชีวิตจริง?", c: ["รถยนต์", "กล่องเก็บของที่เราสามารถเปลี่ยนของข้างในได้", "นาฬิกาบอกเวลา", "ป้ายบอกทาง"] },
  { q: "หากต้องการสร้างเกมเก็บแต้ม เมื่อตัวละครเรากินแอปเปิ้ลจะได้คะแนน ควรใช้กลุ่มบล็อกใดเพื่อเก็บคะแนน?", c: ["Motion", "Looks", "Events", "Variables"] },
  { q: "บล็อกสีส้ม 'when [space] key pressed' หมายความว่าอย่างไร?", c: ["เมื่อคลิกธงเขียว", "เมื่อคลิกเมาส์ที่ตัวละคร", "เมื่อกดปุ่มเว้นวรรคบนคีย์บอร์ด", "เมื่อลากตัวละคร"] },
  { q: "การสั่งให้ตัวละคร 2 ตัวคุยสลับกันอย่างสมจริง ต้องใช้บล็อกใดช่วย?", c: ["Wait (รอเวลา) หรือ Broadcast (ส่งข้อความ)", "Move (เดิน)", "Hide (ซ่อน)", "Size (ขนาด)"] },
  { q: "ในระนาบ X-Y ของ Scratch จอตรงกลางมีค่า x และ y เป็นเท่าใด?", c: ["x: 100, y: 100", "x: -100, y: -100", "x: 0, y: 0", "x: 240, y: 180"] },
  { q: "บล็อก 'touching [mouse-pointer]?' นำไปใช้ทำอะไรในเกม?", c: ["เช็คว่าเมาส์คลิกทิ้งไว้หรือไม่", "เช็คว่าตัวละครสัมผัสกับเคอร์เซอร์เมาส์หรือไม่", "ทำให้ตัวละครวิ่งหนีเมาส์", "ซ่อนตัวละครเมื่อเมาส์ขยับ"] },
  { q: "Bug ในการเขียนโปรแกรมหมายถึงข้อใด?", c: ["แมลงที่ติดอยู่ในคอมพิวเตอร์", "ข้อผิดพลาดหรือจุดบกพร่องในโค้ดโปรแกรม", "ฟีเจอร์พิเศษของแอปพลิเคชัน", "การอัปเดตเวอร์ชันใหม่"] },

  // หมวด 3
  { q: "อุปกรณ์ใดนำเข้าข้อมูล (Input) เสียงเข้าสู่คอมพิวเตอร์?", c: ["ลำโพง", "หูฟัง", "ไมโครโฟน", "เครื่องพิมพ์"] },
  { q: "หน่วยความจำใดที่ข้อมูลจะหายไปเมื่อปิดเครื่องคอมพิวเตอร์?", c: ["ROM", "Hard Disk", "Flash Drive", "RAM"] },
  { q: "ข้อใดคือหน้าที่หลักของ CPU?", c: ["เก็บข้อมูลไฟล์งาน", "เป็นแหล่งจ่ายไฟให้คอมพิวเตอร์", "ประมวลผลคำสั่งและข้อมูลทั้งหมด", "แสดงภาพบนหน้าจอ"] },
  { q: "อุปกรณ์ใดทำหน้าที่เป็น 'อุปกรณ์ส่งออก' (Output Device) ทั้งหมด?", c: ["จอภาพ, เครื่องพิมพ์, ลำโพง", "คีย์บอร์ด, เมาส์, สแกนเนอร์", "จอภาพ, ไมโครโฟน, แฟลชไดรฟ์", "กล้องวงจรปิด, เมาส์, ลำโพง"] },
  { q: "หน่วยบรรจุข้อมูลใดมีความจุมากที่สุด?", c: ["Megabyte (MB)", "Gigabyte (GB)", "Terabyte (TB)", "Kilobyte (KB)"] },
  { q: "ข้อใดต่อไปนี้เป็น 'ซอฟต์แวร์ระบบ' (System Software)?", c: ["Microsoft Word", "Google Chrome", "เกม Minecraft", "Windows 11"] },
  { q: "การกดแฟ้นคีย์บอร์ด Ctrl + C คือการสั่งงานใด?", c: ["คัดลอก (Copy)", "วาง (Paste)", "ตัด (Cut)", "พิมพ์ (Print)"] },
  { q: "หากต้องการย้ายไฟล์จากคอมพิวเตอร์เครื่องหนึ่งไปอีกเครื่องหนึ่งโดยไม่ใช้อินเทอร์เน็ต ควรใช้อุปกรณ์ใด?", c: ["RAM", "CPU", "USB Flash Drive", "จอภาพ"] },
  { q: "สแกนเนอร์ (Scanner) เทียบได้กับอวัยวะใดของมนุษย์ในการรับรู้ข้อมูล?", c: ["หู", "ตา", "ปาก", "สมอง"] },
  { q: "SSD และ HDD ทำหน้าที่เหมือนกัน แต่อะไรคือสิ่งที่ SSD ดีกว่า?", c: ["สวยงามกว่า", "ราคาถูกกว่ามากๆ ในความจุเท่ากัน", "มีความเร็วในการอ่านและเขียนข้อมูลสูงกว่า", "สามารถใช้แทนจอภาพได้"] },
  { q: "ซอฟต์แวร์นำเสนอ (Presentation Software) คือโปรแกรมใด?", c: ["Microsoft Excel", "Microsoft PowerPoint", "Microsoft Word", "Microsoft Access"] },
  { q: "ข้อใดไม่ใช่อุปกรณ์ที่อยู่ภายในเคสคอมพิวเตอร์ (Computer Case)?", c: ["Mainboard", "Power Supply", "CPU", "Projector"] },

  // หมวด 4
  { q: "บริการใดบนอินเทอร์เน็ตที่ใช้ค้นหาข้อมูลต่างๆ ได้ทั่วโลก?", c: ["E-mail", "Search Engine", "E-commerce", "Social Media"] },
  { q: "Keyword ในการค้นหา (Search) หมายถึงอะไร?", c: ["รหัสผ่านสำหรับเข้าเว็บไซต์", "คำหลักหรือคำสำคัญที่ใช้ค้นหาข้อมูล", "การเข้ารหัสข้อมูล", "รูปภาพที่ใช้ค้นหาประวัติ"] },
  { q: "โดเมนเนม (Domain name) นามสกุล '.ac.th' หมายถึงหน่วยงานใดในไทย?", c: ["บริษัทเอกชน", "หน่วยงานทหาร", "สถานศึกษา/สถาบันการศึกษา", "องค์กรรัฐบาลทั่วไป"] },
  { q: "ข้อใดคือตัวเว็บบราวเซอร์ (Web Browser) ทั้งหมด?", c: ["Google, Yahoo, Bing", "Chrome, Safari, Edge", "Windows, macOS, Android", "Facebook, Instagram, TikTok"] },
  { q: "หากนักเรียนต้องการส่งไฟล์เอกสารด่วนไปให้เพื่อนที่อยู่ต่างจังหวัด ควรใช้บริการใดเหมาะสมที่สุด?", c: ["ไปรษณีย์อิเล็กทรอนิกส์ (E-mail)", "วิดีโอคอล (Video Call)", "โอนเงินผ่านแอปธนาคาร", "พิมพ์แล้วส่งจดหมายกระดาษ"] },
  { q: "'IoT' ย่อมาจาก Internet of Things หมายถึงอะไร?", c: ["การเล่นเกมผ่านอินเทอร์เน็ต", "อุปกรณ์หรือสิ่งของต่างๆ ที่สามารถเชื่อต่อและสื่อสารกันผ่านอินเทอร์เน็ตได้", "สกุลเงินดิจิทัล", "การสืบค้นเอกสาร"] },
  { q: "เมื่อนำมือถือสองเครื่องมาส่งรูปผ่าน Bluetooth เป็นเครือข่ายระดับใด?", c: ["เครือข่ายส่วนบุคคล (PAN)", "เครือข่ายท้องถิ่น (LAN)", "เครือข่ายระดับเมือง (MAN)", "เครือข่ายระดับประเทศ (WAN)"] },
  { q: "URL คืออะไร?", c: ["โปรแกรมดูวิดีโอ", "ที่อยู่ของเว็บไซต์หรือหน้าเว็บ", "อุปกรณ์กระจายสัญญาณไวไฟ", "ไวรัสคอมพิวเตอร์ประเภทหนึ่ง"] },
  { q: "เว็บไซต์ที่มีแม่กุญแจล็อค (https://) แสดงว่าเว็บไซต์นั้นมีคุณสมบัติใด?", c: ["ไม่คิดค่าบริการอินเทอร์เน็ต", "เป็นเว็บไซต์ของรัฐบาล", "มีการเข้ารหัสเพื่อความปลอดภัยของข้อมูล", "ไม่สามารถดาวน์โหลดรูปได้"] },
  { q: "Cloud Computing ในชีวิตประจำวันหมายถึงอะไร?", c: ["การคาดเดาสภาพอากาศ", "การใช้งานพื้นที่เก็บข้อมูลบนอินเทอร์เน็ต เช่น Google Drive", "การทำงานของพัดลมในซ็อกเก็ตชิป", "การอัปโหลดรูปลงโทรศัพท์มือถือเครื่องเดียว"] },
  { q: "หากมีอีเมลแปลกหน้าส่งลิงก์มาบอกว่า 'คุณได้รับรางวัล 1 ล้านบาท คลิกที่นี่' นักเรียนควรทำอย่างไร?", c: ["รีบส่งต่อให้เพื่อน 10 คน", "คลิกดูเพื่อตรวจสอบว่าจริงหรือไม่", "กรอกข้อมูลส่วนตัวเพื่อรับรางวัล", "ห้ามคลิก และลบอีเมลฉบับนั้นทิ้งทันที"] },
  { q: "แฮกเกอร์ (Hacker) ด้านมืด มีพฤติกรรมอย่างไร?", c: ["ลักลอบเจาะระบบเพื่อขโมยข้อมูลหรือสร้างความเสียหาย", "ซื้อขายอุปกรณ์คอมพิวเตอร์มือสอง", "สร้างเกมให้คนเล่นฟรี", "สอนคอมพิวเตอร์บนอินเทอร์เน็ต"] },

  // หมวด 5
  { q: "ข้อมูลใดต่อไปนี้ เป็น 'ข้อมูลส่วนบุคคล' ที่ไม่ควรโพสต์ลงสื่อสาธารณะ?", c: ["ชื่อภาพยนตร์ที่ชอบดู", "รหัสบัตรประชาชน 13 หลัก", "สีที่ชอบ", "อาหารมื้อเช้าที่เพิ่งทาน"] },
  { q: "พฤติกรรมใดเสี่ยงต่อการถูกขโมยบัญชีผู้ใช้มากที่สุด?", c: ["เปลี่ยนรหัสผ่านทุก 6 เดือน", "ใช้คอมพิวเตอร์สาธารณะแล้วลืม Log out", "ตั้งรหัสผ่านที่เดายาก", "เชื่อมต่อ Wi-Fi ที่บ้านที่มีรหัสผ่าน"] },
  { q: "ข้อใดถือเป็นมารยาทที่ดีในการใช้อินเทอร์เน็ต (Netiquette)?", c: ["พิมพ์ข้อความว่าร้ายผู้อื่นโดยใช้นามแฝง", "คัดลอกงานของคนอื่นมาส่งครูโดยไม่ให้เครดิต", "เคารพสิทธิและไม่ละเมิดความเป็นส่วนตัวของผู้อื่น", "ส่งข้อความขยะ (Spam) ไปก่อกวนเพื่อน"] },
  { q: "Fake News หรือข่าวปลอม มักมีลักษณะอย่างไร?", c: ["แหล่งที่มาน่าเชื่อถือ อ้างอิงจากสำนักข่าวหลัก", "วันที่ปัจจุบัน ข้อมูลตรงตามความเป็นจริง", "มีการพาดหัวข่าวเกินจริงกระตุ้นความตื่นตระหนก ไม่ระบุที่มา", "มีแหล่งอ้างอิงทางวิชาการชัดเจน"] },
  { q: "การ 'รังแกกันบนโลกไซเบอร์' (Cyberbullying) ส่งผลกระทบอย่างไรต่อผู้ถูกกระทำ?", c: ["มีสุขภาพร่างกายแข็งแรงขึ้น", "มีภูมิคุ้มกันในการใช้ชีวิตมากขึ้น", "เกิดความเครียด ซึมเศร้า และอาจทำร้ายตัวเอง", "เป็นที่นิยมในสังคม"] },
  { q: "การเช็คความน่าเชื่อถือของข่าวแชร์ใน LINE ทำได้อย่างไร?", c: ["เช็คจากศูนย์ต่อต้านข่าวปลอม (Anti-Fake News Center)", "ถามเพื่อนที่เพิ่งแชร์มา", "พิจารณาจากยอดกดไลก์ ถ้าไลก์เยอะคือเรื่องจริง", "เชื่อเลยเพราะมาจากคนรู้จัก"] },
  { q: "หากมีคนเอารูปนักเรียนไปตัดต่อล้อเลียนใน Facebook นักเรียนควรทำข้อใดเป็นอันดับแรก?", c: ["ถ่ายรูปแบบเดียวกันคืน", "แคปหน้าจอเป็นหลักฐาน แล้วแจ้งผู้ปกครองหรือครูเพื่อรายงานบัญชี (Report)", "เข้าไปด่ากลับในคอมเมนต์", "ย้ายโรงเรียน"] },
  { q: "เครื่องหมาย CC (Creative Commons) มีไว้สำหรับอะไร?", c: ["บอกราคาของซอฟต์แวร์", "แสดงเงื่อนไขการนำผลงานไปใช้งานหรือเผยแพร่ต่อ", "บอกขนาดไฟล์ของรูปภาพ", "ปุ่มเปิด/ปิดโปรแกรม"] },
  { q: "หากติดตั้งโปรแกรมเถื่อน (Crack) ในคอมพิวเตอร์ จะมีความเสี่ยงเรื่องใดมากที่สุด?", c: ["คอมพิวเตอร์จะมีราคาแพงขึ้น", "โดนไวรัส มัลแวร์ขโมยข้อมูลหรือทำให้เครื่องพัง", "ความเร็วอินเทอร์เน็ตเพิ่มขึ้น", "จอภาพสว่างขึ้นเอง"] },
  { q: "รหัสผ่านที่ปลอดภัยควรมีลักษณะใด?", c: ["123456", "ชื่อเล่น+วันเกิด", "ตัวอักษรพิมพ์ใหญ่ พิมพ์เล็ก ตัวเลข และอักขระพิเศษผสมกันยาว 8 ตัวขึ้นไป", "เบอร์โทรศัพท์"] },
  { q: "ข้อใดคือตัวอย่างของ Phishing?", c: ["การตกปลาในเกม", "การสร้างหน้าเว็บปลอมเลียนแบบธนาคารเพื่อหลอกขอรหัสผ่าน", "การซื้อขายสินค้าออนไลน์ผ่านเว็บไซต์ต่างประเทศ", "โปรแกรมแชทสำหรับองค์กร"] },
  { q: "เพราะเหตุใดเราจึงไม่ควรเล่นโทรศัพท์มือถือก่อนนอนในที่มืด?", c: ["ทำให้โทรศัพท์พังเร็ว", "แสงสีฟ้าจากจอจะทำลายสายตาและทำให้นอนไม่หลับ", "ทำให้ค่าไฟพุ่งสูงขึ้นมาก", "แบตเตอรี่จะเสื่อมอย่างรวดเร็ว"] }
];

const keys = ["ข", "ข", "ค", "ง", "ก", "ก", "ค", "ค", "ข", "ค", "ข", "ก",
              "ก", "ข", "ข", "ค", "ค", "ข", "ง", "ค", "ก", "ค", "ข", "ข",
              "ค", "ง", "ค", "ก", "ค", "ง", "ก", "ค", "ข", "ค", "ข", "ง",
              "ข", "ข", "ค", "ข", "ก", "ข", "ก", "ข", "ค", "ข", "ง", "ก",
              "ข", "ข", "ค", "ค", "ค", "ก", "ข", "ข", "ข", "ค", "ข", "ข"];


const docElements = [
  // Header banner
  new Table({
    width: { size: 9638, type: WidthType.DXA },
    columnWidths: [9638],
    rows: [new TableRow({ children: [
      new TableCell({
        borders,
        shading: { fill: "1565C0", type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 200, right: 200 },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "แบบทดสอบปลายภาค (ฉบับ 60 ข้อ)", font: "TH Sarabun New", size: 42, bold: true, color: "FFFFFF" })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "วิชาวิทยาการคำนวณ (Computing Science)", font: "TH Sarabun New", size: 32, bold: true, color: "FFD700" })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "ชั้นประถมศึกษาปีที่ 6   ภาคเรียนที่ 2   ปีการศึกษา ............", font: "TH Sarabun New", size: 28, color: "FFFFFF" })] }),
        ]
      })
    ]})]
  }),
  new Paragraph({ spacing: { before: 80, after: 0 }, children: [] }),

  // Info row
  new Table({
    width: { size: 9638, type: WidthType.DXA },
    columnWidths: [4819, 4819],
    rows: [
      new TableRow({ children: [
        cell("เวลา  90  นาที", { width: 4819, bold: true }),
        cell("คะแนนเต็ม  60  คะแนน", { width: 4819, bold: true, align: AlignmentType.RIGHT }),
      ]}),
      new TableRow({ children: [
        new TableCell({
          borders,
          columnSpan: 2,
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [
            new TextRun({ text: "ชื่อ–สกุล  ", font: "TH Sarabun New", size: 28, bold: true }),
            new TextRun({ text: ".......................................................................  ", font: "TH Sarabun New", size: 28 }),
            new TextRun({ text: "ชั้น  ", font: "TH Sarabun New", size: 28, bold: true }),
            new TextRun({ text: "..............  ", font: "TH Sarabun New", size: 28 }),
            new TextRun({ text: "เลขที่  ", font: "TH Sarabun New", size: 28, bold: true }),
            new TextRun({ text: "............", font: "TH Sarabun New", size: 28 }),
          ]})]
        })
      ]}),
    ]
  }),
  new Paragraph({ spacing: { before: 120, after: 0 }, children: [] }),

  sectionHeader("ส่วนที่ 1", "แบบปรนัย (60 ข้อ  ข้อละ 1 คะแนน  รวม 60 คะแนน)"),
  new Paragraph({
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: "คำสั่ง:  เลือกคำตอบที่ถูกต้องที่สุดเพียงข้อเดียว แล้วทำเครื่องหมาย  X  ลงในกระดาษคำตอบ", font: "TH Sarabun New", size: 26, bold: true })]
  })
];

const sectionTitles = [
  "หมวดที่ 1: การแก้ปัญหาเชิงตรรกะและอัลกอริทึม",
  "หมวดที่ 2: การเขียนโปรแกรม (Scratch พื้นฐาน)",
  "หมวดที่ 3: ระบบคอมพิวเตอร์และฮาร์ดแวร์",
  "หมวดที่ 4: เครือข่ายคอมพิวเตอร์และอินเทอร์เน็ต",
  "หมวดที่ 5: การพิจารณาข้อมูลและความปลอดภัยทางไซเบอร์"
];

let qCount = 0;
for (let section = 0; section < 5; section++) {
  docElements.push(subHeader(sectionTitles[section]));
  for (let i = 0; i < 12; i++) {
    const qObj = questions[qCount];
    docElements.push(...mcq(qCount + 1, qObj.q, qObj.c));
    qCount++;
  }
}

// Answer Key Section
docElements.push(new Paragraph({ spacing: { before: 200, after: 80 }, children: [] }));
docElements.push(sectionHeader("ส่วนที่ 2", "เฉลยข้อสอบ (Answer Key)", "D32F2F"));

const answerRows = [];
for (let r = 0; r < 12; r++) {
  const cells = [];
  for (let c = 0; c < 5; c++) {
    const num = c * 12 + r + 1;
    cells.push(new TableCell({
      borders: {
        top: { style: BorderStyle.SINGLE, size: 2, color: "000000" },
        bottom: { style: BorderStyle.SINGLE, size: 2, color: "000000" },
        left: { style: BorderStyle.SINGLE, size: 2, color: "000000" },
        right: { style: BorderStyle.SINGLE, size: 2, color: "000000" }
      },
      margins: { top: 40, bottom: 40, left: 80, right: 80 },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: `${num}. `, font: "TH Sarabun New", size: 28, bold: true }),
          new TextRun({ text: `${keys[num-1]}`, font: "TH Sarabun New", size: 28, color: "E65100", bold: true })
        ]
      })]
    }));
  }
  answerRows.push(new TableRow({ children: cells }));
}

const answerTable = new Table({
  width: { size: 9638, type: WidthType.DXA },
  rows: answerRows
});
docElements.push(answerTable);


const doc = new Document({
  styles: {
    default: { document: { run: { font: "TH Sarabun New", size: 28 } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: "วิทยาการคำนวณ ป.6 เทอม 2 (ฉบับ 60 ข้อ)", font: "TH Sarabun New", size: 22, color: "666666" })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "หน้า ", font: "TH Sarabun New", size: 22, color: "666666" }),
            new TextRun({ children: [PageNumber.CURRENT], font: "TH Sarabun New", size: 22, color: "666666" }),
            new TextRun({ text: " / ", font: "TH Sarabun New", size: 22, color: "666666" }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "TH Sarabun New", size: 22, color: "666666" }),
          ]
        })]
      })
    },
    children: docElements
  }]
});

Packer.toBuffer(doc).then(buffer => {
  const outPath = "C:/Users/User/Desktop/New folder (2)/ข้อสอบใหม่วิทยาการคำนวณ_ป6_60ข้อ_พร้อมเฉลย_ตาราง.docx";
  fs.writeFileSync(outPath, buffer);
  console.log("Done! Saved to: " + outPath);
});
