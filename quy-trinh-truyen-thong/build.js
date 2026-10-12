const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer,
  ImageRun, AlignmentType, WidthType, BorderStyle, VerticalAlign, HeightRule,
  ShadingType, PageNumber, LevelFormat, HeadingLevel, PageBreak,
} = require('docx');

const OUT = process.argv[2];
const RED = 'A10719', DARK = '404041', GREY = '5A5A5A', ALT = 'EEEEEE';
const FONT = 'Arial';
const W = 9026; // A4 content width (DXA), margins 1440
const NB = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const TB = { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' };
const logo = fs.readFileSync(path.join(__dirname, 'logo.png'));

// ---------- helpers ----------
function runs(text, opts = {}) {
  // **bold** segments inside text
  const parts = String(text).split(/(\*\*[^*]+\*\*)/).filter(Boolean);
  return parts.map(p => p.startsWith('**')
    ? new TextRun({ text: p.slice(2, -2), bold: true, font: FONT, size: opts.size || 21, color: opts.color || '000000' })
    : new TextRun({ text: p, bold: opts.bold, italics: opts.italics, font: FONT, size: opts.size || 21, color: opts.color || '000000' }));
}
const P = (text, o = {}) => new Paragraph({
  spacing: { before: 0, after: o.after ?? 120, line: 300 },
  alignment: o.align || AlignmentType.JUSTIFIED,
  children: runs(text, o),
});
const H1 = t => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 320, after: 140 }, keepNext: true,
  children: [new TextRun({ text: t, bold: true, font: FONT, size: 26, color: RED })] });
const H2 = t => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 }, keepNext: true,
  children: [new TextRun({ text: t, bold: true, font: FONT, size: 22, color: DARK })] });
const BL = (text, level = 0) => new Paragraph({ numbering: { reference: 'bul', level }, spacing: { after: 60, line: 288 },
  alignment: AlignmentType.LEFT, children: runs(text) });
const NL = (text, ref) => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 60, line: 288 },
  alignment: AlignmentType.LEFT, children: runs(text) });

function cellParas(content, o) {
  const items = Array.isArray(content) ? content : [content];
  return items.map(t => {
    const isBullet = typeof t === 'string' && t.startsWith('• ');
    const txt = isBullet ? t.slice(2) : t;
    return new Paragraph({
      spacing: { before: 0, after: 40, line: 264 },
      alignment: o.align || AlignmentType.LEFT,
      ...(isBullet ? { numbering: { reference: 'cbul', level: 0 } } : {}),
      children: runs(txt, { size: 18, bold: o.bold, color: o.color }),
    });
  });
}
function table(widths, header, rows, o = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const mk = (cells, isHead, idx) => new TableRow({
    tableHeader: isHead, cantSplit: true,
    children: cells.map((c, i) => new TableCell({
      width: { size: widths[i], type: WidthType.DXA },
      verticalAlign: isHead ? VerticalAlign.CENTER : VerticalAlign.TOP,
      margins: { top: 70, bottom: 50, left: 100, right: 100 },
      borders: { top: TB, bottom: TB, left: TB, right: TB },
      shading: { type: ShadingType.CLEAR, color: 'auto', fill: isHead ? RED : (idx % 2 ? ALT : 'FFFFFF') },
      children: cellParas(c, isHead
        ? { bold: true, color: 'FFFFFF', align: AlignmentType.CENTER }
        : { bold: o.boldFirst && i === 0, color: '000000' }),
    })),
  });
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: [mk(header, true, 0), ...rows.map((r, i) => mk(r, false, i))],
  });
}
const gap = () => new Paragraph({ spacing: { after: 80 }, children: [] });

// ---------- header/footer (VietCredit Word standard) ----------
const header = new Header({ children: [
  new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [1900, 7126],
    borders: { top: NB, bottom: NB, left: NB, right: NB, insideHorizontal: NB, insideVertical: NB },
    rows: [new TableRow({ height: { value: 360, rule: HeightRule.EXACT }, children: [
      new TableCell({ width: { size: 1900, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
        margins: { top: 0, bottom: 0, left: 0, right: 160 }, borders: { top: NB, bottom: NB, left: NB, right: NB },
        children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [
          new ImageRun({ data: logo, type: 'png', transformation: { width: 106, height: 24 } })] })] }),
      new TableCell({ width: { size: 7126, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
        margins: { top: 0, bottom: 0, left: 160, right: 0 },
        borders: { top: NB, bottom: NB, right: NB, left: { style: BorderStyle.SINGLE, size: 6, color: RED } },
        children: [
          new Paragraph({ spacing: { before: 0, after: 16 }, alignment: AlignmentType.RIGHT, children: [
            new TextRun({ text: 'Công ty Tài chính Tổng hợp Cổ phần Tín Việt', font: FONT, size: 16, bold: true, color: GREY })] }),
          new Paragraph({ spacing: { before: 0, after: 0 }, alignment: AlignmentType.RIGHT, children: [
            new TextRun({ text: 'VietCredit General Finance Joint Stock Company', font: FONT, size: 16, bold: true, color: GREY })] }),
        ] }),
    ] })],
  }),
  new Paragraph({ spacing: { before: 36, after: 0 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: RED, space: 1 } }, children: [] }),
] });
const footer = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
  new TextRun({ text: 'Trang ', font: FONT, size: 18, color: GREY }),
  new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: GREY }),
] })] });

// ---------- content ----------
const C = [];
const SEC = t => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 320, after: 140 }, keepNext: true,
  children: [new TextRun({ text: t, bold: true, font: FONT, size: 26, color: RED })] });
const LEAD = (t, o = {}) => C.push(new Paragraph({ keepNext: true, spacing: { after: 120, line: 300 }, alignment: AlignmentType.JUSTIFIED, children: runs(t, o) }));
let numRef = 0;
const NUMS = items => { numRef++; items.forEach(t => C.push(NL(t, 'n' + numRef))); };
const BLS = items => items.forEach(t => C.push(BL(t)));
const T = (...a) => { C.push(table(...a)); C.push(gap()); };
const NOTE = t => C.push(P(t, { italics: true, color: GREY, size: 19 }));

const SHEET_URL = 'https://docs.google.com/spreadsheets/d/1xp3EIZZXDeLg45_3XUYrvG8A6sjjVtlwraIty1DJtLE';

// ===== Tiêu đề =====
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 60 }, children: [
  new TextRun({ text: 'QUY TRÌNH XỬ LÝ THÔNG TIN TIÊU CỰC TRÊN MẠNG XÃ HỘI', font: FONT, size: 30, bold: true, color: RED })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [
  new TextRun({ text: 'Phối hợp giữa VietCredit và Màn Hình Cộng (MHC)', font: FONT, size: 22, color: GREY })] }));

// ===== 1 =====
C.push(SEC('1. Thông tin chung'));
T([2400, 6626], ['Hạng mục', 'Nội dung'], [
  ['Sản phẩm áp dụng', 'Thẻ tín dụng VC x DMX; TikTok BNPL. Các dự án khác đang có Social listening: MHC theo dõi như hiện tại; sự vụ khẩn thì xử lý theo mục 4.'],
  ['Kênh theo dõi', 'Facebook, TikTok, YouTube, Threads, báo điện tử'],
  ['Link theo dõi chung', SHEET_URL],
  ['Đầu mối VietCredit', 'Marketing: [tên, số điện thoại]'],
  ['Đầu mối MHC', 'Hằng, Lam'],
  ['Nhóm chat', '[Zalo/Telegram]: báo đã cập nhật link và việc gấp. Ý kiến, duyệt, nghiệm thu đều ghi trên link để lưu vết.'],
], { boldFirst: true });

// ===== 2 =====
C.push(SEC('2. Phân nhóm tin tiêu cực'));
T([1300, 3700, 1800, 2226], ['Nhóm', 'Nội dung', 'Seeding', 'Ghi chú'], [
  ['Nhóm 1', 'Hotline không liên hệ được; cho rằng "VietCredit lừa đảo"', 'MHC làm luôn', 'Mời khách hàng liên hệ CSKH, hotline 1900 6515'],
  ['Nhóm 2', 'Lỗi tính năng, không thao tác được, mới phát sinh', 'Chờ VietCredit duyệt', ''],
  ['Nhóm 3', 'Tin tiêu cực về doanh nghiệp, tình hình kinh doanh, lãnh đạo', 'Chờ VietCredit duyệt', 'Thường là sự vụ khẩn (mục 4)'],
  ['Nhóm 4', 'Hỏi bùng nợ được không, có bị đến nhà đòi nợ không, nên vay bên nào', 'MHC làm luôn, trong quota gói', 'Không cần làm 100%, ưu tiên bài nhiều tương tác'],
  ['Khác', 'Ngoài 4 nhóm trên (mạo danh, bên thứ ba dùng tên VietCredit...)', 'Chờ VietCredit duyệt', ''],
], { boldFirst: true });
LEAD('Nhóm 1 và Nhóm 4 vẫn phải chờ VietCredit duyệt nếu bài viết:');
BLS([
  'là khách hàng cụ thể tố VietCredit làm sai, có chi tiết (ví dụ nhắc nợ gọi người thân, đe dọa);',
  'có trên 100 tương tác hoặc lên báo điện tử;',
  'do bên thứ ba dùng tên VietCredit để quảng cáo dịch vụ (rút tiền thẻ, cho vay). Trường hợp này thường report, không seeding.',
]);

// ===== 3 =====
C.push(SEC('3. Các bước xử lý hằng ngày'));
LEAD('Mỗi case là một dòng trên link theo dõi. Cột "Tình trạng" cho biết case đang ở bước nào.');
const W4 = [1200, 1500, 1500, 4826];
T(W4, ['Bước', 'Thời gian', 'Ai làm', 'Làm thế nào'], [
  ['1. Ghi nhận', 'Liên tục. Khung tổng hợp: 9h hôm trước đến 9h hôm nay', 'MHC', ['Theo dõi trên hệ thống, gửi alert cho VietCredit.', 'Sự vụ khẩn thì làm theo mục 4.']],
  ['2. Đề xuất', 'Buổi sáng, cập nhật link lúc 13:30', 'MHC', [
    '• Ghi mỗi case lên link: nền tảng, link, nhóm, tương tác, đánh giá, phương án đề xuất (theo dõi, liên hệ người đăng, seeding, report, gỡ bài).',
    '• Nhóm 1, 4: seeding luôn, ghi tình trạng "Đang triển khai".',
    '• Nhóm còn lại: ghi "Chờ VC duyệt".',
    '• Nhắn nhóm chat báo đã cập nhật.']],
  ['3. Review', '13:30 đến 14:30', 'Marketing VC', [
    'Xem từng case "Chờ VC duyệt" và kiểm tra nhanh các case Nhóm 1, 4:',
    '• Phân nhóm đúng chưa?',
    '• Phương án hợp lý chưa (ví dụ quảng cáo mạo danh thì nên report)?',
    '• Nội dung bình luận đúng chuẩn mục 5 chưa?',
    'Case cần ý kiến chuyên môn thì gửi đơn vị liên quan kèm link bài:',
    '• lỗi tính năng, phí, chính sách: PD/PM;',
    '• nhắc nợ, đòi nợ: Thu hồi nợ;',
    '• khách hàng cụ thể: CS;',
    '• mạo danh, vu khống: Pháp chế;',
    '• doanh nghiệp, lãnh đạo: trình BĐH.']],
  ['4. Cho ý kiến', 'Trước 16:00', 'PD/PM, CS, Thu hồi nợ, Pháp chế', ['Trả lời Marketing: thông tin đúng hay sai, nên trả lời thế nào, có cần liên hệ khách hàng không.', 'Quá 16:00 chưa trả lời: Marketing duyệt phương án an toàn (theo dõi, hoặc mời khách hàng liên hệ CSKH) và báo trưởng đơn vị.']],
  ['5. Duyệt', 'Trước 16:30', 'Marketing VC', ['Ghi lên link cột "Ý kiến VC": Đồng ý / Sửa thành... / Đổi phương án / Không xử lý; ghi người duyệt. Đổi tình trạng sang "Đang triển khai". Nhắn MHC.']],
  ['6. Triển khai', 'Trong ngày, chậm nhất 9h hôm sau', 'MHC (CS, Thu hồi nợ nếu phải liên hệ khách hàng)', ['Làm theo phương án đã duyệt. Cập nhật lên link: nội dung bình luận, link hoặc ảnh chụp bình luận đã đăng, kết quả report. Đổi tình trạng sang "Chờ nghiệm thu".', 'MHC không tự liên hệ khách hàng nhân danh VietCredit.']],
  ['7. Nghiệm thu', 'Case đã duyệt: trong 1 ngày làm việc. Nhóm 1, 4: thứ Sáu hằng tuần', 'Marketing VC', [
    'Mở bằng chứng trên link, kiểm tra:',
    '• đăng đúng bài, đúng nội dung, đủ số lượng;',
    '• bình luận còn hiển thị sau 48 giờ;',
    '• report, gỡ bài xong đúng hạn.',
    'Nhóm 1, 4 kiểm tra ngẫu nhiên ít nhất 20% số case trong tuần.',
    'Ghi cột "Nghiệm thu": Đạt (tình trạng "Hoàn thành") hoặc Cần bổ sung (ghi rõ cần làm gì, hạn).']],
  ['8. Bổ sung', 'Xác nhận trong 2 giờ, làm xong trong 1 ngày làm việc', 'MHC', ['Sửa, xóa, thêm bình luận, đăng lại hoặc đổi phương án theo yêu cầu. Cập nhật bằng chứng, đổi lại "Chờ nghiệm thu". Marketing nghiệm thu lại như bước 7.', 'Bình luận bị ẩn, xóa trong 48 giờ hoặc bài đã gỡ đăng lại trong thời gian bảo hành: MHC làm lại, không tính thêm phí, quota.']],
  ['9. Báo cáo', 'Thứ Hai (tuần); ngày làm việc thứ 3 (tháng)', 'MHC, Marketing VC', [
    '• MHC gửi báo cáo tuần, tháng: số case theo sản phẩm, nhóm; kết quả xử lý; quota đã dùng; các vấn đề bị phàn nàn nhiều nhất.',
    '• Marketing chuyển vấn đề có thật cho đơn vị phụ trách sửa (tổng đài quá tải: CS; lỗi lặp lại: PD/PM; nhắc nợ sai: Thu hồi nợ).',
    '• Cuối tháng đối chiếu khối lượng đã nghiệm thu "Đạt" trước khi thanh toán. Marketing báo cáo BĐH.']],
], { boldFirst: true });

// ===== 4 =====
C.push(SEC('4. Sự vụ khẩn'));
LEAD('Áp dụng khi tin liên quan lãnh đạo, tình hình kinh doanh, bảo mật, lỗi sản phẩm hàng loạt, hoặc bài có trên 100 tương tác, lên báo điện tử.');
T(W4, ['Bước', 'Thời gian', 'Ai làm', 'Làm thế nào'], [
  ['1. Báo động', 'Trong 30 phút từ khi phát hiện, cả ngoài giờ', 'MHC', 'Gọi điện và nhắn nhóm chat. Gửi ngay form sự vụ: tên sự vụ, số bài và nền tảng, nguồn đăng, nội dung và bình luận nổi bật, tương tác, đánh giá, đề xuất (kèm chi phí, thời gian nếu report, gỡ bài), những việc VietCredit nên tránh.'],
  ['2. Tiếp nhận', 'Trong 1 giờ', 'Marketing VC', 'Xác nhận với MHC, báo BĐH, gọi nhóm xử lý: PD/PM liên quan, CS, Thu hồi nợ (nếu liên quan), Pháp chế.'],
  ['3. Chốt phương án', 'Trong 4 giờ', 'Nhóm xử lý, BĐH', 'Kiểm tra sự việc, chốt phương án, nội dung phát ngôn, chi phí. CS dùng cùng nội dung trả lời khách hàng gọi tổng đài.'],
  ['4. Theo dõi, đóng', 'Đến khi ổn định; báo cáo trong 5 ngày làm việc sau đó', 'MHC, Marketing VC', 'MHC cập nhật diễn biến theo mốc đã hẹn. Marketing báo cáo BĐH nguyên nhân, cách xử lý, chi phí.'],
], { boldFirst: true });

// ===== 5 =====
C.push(SEC('5. Quy định chung'));
C.push(P('**Bình luận seeding:**', { after: 60 }));
BLS([
  'Không đăng phần nhãn trong ngoặc vuông; viết đúng tên "VietCredit".',
  'Không viết câu có thể hiểu là VietCredit làm sai khi chưa xác minh; không lên lớp, dọa người đăng; không gợi ý trốn nợ; không hứa ưu đãi, giãn nợ chưa được xác nhận.',
  'Các bình luận trong một bài khác câu chữ, đăng giãn cách, từ tài khoản khác nhau.',
]);
C.push(P('**Chi phí ngoài gói (report, gỡ bài):** Trưởng Marketing duyệt đến [hạn mức] mỗi lần; vượt mức hoặc sự vụ khẩn do BĐH duyệt.', { after: 60 }));
C.push(P('**Phát ngôn:** trong sự vụ khẩn, chỉ người được BĐH giao mới phát ngôn. Nhân viên không tự bình luận trên tài khoản cá nhân, không xóa bình luận hàng loạt trên fanpage khi chưa có ý kiến nhóm xử lý.'));

// ===== 6 =====
C.push(SEC('6. Nội dung cần chốt'));
T([1800, 7226], ['Với ai', 'Nội dung'], [
  ['BĐH', ['• Phạm vi giai đoạn 1 (mục 1).', '• Cho MHC tự seeding Nhóm 1, 4 theo kịch bản đã duyệt (mục 2).', '• Hạn mức chi phí Trưởng Marketing được duyệt; người phát ngôn khi có sự vụ khẩn.']],
  ['MHC', ['• Cập nhật link cố định lúc 13:30 (MHC đang ghi "đầu giờ chiều").', '• Các trường hợp Nhóm 1, 4 vẫn phải chờ duyệt (mục 2) và quy định bình luận (mục 5).', '• Thêm cột trên link: Tương tác, Ý kiến VC, Người duyệt, Nghiệm thu, Kiểm tra sau 48 giờ.', '• Thời hạn bổ sung (bước 8) và luồng sự vụ khẩn (mục 4).']],
], { boldFirst: true });

// ---------- document ----------
const doc = new Document({
  creator: 'Marketing VietCredit',
  title: 'Quy trình xử lý thông tin tiêu cực trên mạng xã hội',
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  numbering: { config: [
    { reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
    { reference: 'cbul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 220, hanging: 180 } } } }] },
    ...Array.from({ length: 20 }, (_, i) => ({ reference: 'n' + (i + 1), levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] })),
  ] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1200, left: 1440, right: 1440, header: 560 } } },
    headers: { default: header }, footers: { default: footer },
    children: C,
  }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log('wrote', OUT); });
