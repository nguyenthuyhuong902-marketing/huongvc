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
const PART = (code, title, note) => {
  C.push(new Paragraph({
    heading: HeadingLevel.HEADING_1, pageBreakBefore: true, keepNext: true,
    spacing: { before: 0, after: note ? 0 : 160 },
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: RED },
    children: [new TextRun({ text: ` ${code}. ${title}`, bold: true, font: FONT, size: 26, color: 'FFFFFF' })],
  }));
  if (note) C.push(new Paragraph({ spacing: { before: 80, after: 160 }, keepNext: true,
    children: runs(note, { italics: true, color: GREY, size: 19 }) }));
};
const SEC = t => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 280, after: 120 }, keepNext: true,
  children: [new TextRun({ text: t, bold: true, font: FONT, size: 24, color: RED })] });
const SUB = t => new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 200, after: 100 }, keepNext: true,
  children: [new TextRun({ text: t, bold: true, font: FONT, size: 22, color: DARK })] });
const LEAD = t => new Paragraph({ keepNext: true, spacing: { after: 120, line: 300 }, alignment: AlignmentType.JUSTIFIED, children: runs(t) });
let numRef = 0;
const NUMS = items => { numRef++; items.forEach(t => C.push(NL(t, 'n' + numRef))); };
const BLS = items => items.forEach(t => C.push(BL(t)));
const T = (...a) => { C.push(table(...a)); C.push(gap()); };

const SHEET_URL = 'https://docs.google.com/spreadsheets/d/1xp3EIZZXDeLg45_3XUYrvG8A6sjjVtlwraIty1DJtLE';

// ===== Bìa =====
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 240, after: 80 }, children: [
  new TextRun({ text: 'QUY TRÌNH', font: FONT, size: 32, bold: true, color: RED })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [
  new TextRun({ text: 'QUẢN TRỊ RỦI RO TRUYỀN THÔNG VÀ XỬ LÝ THÔNG TIN TIÊU CỰC TRÊN MẠNG XÃ HỘI', font: FONT, size: 26, bold: true, color: RED })] }));
T([2600, 6426], ['Thông tin', 'Nội dung'], [
  ['Tình trạng', 'Bản hoàn thiện trình Ban Điều hành, đã cập nhật theo quy trình phối hợp MHC gửi ngày 12/10/2026'],
  ['Đơn vị chủ trì', 'Marketing VietCredit'],
  ['Đơn vị phối hợp', 'Quản lý sản phẩm (PD/PM), Chăm sóc khách hàng (CS), Thu hồi nợ, Pháp chế, đối tác Màn Hình Cộng (MHC)'],
  ['Sản phẩm áp dụng giai đoạn 1', 'Thẻ tín dụng VietCredit x Điện Máy Xanh (VC x DMX); TikTok BNPL'],
  ['Link theo dõi chung', SHEET_URL],
], { boldFirst: true });
C.push(SUB('Tài liệu gồm 4 phần'));
T([1000, 4626, 3400], ['Phần', 'Nội dung', 'Người đọc'], [
  ['A', 'Tổng quan: phạm vi, vòng đời một case, lịch một ngày làm việc', 'Ban Điều hành, tất cả các bên'],
  ['B', 'Phối hợp VietCredit và MHC: phân nhóm, phương án, cách VietCredit duyệt, cập nhật link, nghiệm thu, yêu cầu bổ sung, sự vụ khẩn', 'MHC và Marketing. Gửi MHC làm thỏa thuận làm việc'],
  ['C', 'Quy trình nội bộ VietCredit: review đề xuất, lấy ý kiến phòng ban, nghiệm thu, xử lý sự vụ nghiêm trọng, chi phí, phòng ngừa', 'Các phòng ban VietCredit. Không gửi MHC'],
  ['D', 'Triển khai, nội dung cần BĐH quyết định, điểm cần thống nhất thêm với MHC', 'Ban Điều hành, Marketing'],
], { boldFirst: true });

// ===== PHẦN A =====
PART('PHẦN A', 'TỔNG QUAN');

C.push(SEC('A1. Mục đích'));
BLS([
  'Phát hiện sớm thông tin tiêu cực về VietCredit và sản phẩm trên mạng xã hội, báo điện tử.',
  'Quy định rõ việc MHC được làm ngay, việc phải chờ VietCredit duyệt, việc phải báo Ban Điều hành.',
  'Kiểm soát chất lượng công việc của MHC qua bước nghiệm thu, có cơ chế yêu cầu làm lại hoặc bổ sung.',
  'Chuyển các phàn nàn có thật về cho phòng ban liên quan để xử lý tận gốc.',
]);

C.push(SEC('A2. Phạm vi áp dụng'));
T([2200, 6826], ['Hạng mục', 'Áp dụng'], [
  ['Sản phẩm giai đoạn 1', 'Thẻ tín dụng VC x DMX; TikTok BNPL: áp dụng toàn bộ quy trình.'],
  ['Các dự án khác', 'Các dự án đang có báo cáo Social listening: MHC theo dõi và đề xuất như hiện tại. Sự vụ nghiêm trọng (Cấp độ 1) của bất kỳ dự án nào đều xử lý theo luồng khẩn.'],
  ['Mở rộng sau này', ['Marketing đề xuất BĐH bổ sung sản phẩm đáp ứng ít nhất một tiêu chí:', '• VietCredit chịu toàn bộ rủi ro;', '• dự án có margin tốt;', '• dự án đang thử nghiệm hoặc có vấn đề cần đo lường, giám sát.']],
  ['Kênh theo dõi', 'Facebook, TikTok, YouTube, Threads, báo điện tử và trang tin.'],
], { boldFirst: true });

C.push(SEC('A3. Thuật ngữ'));
T([2300, 6726], ['Thuật ngữ', 'Giải thích'], [
  ['MHC', 'Màn Hình Cộng, đối tác cung cấp dịch vụ Social listening và quản trị tiêu cực (hệ thống Monitaz)'],
  ['Hệ thống Crisis', 'Hệ thống của MHC ghi nhận các tin đã cảnh báo (alert) cho VietCredit'],
  ['Link theo dõi', 'Google Sheet dùng chung giữa hai bên: sheet "Tư vấn XLKH" (danh sách case, phương án) và sheet "Seeding" (bình luận, nghiệm thu)'],
  ['Case', 'Một bài viết hoặc một sự việc tiêu cực cần xem xét'],
  ['Seeding', 'Viết bình luận giải thích, định hướng, pha loãng thảo luận tiêu cực dưới bài viết công khai'],
  ['Report', 'Báo cáo bài viết với nền tảng để hạn chế tương tác hoặc gỡ bài'],
  ['Nghiệm thu', 'VietCredit kiểm tra việc MHC đã làm có đúng phương án, đủ số lượng, còn hiển thị, có bằng chứng hay không'],
  ['Yêu cầu bổ sung', 'Yêu cầu MHC làm thêm, làm lại hoặc đổi phương án sau khi nghiệm thu hoặc khi tình hình thay đổi'],
], { boldFirst: true });

C.push(SEC('A4. Vòng đời một case'));
LEAD('Mỗi case đi qua các bước sau. Cột "Tình trạng xử lý" trên link theo dõi ghi bước hiện tại.');
T([700, 2300, 1700, 4326], ['Bước', 'Việc', 'Ai làm', 'Tình trạng ghi trên link'], [
  ['1', 'Phát hiện, cảnh báo', 'MHC', 'Tin đã alert, ghi nhận trên hệ thống Crisis'],
  ['2', 'Tổng hợp, phân nhóm, đề xuất phương án', 'MHC', 'Nhóm 1, 4: "Đang triển khai". Các nhóm khác: "Chờ VC duyệt"'],
  ['3', 'Review, duyệt phương án', 'VietCredit', '"Đã duyệt" (kèm ý kiến) hoặc "Không xử lý"'],
  ['4', 'Triển khai', 'MHC (CS, Thu hồi nợ nếu phải liên hệ khách hàng)', '"Đang triển khai"'],
  ['5', 'Cập nhật kết quả, bằng chứng', 'MHC', '"Chờ nghiệm thu"'],
  ['6', 'Nghiệm thu', 'VietCredit', '"Hoàn thành" hoặc "Cần bổ sung"'],
  ['7', 'Bổ sung (nếu có), theo dõi tiếp', 'MHC', 'Quay lại bước 5; hoặc "Theo dõi" nếu cần xem thêm diễn biến'],
], { boldFirst: true });

C.push(SEC('A5. Lịch một ngày làm việc'));
LEAD('Theo quy trình MHC chốt ngày 12/10/2026, có bổ sung các mốc phía VietCredit.');
T([2200, 1700, 5126], ['Thời điểm', 'Ai', 'Việc'], [
  ['Liên tục trong ngày', 'MHC', 'Gửi alert khi phát hiện tin tiêu cực. Sự vụ nghiêm trọng xử lý theo luồng khẩn (mục B8)'],
  ['Buổi sáng', 'MHC', 'Tra soát tin trên hệ thống Crisis, khung ghi nhận từ 9h00 hôm trước đến 9h00 hôm nay. Nhóm 1 và Nhóm 4 seeding luôn'],
  ['13:30', 'MHC', 'Cập nhật phương án cho từng case lên link theo dõi, báo đầu mối VietCredit qua nhóm chat'],
  ['13:30 đến 16:00', 'VietCredit nội bộ', 'Marketing review, hỏi phòng ban liên quan (Phần C)'],
  ['Trước 16:30', 'Marketing', 'Ghi ý kiến, duyệt phương án trên link, báo MHC'],
  ['Trong ngày, chậm nhất 9h00 hôm sau', 'MHC', 'Triển khai phương án đã duyệt, cập nhật kết quả và bằng chứng lên link'],
  ['Hằng ngày (case cần duyệt), thứ Sáu (Nhóm 1, 4)', 'Marketing', 'Nghiệm thu, ghi nhận xét, yêu cầu bổ sung nếu cần'],
], { boldFirst: true });
C.push(P('MHC ghi "đầu giờ chiều". Quy trình đề xuất chốt mốc 13:30 để hai bên chủ động sắp xếp; cần MHC xác nhận.', { italics: true, color: GREY, size: 19 }));

// ===== PHẦN B =====
PART('PHẦN B', 'PHỐI HỢP GIỮA VIETCREDIT VÀ MHC',
  'Phần này gửi MHC. Các mốc thời gian và bốn nhóm seeding lấy theo quy trình MHC gửi ngày 12/10/2026. Các điểm VietCredit đề nghị bổ sung được ghi rõ.');

C.push(SEC('B1. Đầu mối và kênh trao đổi'));
T([2200, 3400, 3426], ['Bên', 'Đầu mối', 'Kênh'], [
  ['VietCredit', 'Marketing: [tên, số điện thoại]. Mọi trao đổi đi qua đầu mối này', 'Link theo dõi cho việc hằng ngày; nhóm chat cho thông báo và việc gấp'],
  ['MHC', 'Hằng, Lam', 'Như trên'],
  ['Nhóm chat khẩn', 'Đầu mối hai bên, Trưởng Marketing VietCredit', '[Zalo hoặc Telegram, hai bên chốt]'],
], { boldFirst: true });
C.push(P('Ý kiến, duyệt, nhận xét, yêu cầu bổ sung đều ghi trên link theo dõi để lưu vết. Nhóm chat chỉ dùng để báo "đã cập nhật" hoặc trao đổi việc gấp.'));

C.push(SEC('B2. Phân nhóm và quyền seeding'));
T([1100, 3300, 2000, 2626], ['Nhóm', 'Nội dung', 'Quyền seeding', 'Ghi chú'], [
  ['Nhóm 1', 'Hotline không liên hệ được; cho rằng "VietCredit lừa đảo"', 'MHC chủ động', 'Hướng bình luận: mời khách hàng liên hệ CSKH VietCredit, hotline 1900 6515'],
  ['Nhóm 2', 'Lỗi tính năng, không thao tác được, mới phát sinh', 'Thống nhất với VietCredit trước', 'VietCredit xác nhận lỗi và hướng trả lời'],
  ['Nhóm 3', 'Tin tiêu cực về doanh nghiệp, tình hình kinh doanh, lãnh đạo', 'Thống nhất với VietCredit trước', 'Thường là Cấp độ 1, xem mục B8'],
  ['Nhóm 4', 'Hỏi bùng nợ được không, có bị về nhà đòi nợ không, nên vay bên nào', 'MHC chủ động, trong quota gói', 'Không cần xử lý 100%; ưu tiên bài nhiều tương tác'],
  ['Ngoài 4 nhóm', 'Ví dụ: bên thứ ba quảng cáo rút tiền thẻ tín dụng dùng tên VietCredit; mạo danh; tin tuyển dụng, nhân sự', 'Thống nhất với VietCredit trước', 'VietCredit đề nghị bổ sung'],
], { boldFirst: true });
C.push(SUB('Các trường hợp MHC không tự seeding, kể cả Nhóm 1 và Nhóm 4'));
C.push(P('VietCredit đề nghị MHC chuyển sang "Chờ VC duyệt" khi bài viết thuộc một trong các trường hợp:'));
BLS([
  'Khách hàng cụ thể tố VietCredit làm sai, có chi tiết hoặc bằng chứng (ví dụ: nhắc nợ làm phiền người thân, đe dọa, thu phí sai). VietCredit cần kiểm tra thực tế trước khi bình luận.',
  'Bài có trên 100 tương tác hoặc xuất hiện trên báo điện tử.',
  'Bài do bên thứ ba dùng tên, logo VietCredit để quảng cáo dịch vụ (rút tiền thẻ, cho vay, hỗ trợ nợ xấu). Phương án phù hợp thường là report, không phải seeding.',
]);

C.push(SEC('B3. Phương án xử lý và quyền quyết định'));
T([1900, 3200, 3926], ['Phương án', 'Ai quyết định', 'Ghi chú'], [
  ['Theo dõi thêm', 'MHC đề xuất, VietCredit duyệt', 'MHC báo lại nếu tương tác tăng nhanh'],
  ['Liên hệ người đăng', 'VietCredit duyệt', 'MHC cung cấp thông tin người đăng nếu có. CS hoặc Thu hồi nợ của VietCredit liên hệ khách hàng; MHC không tự liên hệ nhân danh VietCredit'],
  ['Seeding', 'Nhóm 1, 4: MHC. Còn lại: VietCredit duyệt', 'Theo chuẩn bình luận ở mục B4'],
  ['Report', 'VietCredit duyệt', 'MHC ghi thời hạn dự kiến và báo kết quả'],
  ['Gỡ bài', 'VietCredit duyệt (có chi phí ngoài gói)', 'Đề xuất phải ghi chi phí, thời gian, bảo hành, tỷ lệ thành công'],
], { boldFirst: true });

C.push(SEC('B4. Chuẩn nội dung bình luận seeding'));
BLS([
  'Đăng đúng câu đã duyệt. Phần nhãn trong ngoặc vuông (ví dụ "[Hướng dẫn liên hệ kênh hỗ trợ chính thức]") chỉ dùng để phân loại trên link, không đăng lên mạng xã hội.',
  'Viết tên thương hiệu là "VietCredit", không viết "VIETCREDIT", "Viet Credit".',
  'Không viết câu có thể hiểu là VietCredit làm sai khi chưa xác minh (ví dụ "việc liên hệ người thân cần tuân thủ đúng quy định pháp luật" dưới bài tố nhắc nợ sai).',
  'Không công kích, dọa dẫm, lên lớp người đăng. Không nêu thông tin cá nhân, khoản vay của khách hàng.',
  'Không gợi ý cách trốn nợ; không hứa ưu đãi, giãn nợ, miễn giảm mà VietCredit chưa xác nhận.',
  'Các bình luận trong cùng một bài khác nhau về câu chữ, đăng giãn cách, từ các tài khoản khác nhau. Số bình luận mỗi bài theo đề xuất từng case.',
]);

C.push(SEC('B5. Cập nhật link theo dõi'));
C.push(P('Link hiện có hai sheet. VietCredit đề nghị bổ sung các cột in đậm dưới đây để hai bên duyệt và nghiệm thu ngay trên link.'));
C.push(SUB('Sheet "Tư vấn XLKH" (một dòng một case)'));
T([2700, 1500, 4826], ['Cột', 'Bên ghi', 'Cách ghi'], [
  ['STT, Ngày tổng hợp', 'MHC', 'STT liên tục; ngày ghi theo dd/mm/yyyy'],
  ['Nền tảng, Link', 'MHC', 'Link trực tiếp tới bài viết'],
  ['Nhóm', 'MHC', '1, 2, 3, 4 hoặc "Khác"'],
  ['**Tương tác, Cấp độ**', 'MHC', 'Lượt thích, bình luận, chia sẻ lúc ghi nhận; Cấp độ 1, 2, 3'],
  ['Đánh giá case', 'MHC', 'Tóm tắt nội dung và mức ảnh hưởng'],
  ['Phương án tư vấn/xử lý', 'MHC', 'Phương án và lý do; với report, gỡ bài: chi phí, thời hạn'],
  ['Bên triển khai', 'MHC', 'MHC / VietCredit (CS, Thu hồi nợ)'],
  ['**Ý kiến VC**', 'VietCredit', 'Đồng ý / Đồng ý có điều chỉnh / Đổi phương án / Không xử lý, kèm nội dung'],
  ['**Người duyệt, ngày duyệt**', 'VietCredit', 'Tên người ở Marketing ghi ý kiến'],
  ['Tình trạng xử lý', 'Hai bên', 'Chọn một trong các tình trạng ở mục A4'],
  ['**Kết quả nghiệm thu**', 'VietCredit', 'Đạt / Đạt có góp ý / Cần bổ sung'],
  ['**Yêu cầu bổ sung, hạn**', 'VietCredit', 'Việc cần làm thêm và hạn hoàn thành'],
  ['Ghi chú', 'Hai bên', ''],
], { boldFirst: true });
C.push(SUB('Sheet "Seeding" (một dòng một bình luận)'));
T([2700, 1500, 4826], ['Cột', 'Bên ghi', 'Cách ghi'], [
  ['STT, Link bài', 'MHC', 'Trùng STT với sheet "Tư vấn XLKH"'],
  ['Comment', 'MHC', 'Nội dung bình luận dự kiến, có nhãn hướng nội dung'],
  ['Nghiệm thu', 'MHC', 'Link trực tiếp tới bình luận đã đăng hoặc ảnh chụp màn hình, kèm thời điểm đăng'],
  ['**Kiểm tra sau 48 giờ**', 'MHC', 'Còn hiển thị / Bị ẩn / Bị xóa'],
  ['**Nhận xét VC**', 'VietCredit', 'Đạt / Cần sửa, lý do'],
], { boldFirst: true });

C.push(SEC('B6. Nghiệm thu'));
C.push(SUB('Thời điểm'));
BLS([
  'Case đã qua VietCredit duyệt (Nhóm 2, 3, ngoài nhóm, report, gỡ bài): Marketing nghiệm thu trong 1 ngày làm việc sau khi MHC chuyển sang "Chờ nghiệm thu".',
  'Nhóm 1 và Nhóm 4: Marketing nghiệm thu theo tuần, vào thứ Sáu, kiểm tra ngẫu nhiên ít nhất 20% số case của tuần.',
  'Cuối tháng: hai bên đối chiếu số case, số bình luận, số report, gỡ bài đã nghiệm thu "Đạt" để làm cơ sở nghiệm thu gói dịch vụ và thanh toán chi phí ngoài gói.',
]);
C.push(SUB('Tiêu chí'));
T([1900, 7126], ['Phương án', 'Đạt khi'], [
  ['Seeding', ['• Bình luận đăng đúng bài, đúng nội dung đã duyệt hoặc kịch bản mẫu, đúng chuẩn mục B4;', '• đủ số lượng đã đề xuất;', '• có bằng chứng (link bình luận hoặc ảnh chụp);', '• còn hiển thị sau 48 giờ.']],
  ['Report', 'Bài bị gỡ hoặc hạn chế trong thời hạn MHC đã cam kết; MHC có bằng chứng.'],
  ['Gỡ bài', 'Bài bị gỡ trong thời hạn cam kết và không xuất hiện lại trong thời gian bảo hành.'],
  ['Theo dõi', 'MHC cập nhật diễn biến theo mốc đã hẹn và báo ngay khi tương tác tăng.'],
], { boldFirst: true });
C.push(SUB('Kết quả'));
T([2000, 7026], ['Kết quả', 'Ý nghĩa'], [
  ['Đạt', 'Case chuyển "Hoàn thành".'],
  ['Đạt có góp ý', 'Case chuyển "Hoàn thành". MHC áp dụng góp ý cho các case sau.'],
  ['Cần bổ sung', 'MHC xử lý theo mục B7. Không tính vào khối lượng nghiệm thu cho đến khi đạt.'],
], { boldFirst: true });

C.push(SEC('B7. Yêu cầu bổ sung'));
C.push(P('VietCredit có thể yêu cầu bổ sung khi nghiệm thu chưa đạt, hoặc khi case diễn biến thêm. Các loại yêu cầu:'));
BLS([
  'Sửa hoặc xóa bình luận sai nội dung, sai chuẩn.',
  'Thêm bình luận khi thảo luận tiêu cực tăng.',
  'Đổi phương án (ví dụ từ seeding sang report).',
  'Tiếp tục theo dõi và báo cáo thêm.',
  'Đăng lại khi bình luận bị ẩn, bị xóa.',
]);
T([700, 3400, 4926], ['Bước', 'Việc', 'Thời hạn'], [
  ['1', 'VietCredit ghi yêu cầu vào cột "Yêu cầu bổ sung, hạn", chuyển tình trạng "Cần bổ sung", báo qua nhóm chat', 'Ngay khi nghiệm thu'],
  ['2', 'MHC xác nhận đã nhận', 'Trong 2 giờ làm việc'],
  ['3', 'MHC thực hiện, cập nhật bằng chứng, chuyển lại "Chờ nghiệm thu"', 'Trong 1 ngày làm việc, hoặc theo hạn VietCredit ghi'],
  ['4', 'VietCredit nghiệm thu lại', 'Trong 1 ngày làm việc'],
], { boldFirst: true });
C.push(P('Bình luận bị ẩn, bị xóa trong 48 giờ đầu và bài gỡ xuất hiện lại trong thời gian bảo hành thì MHC xử lý lại, không tính thêm vào quota hoặc chi phí.'));

C.push(SEC('B8. Sự vụ khẩn'));
C.push(P('VietCredit đề nghị bổ sung luồng này vì quy trình MHC gửi chưa nêu. Áp dụng với sự vụ Cấp độ 1: tin về lãnh đạo, hoạt động kinh doanh, bảo mật, lỗi sản phẩm hàng loạt, truyền thông sai thực tế, hoặc bài có lan tỏa cao (trên 100 tương tác, hoặc lên báo điện tử). Bảng xếp cấp độ chi tiết ở Phụ lục 1.'));
T([2900, 1700, 4426], ['Mốc', 'Ai', 'Việc'], [
  ['Trong 30 phút từ khi phát hiện', 'MHC', 'Gọi điện và nhắn nhóm chat khẩn, kể cả ngoài giờ và cuối tuần. Gửi Form sự vụ (Phụ lục 2) ngay sau đó'],
  ['Trong 1 giờ', 'Marketing', 'Xác nhận đã nhận'],
  ['Trong 4 giờ', 'VietCredit', 'Chốt phương án'],
  ['Theo mốc hai bên thống nhất', 'MHC', 'Cập nhật diễn biến đến khi đóng sự vụ'],
], { boldFirst: true });

C.push(SEC('B9. Báo cáo định kỳ MHC gửi'));
T([1800, 7226], ['Báo cáo', 'Nội dung'], [
  ['Hằng tuần (thứ Hai)', 'Số case theo sản phẩm, nền tảng, nhóm, cấp độ; tình trạng xử lý; số bình luận đã đăng và tỷ lệ còn hiển thị; case nổi bật; quota đã dùng'],
  ['Hằng tháng (ngày làm việc thứ 3)', 'Như báo cáo tuần, thêm: xu hướng so với tháng trước; tỷ lệ tiêu cực trong thảo luận; kết quả report, gỡ bài; các vấn đề khách hàng phàn nàn nhiều nhất; bảng đối chiếu khối lượng đã nghiệm thu'],
], { boldFirst: true });

// ===== PHẦN C =====
PART('PHẦN C', 'QUY TRÌNH NỘI BỘ VIETCREDIT',
  'Phần này dùng nội bộ, không gửi MHC. Sau giai đoạn thí điểm, Marketing hoàn thiện để ban hành thành quy trình nội bộ.');

C.push(SEC('C1. Vai trò các đơn vị'));
T([1700, 7326], ['Đơn vị', 'Trách nhiệm'], [
  ['Marketing', ['• Đầu mối duy nhất làm việc với MHC.', '• Review đề xuất hằng ngày, lấy ý kiến phòng ban, ghi ý kiến lên link.', '• Nghiệm thu, nhận xét, yêu cầu bổ sung.', '• Duyệt kịch bản seeding mẫu; theo dõi quota và chi phí.', '• Báo BĐH sự vụ Cấp độ 1; báo cáo tháng.']],
  ['PD/PM sản phẩm', ['• Xác nhận lỗi tính năng, thông tin phí, chính sách, điều kiện sản phẩm.', '• Báo trước cho Marketing các thay đổi có thể gây phản ứng.', '• Khắc phục các lỗi lặp lại.']],
  ['CS', ['• Tra cứu khách hàng khi bài viết đủ thông tin nhận diện.', '• Liên hệ khách hàng khi phương án là "Liên hệ người đăng".', '• Báo Marketing khi tổng đài quá tải hoặc nhiều khiếu nại cùng một vấn đề.']],
  ['Thu hồi nợ', ['• Kiểm tra các phản ánh về cách nhắc nợ, đòi nợ (gọi người thân, đến nhà, đe dọa).', '• Cho ý kiến hướng trả lời các câu hỏi Nhóm 4.', '• Báo Marketing khi có thay đổi trong quy trình nhắc nợ.']],
  ['Pháp chế', ['• Cho ý kiến khi bài viết có dấu hiệu vu khống, mạo danh thương hiệu, lộ thông tin khách hàng.', '• Soạn yêu cầu gỡ bài chính thức khi cần.']],
  ['BĐH', ['• Chỉ đạo sự vụ Cấp độ 1; duyệt phát ngôn và chi phí vượt hạn mức.']],
], { boldFirst: true });

C.push(SEC('C2. Review đề xuất hằng ngày'));
C.push(SUB('Các bước và thời hạn'));
T([1900, 1900, 5226], ['Thời hạn', 'Ai', 'Việc'], [
  ['13:30', 'MHC', 'Cập nhật link phương án'],
  ['Trước 14:30', 'Marketing', 'Review theo checklist bên dưới. Case cần ý kiến chuyên môn thì gửi đơn vị theo bảng "Hỏi ai", kèm link bài và đề xuất của MHC'],
  ['Trước 16:00', 'Đơn vị được hỏi', 'Trả lời: thông tin đúng hay sai, hướng trả lời, có cần liên hệ khách hàng không'],
  ['Trước 16:30', 'Marketing', 'Ghi ý kiến, người duyệt, ngày duyệt lên link; chuyển tình trạng; báo MHC'],
], { boldFirst: true });
C.push(P('Đơn vị không trả lời trước 16:00: Marketing duyệt phương án an toàn (theo dõi thêm, hoặc bình luận mời khách hàng liên hệ CSKH qua hotline 1900 6515) và báo trưởng đơn vị. Hướng chi tiết gửi MHC bổ sung khi có ý kiến.'));
C.push(SUB('Checklist review mỗi case'));
NUMS([
  'MHC phân nhóm đúng chưa? Case tố VietCredit làm sai, case bên thứ ba dùng tên VietCredit, case lan tỏa cao có bị xếp nhầm vào Nhóm 1, 4 để tự seeding không?',
  'Thông tin trong bài đúng hay sai? Có cần CS, Thu hồi nợ tra cứu khách hàng không?',
  'Phương án có hợp lý không? Ví dụ: bài quảng cáo dịch vụ rút tiền thẻ dùng tên VietCredit nên report hơn là seeding.',
  'Nội dung bình luận đúng chuẩn mục B4 chưa: tên thương hiệu, không ám chỉ VietCredit sai, không đăng nhãn, không lên lớp khách hàng?',
  'Có phát sinh chi phí ngoài gói không? Nếu có thì duyệt theo mục C6.',
]);
C.push(SUB('Hỏi ai'));
T([4026, 5000], ['Nội dung case', 'Đơn vị cho ý kiến'], [
  ['Lỗi tính năng, ứng dụng, không thao tác được', 'PD/PM sản phẩm'],
  ['Phí, lãi, hạn mức, điều kiện, chính sách', 'PD/PM sản phẩm'],
  ['Nhắc nợ, đòi nợ, liên hệ người thân, đến nhà', 'Thu hồi nợ; CS nếu cần liên hệ khách hàng'],
  ['Khách hàng cụ thể phàn nàn về dịch vụ', 'CS'],
  ['Mạo danh, bên thứ ba dùng tên VietCredit, vu khống', 'Pháp chế'],
  ['Doanh nghiệp, tình hình kinh doanh, lãnh đạo', 'Marketing trình BĐH'],
], { boldFirst: true });

C.push(SEC('C3. Nghiệm thu nội bộ'));
T([2600, 6426], ['Việc', 'Cách làm'], [
  ['Nghiệm thu case đã duyệt', 'Marketing mở link bằng chứng, kiểm tra theo tiêu chí mục B6, ghi kết quả trong 1 ngày làm việc'],
  ['Kiểm tra mẫu Nhóm 1, 4', 'Thứ Sáu, chọn ngẫu nhiên ít nhất 20% case trong tuần. Nếu trên 1/5 số case kiểm tra không đạt, tuần sau kiểm tra 100%'],
  ['Kiểm tra sau 48 giờ', 'Đối chiếu cột "Kiểm tra sau 48 giờ" của MHC với một vài bình luận thực tế'],
  ['Đối chiếu cuối tháng', 'Đối chiếu bảng khối lượng MHC gửi với kết quả nghiệm thu trên link trước khi xác nhận thanh toán'],
  ['Đánh giá MHC hằng tháng', 'Tỷ lệ case đạt ngay lần đầu, tỷ lệ đúng hạn, số yêu cầu bổ sung, số lần phân nhóm sai. Gửi nhận xét cho MHC trong buổi họp tháng'],
], { boldFirst: true });

C.push(SEC('C4. Chuyển thông tin để xử lý tận gốc'));
C.push(P('Seeding chỉ xử lý phần hiển thị trên mạng. Khi review hoặc nghiệm thu thấy vấn đề có thật, Marketing chuyển cho đơn vị phụ trách:'));
BLS([
  'Phản ánh về cách nhắc nợ (gọi người thân, đe dọa): chuyển Thu hồi nợ kiểm tra cuộc gọi, nhân sự hoặc đối tác nhắc nợ liên quan.',
  'Nhiều phản ánh hotline không gọi được: chuyển CS kiểm tra năng lực tổng đài.',
  'Lỗi tính năng lặp lại: chuyển PD/PM, theo dõi đến khi sửa xong.',
  'Bên thứ ba quảng cáo dịch vụ dùng tên VietCredit: chuyển Pháp chế, cân nhắc cảnh báo khách hàng trên kênh chính thức.',
]);

C.push(SEC('C5. Sự vụ Cấp độ 1'));
NUMS([
  'Trong 1 giờ từ khi MHC báo: Marketing xác nhận với MHC, báo BĐH, gọi nhóm xử lý (Trưởng Marketing, PD/PM liên quan, Trưởng CS, Thu hồi nợ nếu liên quan, Pháp chế, thành viên BĐH phụ trách).',
  'Các đơn vị kiểm tra sự việc: đúng hay sai, bao nhiêu khách hàng bị ảnh hưởng, tổng đài đã nhận khiếu nại chưa.',
  'Trong 4 giờ: nhóm xử lý đề xuất; BĐH duyệt phương án, nội dung phát ngôn, chi phí.',
  'Marketing gửi phương án cho MHC; CS dùng cùng nội dung để trả lời khách hàng gọi tổng đài.',
  'Trong 5 ngày làm việc sau khi đóng sự vụ: Marketing báo cáo BĐH nguyên nhân, cách xử lý, chi phí, việc cần sửa.',
]);
C.push(P('Trong sự vụ: chỉ người được BĐH giao mới phát ngôn; nhân viên không tự bình luận trên tài khoản cá nhân; không xóa bình luận hàng loạt trên fanpage VietCredit khi chưa có ý kiến nhóm xử lý.'));

C.push(SEC('C6. Duyệt chi phí ngoài gói'));
T([3500, 5526], ['Khoản', 'Người duyệt'], [
  ['Seeding, theo dõi trong quota gói', 'Không cần duyệt riêng'],
  ['Report, gỡ bài đến [hạn mức, BĐH chốt] mỗi lần', 'Trưởng Marketing'],
  ['Vượt hạn mức, hoặc sự vụ Cấp độ 1', 'BĐH'],
], { boldFirst: true });

C.push(SEC('C7. Phòng ngừa chủ động'));
NUMS([
  '**Bộ từ khóa.** PD/PM cung cấp tên sản phẩm, cách khách hàng hay gọi, tên đối tác (Điện Máy Xanh, TikTok Shop). Marketing gửi MHC, rà lại mỗi quý.',
  '**Kịch bản mẫu.** Marketing duyệt bộ bình luận mẫu cho Nhóm 1 và Nhóm 4; Thu hồi nợ, CS xác nhận nội dung. Bổ sung kịch bản sau mỗi case mới đã duyệt.',
  '**Báo trước.** Trước khi bảo trì hệ thống, đổi phí, chính sách, chạy chiến dịch lớn, PD/PM báo Marketing ít nhất 2 ngày làm việc (lỗi đột xuất thì báo ngay) để Marketing chuẩn bị câu trả lời và báo MHC.',
  '**Theo dõi khi ra mắt.** 2 tuần đầu sau khi ra mắt sản phẩm, tính năng, chiến dịch: Marketing yêu cầu MHC báo cáo nhanh hằng ngày cho sản phẩm đó.',
  '**Họp tháng.** Marketing, PD/PM, CS, Thu hồi nợ xem 5 vấn đề bị phàn nàn nhiều nhất trong báo cáo tháng của MHC, thống nhất việc khắc phục và thời hạn.',
]);

C.push(SEC('C8. Báo cáo Ban Điều hành'));
C.push(P('Marketing gửi BĐH báo cáo tháng, gồm: số case theo sản phẩm, nhóm, cấp độ; tỷ lệ tiêu cực trong thảo luận; tỷ lệ case được duyệt đúng hạn; kết quả nghiệm thu và đánh giá MHC; chi phí ngoài gói; các vấn đề lặp lại và tiến độ khắc phục.'));

// ===== PHẦN D =====
PART('PHẦN D', 'TRIỂN KHAI VÀ CÁC NỘI DUNG CẦN CHỐT');

C.push(SEC('D1. Kế hoạch triển khai'));
T([2000, 4826, 2200], ['Thời gian', 'Việc', 'Thực hiện'], [
  ['Tuần 1', 'BĐH duyệt quy trình. Gửi Phần B cho MHC; thống nhất các điểm ở mục D3. Bổ sung cột vào link theo dõi. Lập nhóm chat khẩn', 'Marketing'],
  ['Tuần 1 đến 2', 'Rà lại kịch bản mẫu Nhóm 1, 4 theo chuẩn B4; chốt bộ từ khóa cho VC x DMX và TikTok BNPL', 'Marketing, PD/PM, CS, Thu hồi nợ, MHC'],
  ['Tuần 3 đến 6', 'Chạy thí điểm, họp 15 phút mỗi tuần với MHC', 'Marketing, MHC'],
  ['Cuối tuần 6', 'Đánh giá thí điểm; ban hành Phần C thành quy trình nội bộ; đề xuất sản phẩm áp dụng tiếp theo', 'Marketing'],
], { boldFirst: true });

C.push(SEC('D2. Nội dung cần BĐH quyết định'));
NUMS([
  'Phạm vi giai đoạn 1: Thẻ tín dụng VC x DMX và TikTok BNPL; các dự án khác áp dụng luồng khẩn cho Cấp độ 1.',
  'Cho phép MHC tự seeding Nhóm 1 và Nhóm 4 theo kịch bản đã duyệt, trừ các trường hợp ở mục B2.',
  'Hạn mức chi phí report, gỡ bài Trưởng Marketing được duyệt mỗi lần.',
  'Người phát ngôn và thành viên BĐH phụ trách khi có sự vụ Cấp độ 1.',
]);

C.push(SEC('D3. Điểm cần thống nhất thêm với MHC'));
C.push(P('Quy trình MHC gửi ngày 12/10/2026 được giữ nguyên về khung giờ ghi nhận (9h00 đến 9h00), bốn nhóm seeding và việc cập nhật kết quả lên link. VietCredit đề nghị bổ sung:'));
NUMS([
  'Mốc cố định 13:30 cho link phương án (thay cho "đầu giờ chiều").',
  'Các trường hợp không tự seeding dù thuộc Nhóm 1, 4, và nhóm "Ngoài 4 nhóm" (mục B2).',
  'Chuẩn nội dung bình luận (mục B4).',
  'Các cột mới trên link: Tương tác, Cấp độ, Ý kiến VC, Người duyệt, Kết quả nghiệm thu, Yêu cầu bổ sung, Kiểm tra sau 48 giờ, Nhận xét VC (mục B5).',
  'Tiêu chí nghiệm thu và quy trình yêu cầu bổ sung (mục B6, B7).',
  'Luồng sự vụ khẩn (mục B8) và báo cáo tuần, tháng (mục B9).',
]);

// ===== PHỤ LỤC =====
PART('PHỤ LỤC', 'TIÊU CHÍ CẤP ĐỘ VÀ MẪU FORM');
C.push(SEC('Phụ lục 1. Xếp cấp độ sự vụ'));
C.push(P('Ghép loại thông tin với mức lan tỏa. Lan tỏa cao: trên 100 tương tác hoặc lên báo điện tử. Trung bình: 50 đến 100. Thấp: dưới 50. "Nghiêm trọng" là Cấp độ 1, "Vừa" là Cấp độ 2, "Thấp" là Cấp độ 3.'));
T([4826, 1400, 1400, 1400], ['Loại thông tin tiêu cực', 'Lan tỏa cao', 'Lan tỏa trung bình', 'Lan tỏa thấp'], [
  ['Lãnh đạo; bảo mật, bí mật kinh doanh; hoạt động kinh doanh; truyền thông sai thực tế; lỗi sản phẩm mang tính hệ thống; nhân viên vi phạm pháp luật khi làm việc', 'Nghiêm trọng', 'Nghiêm trọng', 'Nghiêm trọng'],
  ['Phàn nàn về sản phẩm, dịch vụ của một hoặc một nhóm khách hàng', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Quy định, chính sách, thủ tục; thái độ nhân viên; nhắc nợ, thu hồi nợ', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Tuyển dụng, đãi ngộ, điều kiện làm việc của nhiều nhân viên', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Đầu tư, đấu thầu; cơ sở vật chất, công nghệ lạc hậu', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Nhân viên vi phạm luật giao thông, đánh nhau, phát ngôn nhạy cảm liên quan công việc', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Chia sẻ lại link website tiêu cực', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Chia sẻ lại bài tiêu cực đã được cảnh báo', 'Thấp', 'Thấp', 'Thấp'],
  ['Đời tư nhân viên vi phạm pháp luật, không liên quan công việc', 'Thấp', 'Thấp', 'Thấp'],
]);
C.push(SEC('Phụ lục 2. Form sự vụ lớn'));
T([2300, 6726], ['Mục', 'Nội dung'], [
  ['0. Tên sự vụ', 'Tên ngắn, sản phẩm liên quan, cấp độ'],
  ['1. Tổng quan', 'Tổng số bài đăng; số bài trên từng nền tảng'],
  ['2. Nguồn đăng', 'Tên, link; số điện thoại, email, kênh khác (nếu có)'],
  ['3. Nội dung', ['• Nội dung sự vụ', '• Tương tác hiện tại, thời điểm thảo luận gần nhất', '• Tỷ lệ bình luận tiêu cực, trung tính, tích cực và ý chính', '• Các bài nổi bật']],
  ['4. Đánh giá', 'Mức lan tỏa, mức ảnh hưởng, lý do xếp cấp độ'],
  ['5. Đề xuất', ['• Phương án và nội dung', '• Với report, gỡ bài: chi phí, thời gian, bảo hành, tỷ lệ thành công', '• Việc VietCredit nên tránh trong lúc chờ phương án']],
], { boldFirst: true });

// ---------- document ----------
const doc = new Document({
  creator: 'Marketing VietCredit',
  title: 'Quy trình quản trị rủi ro truyền thông và xử lý thông tin tiêu cực',
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
