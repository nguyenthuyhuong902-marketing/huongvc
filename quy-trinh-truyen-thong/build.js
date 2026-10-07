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
const H3 = t => new Paragraph({ spacing: { before: 160, after: 80 }, keepNext: true,
  children: [new TextRun({ text: t, bold: true, italics: true, font: FONT, size: 21, color: DARK })] });
// Thanh đỏ đánh dấu đầu mỗi phần; phần B, C sang trang mới để tách riêng được
const PART = (code, title, note, newPage = true) => {
  C.push(new Paragraph({
    heading: HeadingLevel.HEADING_1, pageBreakBefore: newPage, keepNext: true,
    spacing: { before: newPage ? 0 : 320, after: note ? 0 : 160 },
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: RED },
    indent: { left: 0, right: 0 },
    children: [new TextRun({ text: ` ${code}. ${title}`, bold: true, font: FONT, size: 26, color: 'FFFFFF' })],
  }));
  if (note) C.push(new Paragraph({ spacing: { before: 80, after: 160 }, keepNext: true,
    children: runs(note, { italics: true, color: GREY, size: 19 }) }));
};
const SEC = t => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 280, after: 120 }, keepNext: true,
  children: [new TextRun({ text: t, bold: true, font: FONT, size: 24, color: RED })] });
const SUB = t => new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 200, after: 100 }, keepNext: true,
  children: [new TextRun({ text: t, bold: true, font: FONT, size: 22, color: DARK })] });
let numRef = 0;
const NUMS = items => { numRef++; items.forEach(t => C.push(NL(t, 'n' + numRef))); };
const NOTE = t => P(t, { italics: true, color: GREY, size: 19 });

// ===== Bìa =====
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 240, after: 80 }, children: [
  new TextRun({ text: 'QUY TRÌNH', font: FONT, size: 32, bold: true, color: RED })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [
  new TextRun({ text: 'QUẢN TRỊ RỦI RO TRUYỀN THÔNG VÀ XỬ LÝ THÔNG TIN TIÊU CỰC TRÊN MẠNG XÃ HỘI', font: FONT, size: 26, bold: true, color: RED })] }));
C.push(table([2600, 6426], ['Thông tin', 'Nội dung'], [
  ['Tình trạng', 'Dự thảo lần 2, trình Ban Điều hành'],
  ['Ngày soạn', '07/10/2026'],
  ['Đơn vị chủ trì', 'Marketing VietCredit'],
  ['Đơn vị phối hợp', 'Quản lý sản phẩm (PD/PM), Chăm sóc khách hàng (CS), Pháp chế, đối tác Màn Hình Cộng (MHC)'],
  ['Sản phẩm áp dụng giai đoạn 1', 'Thẻ tín dụng VietCredit x Điện Máy Xanh (VC x DMX); TikTok BNPL'],
], { boldFirst: true }));
C.push(gap());
C.push(SUB('Tài liệu gồm 4 phần'));
C.push(table([1000, 4426, 3600], ['Phần', 'Nội dung', 'Người đọc'], [
  ['A', 'Tổng quan: mục đích, phạm vi, sơ đồ quy trình trên một trang', 'Ban Điều hành, tất cả các bên'],
  ['B', 'Phối hợp giữa VietCredit và MHC: tiêu chí phân loại, quyền tự xử lý, lịch gửi file, thời hạn, báo cáo, mẫu biểu', 'MHC và Marketing. Phần này gửi MHC làm thỏa thuận làm việc'],
  ['C', 'Quy trình nội bộ VietCredit: ai cho ý kiến, thời hạn nội bộ, xử lý sự vụ nghiêm trọng, duyệt chi phí, phòng ngừa', 'Marketing, PD/PM, CS, Pháp chế, BĐH. Không gửi MHC; sau này ban hành thành quy trình nội bộ'],
  ['D', 'Kế hoạch triển khai và các nội dung cần Ban Điều hành quyết định', 'Ban Điều hành'],
], { boldFirst: true }));

// ===== PHẦN A =====
PART('PHẦN A', 'TỔNG QUAN', null);

C.push(SEC('A1. Mục đích'));
C.push(BL('Phát hiện sớm thông tin tiêu cực về VietCredit và sản phẩm trên mạng xã hội, báo điện tử.'));
C.push(BL('Quy định rõ việc nào MHC được làm ngay, việc nào phải hỏi VietCredit, việc nào phải báo Ban Điều hành.'));
C.push(BL('Có phương án trong ngày cho tin thông thường và trong 4 giờ cho sự vụ nghiêm trọng.'));
C.push(BL('Chuyển các phàn nàn lặp lại về cho PD/PM, CS để sửa tận gốc.'));

C.push(SEC('A2. Phạm vi áp dụng'));
C.push(table([2200, 6826], ['Hạng mục', 'Áp dụng'], [
  ['Sản phẩm giai đoạn 1', 'Thẻ tín dụng VC x DMX; TikTok BNPL. Áp dụng toàn bộ quy trình.'],
  ['Các dự án khác', 'Các dự án đang có báo cáo Social listening: MHC tiếp tục theo dõi, tổng hợp như hiện tại. Sự vụ nghiêm trọng (Cấp độ 1) của bất kỳ dự án nào đều xử lý theo luồng khẩn.'],
  ['Mở rộng sau này', ['Marketing đề xuất BĐH bổ sung sản phẩm đáp ứng ít nhất một tiêu chí:',
    '• VietCredit chịu toàn bộ rủi ro của dự án;',
    '• dự án có margin tốt;',
    '• dự án đang thử nghiệm hoặc có vấn đề cần đo lường, giám sát.']],
  ['Kênh theo dõi', 'Facebook (fanpage, nhóm, trang cá nhân công khai), TikTok, YouTube, Threads, báo điện tử và trang tin.'],
], { boldFirst: true }));

C.push(SEC('A3. Thuật ngữ'));
C.push(table([2300, 6726], ['Thuật ngữ', 'Giải thích'], [
  ['MHC', 'Màn Hình Cộng, đối tác cung cấp dịch vụ Social listening và gói quản trị tiêu cực chủ động cho VietCredit'],
  ['PD/PM', 'Đơn vị phát triển và quản lý sản phẩm phụ trách sản phẩm liên quan'],
  ['CS', 'Chăm sóc khách hàng'],
  ['BĐH', 'Ban Điều hành'],
  ['Seeding', 'Viết bình luận để giải thích thông tin, định hướng và pha loãng thảo luận tiêu cực dưới bài viết công khai'],
  ['Report', 'Báo cáo bài viết với nền tảng để hạn chế tương tác, hạn chế bài tiếp cận người khác'],
  ['Gỡ bài', 'Dùng biện pháp kỹ thuật để bài viết bị gỡ khỏi nền tảng'],
  ['Alert OTT', 'Cảnh báo khẩn qua nhóm chat (Zalo, Telegram...) khi có sự vụ lớn'],
  ['Tương tác', 'Tổng lượt thích, bình luận và chia sẻ của một bài viết'],
], { boldFirst: true }));

C.push(SEC('A4. Quy trình trên một trang'));
C.push(new Paragraph({ keepNext: true, spacing: { after: 120 }, children: runs('Mỗi tin tiêu cực đi vào một trong ba luồng. Cách xếp luồng ở mục B3.') }));
C.push(table([1500, 3000, 2000, 2526], ['Luồng', 'Khi nào', 'Ai quyết định', 'Thời hạn'], [
  ['A. MHC tự xử lý', 'Tin mức thấp thuộc nhóm đã có kịch bản duyệt sẵn (hotline, "lừa đảo", hỏi bùng nợ)', 'MHC, theo kịch bản đã duyệt', 'Xử lý ngay, báo lại trong file hằng ngày'],
  ['B. Hỏi ý kiến VietCredit', 'Tin mức trung bình, hoặc nhóm chưa có kịch bản (lỗi tính năng, tin về doanh nghiệp, lãnh đạo...)', 'Marketing, sau khi hỏi PD/PM, CS', 'MHC gửi lúc 10:30, VC trả lời trước 14:00, xử lý xong trong ngày'],
  ['C. Khẩn', 'Sự vụ nghiêm trọng (Cấp độ 1)', 'BĐH', 'Cảnh báo trong 30 phút, có phương án trong 4 giờ'],
], { boldFirst: true }));
C.push(gap());
C.push(new Paragraph({ keepNext: true, spacing: { after: 120 }, children: runs('Ai làm gì theo từng bước:') }));
C.push(table([1700, 2600, 2400, 2326], ['Bước', 'MHC', 'Marketing VC', 'PD/PM, CS, Pháp chế, BĐH'], [
  ['1. Phát hiện', 'Theo dõi liên tục, ghi nhận tin tiêu cực', '', ''],
  ['2. Phân loại', 'Xếp nhóm nội dung và cấp độ, chọn luồng', '', ''],
  ['3. Xử lý Luồng A', 'Seeding theo kịch bản, ghi kết quả vào file 10:30', 'Xem file, không cần duyệt', ''],
  ['4. Xử lý Luồng B', 'Gửi case kèm đề xuất trong file 10:30; triển khai khi có phương án', 'Chuyển đơn vị liên quan, chốt phương án, gửi lại MHC trước 14:00', 'Cho ý kiến trước 13:30'],
  ['5. Xử lý Luồng C', 'Alert OTT trong 30 phút, gửi Form đề xuất, cập nhật diễn biến', 'Xác nhận trong 1 giờ, báo BĐH, họp nhóm xử lý', 'BĐH quyết phương án, duyệt phát ngôn và chi phí'],
  ['6. Sau xử lý', 'Theo dõi bài viết, báo cáo tuần, tháng', 'Báo cáo BĐH, chuyển vấn đề lặp lại cho PD/PM', 'PD/PM, CS sửa nguyên nhân gốc'],
], { boldFirst: true }));

// ===== PHẦN B =====
PART('PHẦN B', 'PHỐI HỢP GIỮA VIETCREDIT VÀ MHC',
  'Phần này gửi MHC. Đây là thỏa thuận làm việc giữa hai bên: MHC làm gì, được tự quyết việc gì, gửi gì cho VietCredit và khi nào.');

C.push(SEC('B1. Đầu mối liên lạc'));
C.push(table([2200, 3400, 3426], ['Bên', 'Đầu mối', 'Kênh'], [
  ['VietCredit', 'Marketing VietCredit: [tên, số điện thoại]. MHC chỉ làm việc qua đầu mối này, không liên hệ trực tiếp các phòng ban khác.', 'Email cho file hằng ngày, báo cáo; nhóm OTT cho cảnh báo khẩn'],
  ['MHC', 'Hằng, Lam (theo đề xuất của MHC)', 'Như trên'],
  ['Nhóm OTT cảnh báo khẩn', 'Đầu mối hai bên và người được Marketing bổ sung', '[Zalo hoặc Telegram, hai bên chốt]'],
], { boldFirst: true }));

C.push(SEC('B2. Việc MHC thực hiện'));
NUMS([
  'Theo dõi liên tục trên hệ thống Social listening theo bộ từ khóa VietCredit cung cấp.',
  'Phân loại mỗi tin tiêu cực theo mục B3.',
  'Tự xử lý các tin thuộc Luồng A theo kịch bản đã duyệt, không cần hỏi lại.',
  'Gửi file tổng hợp lúc 10:30 các ngày làm việc (Mẫu 01), gồm cả tin đã tự xử lý và tin cần ý kiến VietCredit kèm đề xuất.',
  'Gửi Alert OTT và Form đề xuất (Mẫu 02) khi có sự vụ Cấp độ 1, kể cả ngoài giờ.',
  'Triển khai phương án VietCredit đã chốt.',
  'Theo dõi bài viết sau xử lý, báo ngay nếu bài xuất hiện lại hoặc lan sang kênh khác.',
  'Gửi báo cáo tuần, tháng (Mẫu 03).',
  'Trong lúc chờ phương án chính thức, tư vấn VietCredit những việc nên tránh.',
]);

C.push(SEC('B3. Phân loại tin tiêu cực'));
C.push(P('MHC xếp mỗi tin theo ba bước: xác định nhóm nội dung, đo mức lan tỏa, rồi ra cấp độ. Cấp độ quyết định luồng xử lý.'));

C.push(SUB('Bước 1. Nhóm nội dung'));
C.push(table([1100, 3300, 1900, 2726], ['Nhóm', 'Nội dung', 'Quyền xử lý', 'Cách làm'], [
  ['Nhóm 1', 'Gọi hotline mãi không được; cho rằng "VietCredit lừa đảo"', 'MHC tự xử lý', 'Seeding theo kịch bản mẫu: VietCredit đảm bảo dịch vụ nhanh chóng, minh bạch, an toàn; khách hàng cần hỗ trợ xin liên hệ CSKH VietCredit, hotline 1900 6515.'],
  ['Nhóm 2', 'Lỗi tính năng, không thao tác được (mới phát sinh)', 'Hỏi VietCredit', 'Đề xuất hướng trả lời trong file 10:30.'],
  ['Nhóm 3', 'Tin tiêu cực về doanh nghiệp, tình hình kinh doanh, lãnh đạo', 'Hỏi VietCredit', 'Đề xuất hướng trả lời. Nếu là Cấp độ 1 thì chuyển Luồng C.'],
  ['Nhóm 4', 'Hỏi bùng nợ được không, có về nhà đòi nợ không, nên vay bên nào', 'MHC tự xử lý trong quota còn lại của gói', 'Dùng bộ kịch bản đang áp dụng. Không cần xử lý 100% số tin; ưu tiên bài nhiều tương tác.'],
  ['Nhóm 5', 'Nội dung khác, chưa có kịch bản', 'Hỏi VietCredit', 'Đề xuất hướng trả lời. Kịch bản được chốt sẽ thêm vào thư viện để lần sau MHC tự xử lý.'],
], { boldFirst: true }));

C.push(SUB('Bước 2. Mức lan tỏa'));
C.push(table([2000, 7026], ['Mức', 'Tiêu chí'], [
  ['Cao', 'Trên 100 tương tác, hoặc bài xuất hiện trên báo điện tử, trang tin'],
  ['Trung bình', 'Từ 50 đến 100 tương tác'],
  ['Thấp', 'Dưới 50 tương tác'],
], { boldFirst: true }));
C.push(gap());
C.push(P('Bài chia sẻ lại một bài đã cảnh báo: ghi ở mức thấp để tránh cảnh báo trùng. Bài chia sẻ lại link website tiêu cực: tính mức lan tỏa như bài gốc.'));

C.push(SUB('Bước 3. Ra cấp độ'));
C.push(P('Ghép loại thông tin với mức lan tỏa theo bảng sau. "Nghiêm trọng" là Cấp độ 1, "Vừa" là Cấp độ 2, "Thấp" là Cấp độ 3.'));
C.push(table([4826, 1400, 1400, 1400], ['Loại thông tin tiêu cực', 'Lan tỏa cao', 'Lan tỏa trung bình', 'Lan tỏa thấp'], [
  ['Lãnh đạo: đời tư, quá trình công tác, bổ nhiệm; vi phạm pháp luật; phát ngôn bị xuyên tạc hoặc gây hiểu lầm', 'Nghiêm trọng', 'Nghiêm trọng', 'Nghiêm trọng'],
  ['Bảo mật, bí mật kinh doanh', 'Nghiêm trọng', 'Nghiêm trọng', 'Nghiêm trọng'],
  ['Hoạt động kinh doanh', 'Nghiêm trọng', 'Nghiêm trọng', 'Nghiêm trọng'],
  ['Công bố thông tin, truyền thông không đúng thực tế', 'Nghiêm trọng', 'Nghiêm trọng', 'Nghiêm trọng'],
  ['Sản phẩm, dịch vụ có lỗi mang tính hệ thống', 'Nghiêm trọng', 'Nghiêm trọng', 'Nghiêm trọng'],
  ['Nhân viên vi phạm pháp luật khi làm việc hoặc lợi dụng danh nghĩa công ty (làm giả hồ sơ, tham nhũng, lừa đảo, bị bắt, khởi tố)', 'Nghiêm trọng', 'Nghiêm trọng', 'Nghiêm trọng'],
  ['Phàn nàn về sản phẩm, dịch vụ của một hoặc một nhóm khách hàng, một hoặc vài tỉnh', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Quy định, chính sách, thủ tục', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Thái độ của nhân viên', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Tuyển dụng, đãi ngộ, điều kiện làm việc, khen thưởng, kỷ luật của nhiều nhân viên', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Đầu tư, xây dựng cơ bản, đấu thầu', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Cơ sở vật chất, công nghệ yếu kém, lạc hậu', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Nhân viên vi phạm luật giao thông, đánh nhau, phát ngôn nhạy cảm (chính trị, người khuyết tật, cộng đồng LGBT) có liên quan đến công việc', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Chia sẻ lại link website có nội dung tiêu cực', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Chia sẻ lại bài viết tiêu cực trên page, group, trang cá nhân', 'Thấp', 'Thấp', 'Thấp'],
  ['Đời tư vi phạm pháp luật của nhân viên, không liên quan công việc', 'Thấp', 'Thấp', 'Thấp'],
]));

C.push(SUB('Bước 4. Chọn luồng'));
C.push(table([2000, 7026], ['Cấp độ', 'Luồng xử lý'], [
  ['Cấp độ 1', 'Luồng C (khẩn), với mọi nhóm nội dung.'],
  ['Cấp độ 2', 'Luồng B, với mọi nhóm nội dung.'],
  ['Cấp độ 3', 'Nhóm 1 và Nhóm 4: Luồng A (MHC tự xử lý). Nhóm 2, 3, 5: Luồng B; MHC có thể đề xuất chỉ theo dõi.'],
], { boldFirst: true }));
C.push(gap());
C.push(P('MHC không tự xử lý, kể cả với Nhóm 1 và Nhóm 4, khi bài viết nêu một trường hợp khách hàng cụ thể có bằng chứng (ảnh hợp đồng, tin nhắn, ghi âm). Những bài này đưa vào Luồng B để CS kiểm tra.'));

C.push(SEC('B4. Các hình thức xử lý'));
C.push(table([1700, 3500, 3826], ['Hình thức', 'Dùng khi', 'Lưu ý'], [
  ['Theo dõi', 'Lan tỏa thấp, chưa có thảo luận, nguồn đăng ít ảnh hưởng', 'Đề xuất hình thức khác nếu tương tác tăng nhanh'],
  ['Liên hệ', 'Khách hàng có vấn đề thật và nhận diện được', 'MHC cung cấp thông tin nguồn đăng; CS của VietCredit liên hệ khách hàng'],
  ['Seeding', 'Bài có tương tác nhưng chưa lan rộng; bài hỏi đáp giải thích được', 'Dùng kịch bản đã duyệt; giãn cách giữa các bình luận; chỉ ở trang, nhóm, tài khoản công khai cho phép bình luận'],
  ['Report', 'Bài tương tác cao, cần hạn chế tiếp cận', 'Báo kết quả theo deadline đã tư vấn'],
  ['Gỡ bài', 'Thông tin sai sự thật hoặc ảnh hưởng nghiêm trọng; video livestream đã kết thúc không seeding được', 'Có chi phí ngoài gói, chỉ làm khi VietCredit duyệt. Đề xuất phải ghi chi phí, thời gian, bảo hành, tỷ lệ thành công (ví dụ MHC đã đưa: 12.000.000 đồng, 1 đến 15 ngày, bảo hành 7 ngày, tỷ lệ 85%)'],
], { boldFirst: true }));

C.push(SEC('B5. Nguyên tắc khi seeding'));
C.push(BL('Không phủ nhận lỗi mà VietCredit đã xác nhận là có thật. Khi đó bình luận ghi nhận và mời khách hàng liên hệ CSKH.'));
C.push(BL('Không công kích, chế giễu người đăng hay người bình luận.'));
C.push(BL('Không nêu thông tin cá nhân, thông tin khoản vay của khách hàng.'));
C.push(BL('Không hứa ưu đãi, miễn giảm hay kết quả mà sản phẩm không có.'));
C.push(BL('Với Nhóm 4: không gợi ý cách trốn nợ. Bình luận nói về nghĩa vụ trả nợ, ảnh hưởng tới lịch sử tín dụng và việc liên hệ VietCredit khi gặp khó khăn.'));
C.push(BL('Chỉ dùng kịch bản trong thư viện đã duyệt. Nội dung mới phải qua Luồng B.'));

C.push(SEC('B6. Lịch phối hợp và thời hạn'));
C.push(table([2900, 1900, 4226], ['Thời điểm', 'Bên thực hiện', 'Việc'], [
  ['10:30, thứ Hai đến thứ Sáu', 'MHC', 'Gửi file tổng hợp (Mẫu 01)'],
  ['Trước 14:00 cùng ngày', 'VietCredit (Marketing)', 'Trả lời phương án cho các tin Luồng B. VietCredit chỉ xem các tin cần ý kiến, không xác nhận lại toàn bộ file'],
  ['Trong ngày', 'MHC', 'Triển khai phương án đã chốt'],
  ['Trong 30 phút từ khi phát hiện', 'MHC', 'Alert OTT với sự vụ Cấp độ 1, cả ngoài giờ và cuối tuần; gửi Form đề xuất (Mẫu 02) ngay sau đó'],
  ['Trong 1 giờ từ khi có alert', 'VietCredit (Marketing)', 'Xác nhận đã nhận'],
  ['Trong 4 giờ từ khi có alert', 'VietCredit', 'Chốt phương án Cấp độ 1'],
  ['Theo mốc hai bên thống nhất', 'MHC', 'Cập nhật diễn biến sự vụ Cấp độ 1 đến khi đóng'],
], { boldFirst: true }));
C.push(gap());
C.push(P('Nếu quá 14:00 chưa có phản hồi cho một tin Luồng B, MHC nhắc đầu mối Marketing qua nhóm OTT. MHC không tự triển khai nội dung chưa được duyệt.'));

C.push(SEC('B7. Báo cáo MHC gửi VietCredit'));
C.push(table([1800, 7226], ['Báo cáo', 'Nội dung'], [
  ['Hằng ngày', 'File 10:30 (Mẫu 01)'],
  ['Hằng tuần', 'Số tin tiêu cực theo sản phẩm, kênh, nhóm, cấp độ; tình trạng xử lý; tin nổi bật; quota gói đã dùng'],
  ['Hằng tháng', 'Như báo cáo tuần, thêm: xu hướng so với tháng trước; tỷ lệ tiêu cực, trung tính, tích cực trong thảo luận; hiệu quả report, gỡ bài; các vấn đề khách hàng phàn nàn lặp lại nhiều nhất'],
  ['Theo sự vụ', 'Diễn biến và kết quả của sự vụ Cấp độ 1'],
], { boldFirst: true }));

C.push(SEC('B8. VietCredit cung cấp cho MHC'));
C.push(BL('Bộ từ khóa của từng sản phẩm, cập nhật mỗi quý.'));
C.push(BL('Thư viện kịch bản seeding đã duyệt cho từng nhóm nội dung.'));
C.push(BL('Thông báo trước khi có bảo trì hệ thống, lỗi đã biết, thay đổi phí, chính sách hoặc chiến dịch lớn, kèm câu trả lời chuẩn.'));
C.push(BL('Danh sách thành viên nhóm OTT cảnh báo khẩn.'));

C.push(SEC('B9. Mẫu biểu'));
C.push(SUB('Mẫu 01. File tổng hợp hằng ngày'));
C.push(P('Mỗi dòng là một tin, gồm các cột:'));
C.push(table([2600, 6426], ['Cột', 'Cách ghi'], [
  ['Ngày, giờ phát hiện', 'dd/mm/yyyy hh:mm'],
  ['Sản phẩm', 'VC x DMX / TikTok BNPL / sản phẩm khác'],
  ['Kênh, link', 'Facebook, TikTok, YouTube, Threads, báo điện tử; link bài viết'],
  ['Nguồn đăng', 'Tên trang, nhóm hoặc tài khoản; thông tin liên hệ nếu có'],
  ['Tóm tắt nội dung', 'Một đến hai câu'],
  ['Tương tác', 'Lượt thích, bình luận, chia sẻ, lượt xem lúc ghi nhận'],
  ['Nhóm, cấp độ', 'Nhóm 1 đến 5; Cấp độ 1, 2, 3'],
  ['Luồng', 'A (đã tự xử lý) / B (cần ý kiến VietCredit)'],
  ['Đề xuất của MHC', 'Hình thức xử lý và hướng nội dung'],
  ['Phương án VietCredit chốt', 'Marketing ghi'],
  ['Trạng thái, kết quả', 'Đang xử lý / Đã xử lý / Theo dõi; kết quả'],
], { boldFirst: true }));
C.push(SUB('Mẫu 02. Form đề xuất sự vụ lớn'));
C.push(table([2300, 6726], ['Mục', 'Nội dung'], [
  ['0. Tên sự vụ', 'Tên ngắn gọn, sản phẩm liên quan, cấp độ'],
  ['1. Tổng quan', 'Tổng số bài đăng; số bài trên từng nền tảng'],
  ['2. Nguồn đăng', 'Tên, link; số điện thoại, email, kênh xã hội khác (nếu có)'],
  ['3. Nội dung', ['• Nội dung sự vụ', '• Tương tác hiện tại, thời điểm thảo luận gần nhất', '• Tỷ lệ bình luận tiêu cực, trung tính, tích cực và ý chính từng nhóm', '• Các bài nổi bật (nếu có)']],
  ['4. Đánh giá', 'Mức lan tỏa, mức ảnh hưởng, lý do xếp cấp độ'],
  ['5. Đề xuất', ['• Phương án: theo dõi, seeding (số bình luận, giãn cách, hướng nội dung), report, gỡ bài', '• Với report, gỡ bài: chi phí, thời gian, bảo hành, tỷ lệ thành công', '• Việc VietCredit nên tránh trong lúc chờ phương án']],
], { boldFirst: true }));
C.push(SUB('Mẫu 03. Báo cáo tuần, tháng'));
C.push(P('Theo nội dung ở mục B7. Hai bên thống nhất mẫu trình bày trong tuần đầu triển khai.'));

// ===== PHẦN C =====
PART('PHẦN C', 'QUY TRÌNH NỘI BỘ VIETCREDIT',
  'Phần này dùng nội bộ, không gửi MHC. Sau giai đoạn thí điểm, Marketing hoàn thiện phần này để ban hành thành quy trình nội bộ.');

C.push(SEC('C1. Vai trò các đơn vị'));
C.push(table([1700, 7326], ['Đơn vị', 'Trách nhiệm'], [
  ['Marketing', ['• Đầu mối duy nhất làm việc với MHC.', '• Đọc file 10:30, chuyển tin Luồng B cho đơn vị liên quan, chốt và gửi phương án cho MHC.', '• Duyệt và quản lý thư viện kịch bản.', '• Báo BĐH sự vụ Cấp độ 1; báo cáo tháng.', '• Theo dõi quota gói và chi phí ngoài gói.']],
  ['PD/PM sản phẩm', ['• Cho ý kiến về tính năng, lỗi, phí, chính sách, điều kiện sản phẩm.', '• Báo trước cho Marketing các thay đổi có thể gây phản ứng.', '• Lên kế hoạch khắc phục các vấn đề lặp lại.']],
  ['CS', ['• Kiểm tra thông tin khi bài viết nêu khách hàng cụ thể.', '• Liên hệ khách hàng khi phương án là "Liên hệ".', '• Cung cấp câu trả lời chuẩn cho thắc mắc thường gặp.', '• Báo Marketing khi tổng đài quá tải hoặc nhận nhiều khiếu nại cùng một vấn đề.']],
  ['Pháp chế', ['• Cho ý kiến khi bài viết có dấu hiệu vu khống, xúc phạm, lộ thông tin khách hàng.', '• Soạn yêu cầu gỡ bài chính thức gửi nền tảng hoặc cơ quan báo chí khi cần.']],
  ['BĐH', ['• Chỉ đạo xử lý sự vụ Cấp độ 1.', '• Duyệt phát ngôn chính thức và chi phí vượt hạn mức.']],
], { boldFirst: true }));

C.push(SEC('C2. Xử lý tin Luồng B'));
C.push(SUB('Hỏi ai'));
C.push(table([4026, 5000], ['Nội dung tin', 'Đơn vị cho ý kiến'], [
  ['Lỗi tính năng, ứng dụng, không thao tác được', 'PD/PM sản phẩm'],
  ['Phí, lãi, hạn mức, điều kiện, chính sách sản phẩm', 'PD/PM sản phẩm'],
  ['Khách hàng cụ thể phàn nàn, có bằng chứng', 'CS (kiểm tra và liên hệ khách hàng); PD/PM nếu do lỗi sản phẩm'],
  ['Thái độ nhân viên, tư vấn sai', 'CS và đơn vị quản lý nhân viên đó'],
  ['Doanh nghiệp, tình hình kinh doanh, lãnh đạo', 'Marketing trình BĐH'],
  ['Vu khống, xúc phạm, lộ thông tin khách hàng', 'Pháp chế'],
], { boldFirst: true }));
C.push(SUB('Các bước và thời hạn'));
C.push(table([2300, 2000, 4726], ['Thời hạn', 'Người làm', 'Việc'], [
  ['10:30', 'MHC', 'Gửi file tổng hợp'],
  ['Trước 11:30', 'Marketing', 'Lọc tin Luồng B, gửi đơn vị theo bảng "Hỏi ai", kèm link bài và đề xuất của MHC'],
  ['Trước 13:30', 'PD/PM, CS, Pháp chế', 'Trả lời: thông tin đúng hay sai, hướng trả lời, có cần liên hệ khách hàng không'],
  ['Trước 14:00', 'Marketing', 'Chốt kịch bản, gửi MHC; lưu kịch bản mới vào thư viện'],
], { boldFirst: true }));
C.push(gap());
C.push(P('Nếu đơn vị không trả lời trước 13:30, Marketing gửi MHC kịch bản chung (ghi nhận phản ánh, mời khách hàng liên hệ CSKH qua hotline 1900 6515) và báo trưởng đơn vị đó. Kịch bản chi tiết gửi bổ sung khi có ý kiến.'));

C.push(SEC('C3. Xử lý sự vụ Cấp độ 1 (Luồng C)'));
C.push(SUB('Nhóm xử lý'));
C.push(P('Gồm: Trưởng Marketing (điều phối), PD/PM của sản phẩm liên quan, Trưởng CS, Pháp chế, và thành viên BĐH phụ trách. Danh sách và số điện thoại lập sẵn, cập nhật mỗi quý.'));
C.push(SUB('Các bước'));
NUMS([
  'Trong 1 giờ từ khi có alert: Marketing xác nhận với MHC, báo BĐH, gọi nhóm xử lý.',
  'Các đơn vị kiểm tra sự việc: đúng hay sai, bao nhiêu khách hàng bị ảnh hưởng, đã có ai khiếu nại qua tổng đài chưa.',
  'Trong 4 giờ: nhóm xử lý đề xuất phương án; BĐH duyệt phương án, nội dung phát ngôn và chi phí.',
  'Marketing gửi phương án cho MHC; CS dùng cùng nội dung trả lời khách hàng gọi tổng đài.',
  'Nhóm xử lý theo dõi theo mốc MHC cập nhật cho đến khi đóng sự vụ.',
  'Trong 5 ngày làm việc sau khi đóng: Marketing gửi BĐH báo cáo nguyên nhân, cách xử lý, chi phí và việc cần sửa.',
]);
C.push(SUB('Nguyên tắc'));
C.push(BL('Chỉ người được BĐH giao mới phát ngôn với báo chí và trên kênh chính thức.'));
C.push(BL('Nhân viên không tự bình luận, giải thích về sự vụ trên tài khoản cá nhân.'));
C.push(BL('Không xóa bình luận hàng loạt trên fanpage của VietCredit khi chưa có ý kiến nhóm xử lý.'));
C.push(BL('Trong sự vụ lớn, không dùng seeding thay cho phản hồi chính thức.'));

C.push(SEC('C4. Duyệt chi phí ngoài gói'));
C.push(table([3500, 5526], ['Khoản', 'Người duyệt'], [
  ['Seeding, theo dõi trong quota gói', 'Không cần duyệt riêng'],
  ['Report, gỡ bài đến hạn mức [số tiền, BĐH chốt] mỗi lần', 'Trưởng Marketing'],
  ['Vượt hạn mức trên, hoặc sự vụ Cấp độ 1', 'BĐH'],
], { boldFirst: true }));

C.push(SEC('C5. Phòng ngừa chủ động'));
NUMS([
  '**Bộ từ khóa.** PD/PM cung cấp tên sản phẩm, cách khách hàng hay gọi, tên đối tác (Điện Máy Xanh, TikTok Shop), các lỗi và thắc mắc thường gặp. Marketing gửi MHC và rà lại mỗi quý.',
  '**Thư viện kịch bản.** Marketing duyệt câu trả lời mẫu cho từng nhóm nội dung; PD/PM và CS xác nhận nội dung đúng.',
  '**Báo trước.** Trước khi bảo trì hệ thống, đổi phí, chính sách, chạy chiến dịch lớn, hoặc khi có lỗi đã biết, PD/PM báo Marketing ít nhất 2 ngày làm việc (trường hợp lỗi đột xuất thì báo ngay), kèm câu trả lời chuẩn.',
  '**Theo dõi khi ra mắt.** Trong 2 tuần đầu ra mắt sản phẩm, tính năng hoặc chiến dịch, Marketing yêu cầu MHC báo cáo nhanh hằng ngày cho sản phẩm đó.',
  '**Phối hợp với CS.** Khi tin "hotline không gọi được" tăng đột biến, Marketing báo CS kiểm tra tổng đài. Khi CS nhận nhiều khiếu nại cùng một vấn đề, CS báo Marketing để chuẩn bị kịch bản trước.',
  '**Sửa tận gốc.** Họp tháng giữa Marketing, PD/PM, CS xem 5 vấn đề bị phàn nàn nhiều nhất trong báo cáo tháng của MHC; PD/PM đưa ra kế hoạch khắc phục và thời hạn.',
]);

C.push(SEC('C6. Báo cáo nội bộ và chỉ số theo dõi'));
C.push(P('Marketing gửi BĐH báo cáo tháng, gồm báo cáo của MHC và các chỉ số sau:'));
C.push(BL('Số tin tiêu cực mới theo sản phẩm, kênh, nhóm, cấp độ; so với tháng trước.'));
C.push(BL('Tỷ lệ tiêu cực trong tổng thảo luận về từng sản phẩm.'));
C.push(BL('Thời gian từ khi phát hiện đến khi có phương án; tỷ lệ tin Luồng B được trả lời trước 14:00.'));
C.push(BL('Tỷ lệ report, gỡ bài thành công; số bài xuất hiện lại trong thời gian bảo hành.'));
C.push(BL('Quota gói đã dùng; chi phí ngoài gói.'));
C.push(BL('Các vấn đề lặp lại và tiến độ khắc phục của PD/PM.'));

// ===== PHẦN D =====
PART('PHẦN D', 'TRIỂN KHAI VÀ NỘI DUNG CẦN BAN ĐIỀU HÀNH QUYẾT ĐỊNH', null);

C.push(SEC('D1. Kế hoạch triển khai'));
C.push(table([2000, 4826, 2200], ['Thời gian', 'Việc', 'Thực hiện'], [
  ['Tuần 1', 'BĐH duyệt quy trình. Gửi Phần B cho MHC để thống nhất. Lập nhóm OTT và danh sách nhóm xử lý Cấp độ 1.', 'Marketing'],
  ['Tuần 1 đến 2', 'Chốt bộ từ khóa và thư viện kịch bản cho Thẻ tín dụng VC x DMX và TikTok BNPL.', 'PD/PM, CS, Marketing, MHC'],
  ['Tuần 3 đến 6', 'Chạy thí điểm. Marketing và MHC họp 15 phút mỗi tuần để chỉnh cách phân loại.', 'Marketing, MHC'],
  ['Cuối tuần 6', 'Đánh giá thí điểm theo chỉ số ở mục C6; hoàn thiện Phần C thành quy trình nội bộ; đề xuất sản phẩm áp dụng tiếp theo.', 'Marketing'],
], { boldFirst: true }));

C.push(SEC('D2. Nội dung cần BĐH quyết định'));
NUMS([
  'Phạm vi giai đoạn 1: Thẻ tín dụng VC x DMX và TikTok BNPL; các dự án khác chỉ áp dụng luồng khẩn cho Cấp độ 1.',
  'Cho phép MHC tự xử lý Nhóm 1 và Nhóm 4 ở Cấp độ 3 theo kịch bản đã duyệt.',
  'Hạn mức chi phí report, gỡ bài mà Trưởng Marketing được duyệt mỗi lần.',
  'Người được phát ngôn chính thức và thành viên BĐH phụ trách khi có sự vụ Cấp độ 1.',
  'Các thời hạn: VietCredit trả lời trước 14:00; Cấp độ 1 cảnh báo trong 30 phút, có phương án trong 4 giờ.',
]);

C.push(SEC('D3. Điểm đã điều chỉnh so với tài liệu của MHC'));
C.push(BL('Giờ gửi file: tài liệu MHC ghi cả 10h và 10:30; quy trình dùng 10:30.'));
C.push(BL('MHC đề xuất VietCredit xác nhận toàn bộ file mỗi ngày. Quy trình đổi thành: MHC tự phân loại, VietCredit chỉ trả lời các tin Luồng B.'));
C.push(BL('Nhân viên vi phạm pháp luật: hình "Cấp độ khủng hoảng" của MHC xếp ở Cấp độ 3, bảng chi tiết xếp trường hợp vi phạm khi làm việc là Nghiêm trọng. Quy trình theo bảng chi tiết.'));

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
    { reference: 'n1', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
    { reference: 'n2', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
    { reference: 'n3', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
    { reference: 'n4', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
    { reference: 'n5', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
    { reference: 'n6', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
    { reference: 'n7', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
    { reference: 'n8', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
    { reference: 'n9', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
    { reference: 'n10', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1200, left: 1440, right: 1440, header: 560 } } },
    headers: { default: header }, footers: { default: footer },
    children: C,
  }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log('wrote', OUT); });
