const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer,
  ImageRun, AlignmentType, WidthType, BorderStyle, VerticalAlign, HeightRule,
  ShadingType, PageNumber, PageOrientation, LevelFormat, HeadingLevel, PageBreak,
} = require('docx');

const OUT = process.argv[2];
const RED = 'A10719', DARK = '404041', GREY = '5A5A5A', ALT = 'EEEEEE';
const FONT = 'Arial';
const W = 15138; // A4 ngang, lề 850
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
    width: { size: W, type: WidthType.DXA }, columnWidths: [1900, 13238],
    borders: { top: NB, bottom: NB, left: NB, right: NB, insideHorizontal: NB, insideVertical: NB },
    rows: [new TableRow({ height: { value: 360, rule: HeightRule.EXACT }, children: [
      new TableCell({ width: { size: 1900, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
        margins: { top: 0, bottom: 0, left: 0, right: 160 }, borders: { top: NB, bottom: NB, left: NB, right: NB },
        children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [
          new ImageRun({ data: logo, type: 'png', transformation: { width: 106, height: 24 } })] })] }),
      new TableCell({ width: { size: 13238, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
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
const SHEET_URL = 'https://docs.google.com/spreadsheets/d/1xp3EIZZXDeLg45_3XUYrvG8A6sjjVtlwraIty1DJtLE';

// Cột theo vai trò (A4 ngang). Tổng = W
const ROLES = ['step', 'sl', 'xl', 'mkt', 'pm'];
const HEAD = ['Bước / thời hạn', 'MHC - Hệ thống Social listening', 'MHC - Dịch vụ xử lý tin tiêu cực', 'VC - Marketing', 'VC - PM/PD'];
const COLS = [1900, 2900, 3700, 3400, 3238];
const FS = 16; // 8pt

// Một dòng trong ô: "• " gạch đầu dòng, "→ "/"← "/"↓ " mũi tên chuyển việc (đỏ, đậm),
// "~ " ghi chú nghiêng xám, "**x**" in đậm
function line(t) {
  if (t.startsWith('• ')) return new Paragraph({ numbering: { reference: 'cbul', level: 0 }, spacing: { after: 30, line: 252 },
    children: runs(t.slice(2), { size: FS }) });
  if (/^[→←↓]/.test(t)) return new Paragraph({ spacing: { before: 30, after: 30, line: 252 },
    children: [new TextRun({ text: t, bold: true, font: FONT, size: FS, color: RED })] });
  if (t.startsWith('~ ')) return new Paragraph({ spacing: { before: 40, after: 30, line: 252 },
    children: runs(t.slice(2), { size: FS - 1, italics: true, color: GREY }) });
  return new Paragraph({ spacing: { after: 30, line: 252 }, children: runs(t, { size: FS }) });
}
const cellBase = (w, fill, children) => new TableCell({
  width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.TOP,
  margins: { top: 60, bottom: 40, left: 80, right: 80 },
  borders: { top: TB, bottom: TB, left: TB, right: TB },
  shading: { type: ShadingType.CLEAR, color: 'auto', fill }, children,
});
const headRow = new TableRow({ tableHeader: true, cantSplit: true, children: HEAD.map((h, i) =>
  new TableCell({ width: { size: COLS[i], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
    margins: { top: 70, bottom: 50, left: 80, right: 80 }, borders: { top: TB, bottom: TB, left: TB, right: TB },
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: RED },
    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, bold: true, font: FONT, size: 18, color: 'FFFFFF' })] })] })) });

const R = [];
const G = text => R.push(new TableRow({ cantSplit: true, children: [new TableCell({
  columnSpan: COLS.length, width: { size: W, type: WidthType.DXA },
  margins: { top: 70, bottom: 70, left: 100, right: 100 }, borders: { top: TB, bottom: TB, left: TB, right: TB },
  shading: { type: ShadingType.CLEAR, color: 'auto', fill: 'D8D9E0' },
  children: [new Paragraph({ keepNext: true, children: [new TextRun({ text, bold: true, font: FONT, size: 19, color: RED })] })] })] }));
// S(tên bước, [thời hạn, ghi chú...], { mhc: [...], mkt: [...], ... })
const S = (name, meta, cells) => {
  const len = JSON.stringify(cells).length;
  R.push(new TableRow({ cantSplit: len < 1300, children: ROLES.map((role, i) => {
    if (role === 'step') return cellBase(COLS[i], 'FFFFFF', [
      new Paragraph({ spacing: { after: 50 }, children: [new TextRun({ text: name, bold: true, font: FONT, size: 18, color: DARK })] }),
      ...meta.map(line)]);
    const c = cells[role];
    return c ? cellBase(COLS[i], 'FFFFFF', c.map(line)) : cellBase(COLS[i], ALT, [new Paragraph({ children: [] })]);
  }) }));
};

// ===== A. THIẾT LẬP =====
G('A. THIẾT LẬP BAN ĐẦU (làm một lần trước khi áp dụng, rà lại mỗi quý)');
S('0. Chuẩn bị', ['Tuần đầu triển khai'], {
  sl: [
    '← nhận bộ từ khóa từ Marketing',
    '• Cấu hình từ khóa cho Thẻ tín dụng VC x DMX, TikTok BNPL trên hệ thống',
    '• Cài kênh nhận alert: nhóm chat, email của Marketing và đội xử lý',
    '• Rà lại từ khóa mỗi quý',
  ],
  xl: [
    '• Thêm cột vào link theo dõi: Tương tác, Ý kiến VC, Người duyệt, Nghiệm thu, Yêu cầu bổ sung, Kiểm tra sau 48 giờ',
    '• Soạn kịch bản mẫu Nhóm 1, 4',
    '→ gửi Marketing',
    '• Xác nhận: cập nhật link lúc 13:30; thời hạn bổ sung (bước 9); luồng khẩn (K1 đến K4)',
  ],
  mkt: [
    '• Chốt đầu mối VC: [tên, số điện thoại]; đầu mối MHC: Hằng, Lam',
    '• Lập nhóm chat [Zalo/Telegram] gồm đầu mối MHC, Marketing, PM/PD',
    '• Link theo dõi chung: ' + SHEET_URL,
    '← nhận bộ từ khóa từ PM/PD', '→ gửi MHC Social listening cấu hình',
    '← nhận kịch bản mẫu từ MHC xử lý tin',
    '→ gửi PM/PD xác nhận nội dung, sau đó duyệt',
    '→ trình BĐH duyệt: phạm vi giai đoạn 1; cho MHC tự seeding Nhóm 1, 4; hạn mức chi phí Trưởng Marketing được duyệt; người phát ngôn khi có sự vụ khẩn',
  ],
  pm: [
    '• Cung cấp bộ từ khóa: tên sản phẩm, cách khách hàng hay gọi, tên đối tác (Điện Máy Xanh, TikTok Shop)',
    '→ gửi Marketing',
    '← nhận kịch bản mẫu',
    '• Tự lấy xác nhận nội dung của CS (Nhóm 1: hotline, kênh CSKH) và Thu hồi nợ/Collex (Nhóm 4: bùng nợ, nhắc nợ, đến nhà)',
    '→ gửi Marketing',
    '• Từ nay: báo Marketing trước ít nhất 2 ngày làm việc khi đổi phí, chính sách, bảo trì hệ thống, chạy chiến dịch lớn',
  ],
});

// ===== B. HẰNG NGÀY =====
G('B. XỬ LÝ HẰNG NGÀY (thứ Hai đến thứ Sáu)');
S('1. Ghi nhận, cảnh báo', ['Liên tục', 'Khung tổng hợp: 9h00 hôm trước đến 9h00 hôm nay'], {
  sl: [
    '• Quét liên tục các kênh theo bộ từ khóa',
    '• Ghi nhận tin tiêu cực vào hệ thống Crisis',
    '• Gửi alert tự động',
    '→ Marketing',
    '→ đội xử lý tin tiêu cực',
  ],
  xl: [
    '← nhận alert',
    '• Đánh giá nhanh: có phải sự vụ khẩn không (lãnh đạo, tình hình kinh doanh, bảo mật, lỗi sản phẩm hàng loạt, bài trên 100 tương tác, lên báo điện tử)',
    '↓ Khẩn: chuyển ngay bước K1',
    '↓ Không khẩn: tổng hợp ở bước 2',
  ],
  mkt: ['← nhận alert', '• Theo dõi, chưa cần xử lý cho đến bước 3'],
});
S('2. Phân nhóm, đề xuất', ['Buổi sáng', 'Cập nhật link lúc 13:30', '~ Tình trạng trên link: Nhóm 1, 4 "Đang triển khai"; nhóm khác "Chờ VC duyệt"'], {
  xl: [
    '**Phân nhóm:**',
    '• Nhóm 1: hotline không liên hệ được, "VietCredit lừa đảo". Seeding luôn',
    '• Nhóm 2: lỗi tính năng, không thao tác được. Chờ duyệt',
    '• Nhóm 3: doanh nghiệp, tình hình kinh doanh, lãnh đạo. Chờ duyệt',
    '• Nhóm 4: hỏi bùng nợ, đòi nợ đến nhà, nên vay bên nào. Seeding luôn trong quota, không cần 100%',
    '• Khác: mạo danh, bên thứ ba dùng tên VietCredit... Chờ duyệt',
    '**Nhóm 1, 4 vẫn chờ duyệt nếu:**',
    '• khách hàng cụ thể tố VietCredit làm sai, có chi tiết',
    '• bài trên 100 tương tác',
    '• bên thứ ba dùng tên VietCredit quảng cáo dịch vụ',
    '**Ghi mỗi case một dòng trên link:**',
    '• ngày (dd/mm/yyyy), nền tảng, link, nhóm, tương tác, đánh giá',
    '• phương án: theo dõi / liên hệ người đăng / seeding / report / gỡ bài (kèm chi phí, thời hạn, bảo hành)',
    '↓ Nhóm 1, 4: seeding luôn, sang bước 6',
    '→ Nhóm khác: nhắn nhóm chat báo Marketing',
  ],
});
S('3. Review', ['13:30 đến 14:30'], {
  mkt: [
    '← nhận link từ MHC',
    '**Kiểm tra từng case:**',
    '• MHC phân nhóm đúng chưa?',
    '• Phương án hợp lý chưa? (quảng cáo mạo danh thì nên report)',
    '• Bình luận đúng quy định ở bước 6 chưa?',
    '• Có chi phí ngoài gói không?',
    '↓ Tự quyết được (thông tin chung, không cần tra cứu): sang bước 5',
    '→ Cần thông tin sản phẩm, khách hàng, khoản nợ, pháp lý: gửi PM/PD kèm link bài và đề xuất của MHC',
    '→ Nhóm 3 (doanh nghiệp, lãnh đạo): trình BĐH',
  ],
});
S('4. PM tra thông tin, cho ý kiến', ['Trước 16:00', '~ Quá 16:00 chưa có ý kiến: Marketing duyệt phương án an toàn (theo dõi, hoặc mời khách hàng liên hệ CSKH) và nhắc PM/PD'], {
  pm: [
    '← nhận case từ Marketing',
    '**Tự tra thông tin hoặc liên hệ đơn vị liên quan:**',
    '• lỗi tính năng, phí, chính sách: tự kiểm tra',
    '• khách hàng cụ thể phàn nàn dịch vụ: CS tra cứu',
    '• nhắc nợ, đòi nợ, gọi người thân: Thu hồi nợ/Collex kiểm tra lịch sử nhắc nợ, cuộc gọi',
    '• mạo danh, vu khống: Pháp chế đánh giá',
    '**Trả lời Marketing:**',
    '• thông tin đúng hay sai',
    '• hướng trả lời',
    '• có cần liên hệ khách hàng không, đơn vị nào liên hệ',
    '→ Marketing',
  ],
});
S('5. Duyệt', ['Trước 16:30', '~ Tình trạng trên link: "Đang triển khai" hoặc "Không xử lý"'], {
  mkt: [
    '← nhận ý kiến PM/PD',
    '**Ghi cột "Ý kiến VC" trên link:**',
    '• Đồng ý',
    '• Đồng ý, sửa nội dung thành...',
    '• Đổi phương án sang...',
    '• Không xử lý',
    '• Ghi tên người duyệt',
    '**Chi phí ngoài gói (report, gỡ bài):**',
    '• đến [hạn mức] mỗi lần: Trưởng Marketing duyệt',
    '• vượt hạn mức: trình BĐH',
    '→ nhắn nhóm chat báo MHC xử lý tin',
  ],
});
S('6. Triển khai', ['Trong ngày, chậm nhất 9h00 hôm sau', '~ Tình trạng trên link: "Chờ nghiệm thu"'], {
  xl: [
    '← nhận phương án đã duyệt',
    '**Seeding:** ghi vào sheet "Seeding"',
    '• nội dung từng bình luận',
    '• link trực tiếp hoặc ảnh chụp bình luận',
    '• thời điểm đăng',
    '**Report, gỡ bài:** ghi ngày gửi, hạn dự kiến, kết quả',
    '**Quy định bình luận:**',
    '• không đăng phần nhãn trong ngoặc vuông',
    '• viết đúng "VietCredit"',
    '• không viết câu ngầm nhận VietCredit làm sai',
    '• không lên lớp, dọa người đăng',
    '• không gợi ý trốn nợ, không hứa ưu đãi chưa được xác nhận',
    '• mỗi bình luận khác câu chữ, giãn cách, khác tài khoản',
    '**Liên hệ người đăng:** không tự liên hệ nhân danh VietCredit',
    '→ gửi thông tin người đăng cho Marketing',
    '→ báo Marketing đã xong',
  ],
  mkt: ['← nhận thông tin người đăng (nếu phương án là liên hệ)', '→ chuyển PM/PD'],
  pm: [
    '← nhận thông tin người đăng',
    '• Yêu cầu CS (phàn nàn dịch vụ) hoặc Thu hồi nợ/Collex (vấn đề khoản nợ) liên hệ khách hàng',
    '• Nhận lại kết quả',
    '→ báo Marketing',
  ],
});
S('7. Kiểm tra sau 48 giờ', ['48 giờ sau khi đăng'], {
  xl: [
    '• Mở lại từng bình luận đã đăng',
    '• Ghi cột "Kiểm tra sau 48 giờ": Còn / Bị ẩn / Bị xóa',
    '• Bị ẩn, bị xóa: đăng lại, cập nhật bằng chứng, không tính quota',
  ],
});
S('8. Nghiệm thu', ['Case đã duyệt: trong 1 ngày làm việc', 'Nhóm 1, 4: thứ Sáu hằng tuần', '~ Tình trạng trên link: "Hoàn thành" hoặc "Cần bổ sung"'], {
  mkt: [
    '← mở bằng chứng trên link',
    '**Đạt khi:**',
    '• seeding: đúng bài, đúng nội dung duyệt, đủ số lượng, đúng quy định, còn hiển thị sau 48 giờ',
    '• report, gỡ bài: xong đúng hạn, có bằng chứng, không đăng lại trong thời gian bảo hành',
    '• theo dõi: cập nhật đúng hẹn',
    '**Nhóm 1, 4:** kiểm tra ngẫu nhiên ít nhất 20% case trong tuần; hơn 1/5 không đạt thì tuần sau kiểm tra 100%',
    '**Ghi cột "Nghiệm thu":**',
    '• Đạt / Đạt có góp ý',
    '• Cần bổ sung: ghi rõ việc cần làm, hạn',
    '→ Cần bổ sung: báo MHC, sang bước 9',
  ],
});
S('9. Bổ sung, làm lại', ['Xác nhận: trong 2 giờ làm việc', 'Hoàn thành: trong 1 ngày làm việc'], {
  xl: [
    '← nhận yêu cầu trên link',
    '• Sửa, xóa bình luận sai',
    '• Thêm bình luận khi thảo luận tăng',
    '• Đăng lại bình luận bị ẩn, xóa',
    '• Đổi phương án (ví dụ sang report)',
    '• Cập nhật bằng chứng mới',
    '~ Lỗi do MHC hoặc trong thời gian bảo hành: không tính thêm phí, quota',
    '→ Marketing nghiệm thu lại',
  ],
  mkt: ['← nghiệm thu lại như bước 8'],
});
S('10. Xử lý tận gốc', ['Khi review hoặc nghiệm thu thấy vấn đề có thật'], {
  sl: ['• Khi Marketing yêu cầu: trích số liệu thảo luận theo vấn đề (số bài, xu hướng theo ngày) để PM/PD đánh giá mức độ', '→ Marketing'],
  mkt: [
    '• Gom các vấn đề có thật (nhiều tin hotline không gọi được, lỗi lặp lại, phản ánh nhắc nợ sai, mạo danh)',
    '→ chuyển PM/PD',
    '• Cân nhắc cảnh báo khách hàng trên kênh chính thức khi có mạo danh',
  ],
  pm: [
    '← nhận vấn đề',
    '**Tự xử lý hoặc yêu cầu đơn vị liên quan:**',
    '• lỗi tính năng lặp lại: sửa, báo thời hạn',
    '• tổng đài quá tải: CS kiểm tra năng lực',
    '• nhắc nợ bị phản ánh: Thu hồi nợ/Collex kiểm tra cuộc gọi, nhân sự, đối tác nhắc nợ',
    '• mạo danh: Pháp chế xử lý',
    '→ báo Marketing khi xong',
  ],
});

// ===== C. KHẨN =====
G('C. SỰ VỤ KHẨN (lãnh đạo, tình hình kinh doanh, bảo mật, lỗi sản phẩm hàng loạt, bài trên 100 tương tác, lên báo điện tử)');
S('K1. Báo động', ['Trong 30 phút từ khi phát hiện, kể cả ngoài giờ, cuối tuần'], {
  sl: ['• Alert ngay khi phát hiện', '• Thống kê số bài, nền tảng, tương tác, nguồn đăng', '→ đội xử lý tin tiêu cực'],
  xl: [
    '• Gọi điện đầu mối Marketing, nhắn nhóm chat',
    '**Gửi form sự vụ:**',
    '• tên sự vụ; số bài, nền tảng',
    '• nguồn đăng, thông tin liên hệ',
    '• nội dung, bình luận nổi bật, tỷ lệ tiêu cực',
    '• tương tác hiện tại; đánh giá lan tỏa, ảnh hưởng',
    '• đề xuất phương án, chi phí',
    '• việc VietCredit nên tránh trong lúc chờ',
    '→ Marketing',
  ],
});
S('K2. Tiếp nhận', ['Trong 1 giờ'], {
  mkt: ['← nhận form', '• Xác nhận với MHC', '→ báo BĐH, đề nghị phân công thành viên BĐH phụ trách', '→ báo PM/PD kiểm tra sự việc'],
  pm: ['← nhận từ Marketing', '• Gọi các đơn vị liên quan: CS, Thu hồi nợ/Collex, Pháp chế'],
});
S('K3. Chốt phương án', ['Trong 4 giờ', '~ Chỉ người được BĐH giao mới phát ngôn. Nhân viên không tự bình luận trên tài khoản cá nhân. Không xóa bình luận hàng loạt trên fanpage khi chưa có ý kiến BĐH'], {
  sl: ['• Theo dõi sát diễn biến: số bài mới, tương tác, kênh lan sang', '→ đội xử lý tin tiêu cực, Marketing'],
  xl: ['← nhận phương án đã duyệt', '• Triển khai', '• Cập nhật diễn biến theo mốc hẹn'],
  mkt: [
    '← nhận thông tin từ PM/PD',
    '• Soạn phương án, nội dung phát ngôn',
    '→ trình BĐH duyệt phương án, phát ngôn, chi phí, người phát ngôn',
    '→ gửi MHC xử lý tin phương án đã duyệt',
    '→ gửi PM/PD nội dung để CS dùng trả lời khách hàng gọi tổng đài',
  ],
  pm: [
    '**Tổng hợp từ các đơn vị:**',
    '• sự việc đúng hay sai, nguyên nhân',
    '• số khách hàng bị ảnh hưởng',
    '• số khiếu nại tổng đài đã nhận (CS)',
    '• rủi ro pháp lý, góp ý phát ngôn (Pháp chế)',
    '→ Marketing',
    '← nhận nội dung đã duyệt, chuyển CS',
  ],
});
S('K4. Theo dõi, đóng', ['Đến khi ổn định', 'Báo cáo: trong 5 ngày làm việc sau khi đóng'], {
  sl: ['• Theo dõi đến khi thảo luận lắng xuống', '• Cung cấp số liệu diễn biến cho báo cáo sự vụ', '→ Marketing'],
  xl: ['• Cập nhật kết quả xử lý đến khi đóng sự vụ', '→ Marketing'],
  mkt: ['**Báo cáo sự vụ:**', '• nguyên nhân', '• cách xử lý', '• chi phí', '• việc cần sửa', '→ BĐH'],
  pm: ['• Cung cấp nguyên nhân, việc cần sửa phía sản phẩm', '→ Marketing'],
});

// ===== D. ĐỊNH KỲ =====
G('D. ĐỊNH KỲ');
S('D1. Báo cáo tuần', ['Thứ Hai'], {
  sl: ['**Báo cáo Social listening tuần:**', '• Tổng thảo luận về VietCredit, từng sản phẩm', '• Tỷ lệ tích cực, trung tính, tiêu cực', '• Chủ đề tiêu cực nổi bật', '→ Marketing'],
  xl: ['**Báo cáo xử lý tuần:**', '• Số case theo sản phẩm, nền tảng, nhóm', '• Tình trạng xử lý', '• Số bình luận, tỷ lệ còn hiển thị', '• Quota đã dùng', '→ Marketing'],
  mkt: ['← nhận, xem', '• Nhắc các case quá hạn'],
});
S('D2. Báo cáo, đối chiếu tháng', ['Ngày làm việc thứ 3 của tháng'], {
  sl: ['**Báo cáo Social listening tháng:**', '• Xu hướng thảo luận so với tháng trước', '• Tỷ lệ tiêu cực theo sản phẩm', '• 5 vấn đề bị phàn nàn nhiều nhất', '• So sánh với đối thủ', '→ Marketing'],
  xl: ['**Báo cáo xử lý tháng:**', '• Như báo cáo tuần, thêm kết quả report, gỡ bài', '• Bảng khối lượng đã làm (case, bình luận, report, gỡ bài)', '→ Marketing'],
  mkt: [
    '← nhận',
    '• Đối chiếu khối lượng với kết quả "Đạt" trên link trước khi xác nhận thanh toán',
    '**Đánh giá MHC:**',
    '• tỷ lệ đạt ngay lần đầu; tỷ lệ đúng hạn',
    '• số yêu cầu bổ sung; số lần phân nhóm sai',
    '→ gửi nhận xét cho MHC (cả hai mảng)',
  ],
});
S('D3. Họp tháng, báo cáo BĐH', ['Tuần đầu tháng'], {
  mkt: [
    '• Chủ trì họp với PM/PD, trình bày 5 vấn đề bị phàn nàn nhiều nhất',
    '**Báo cáo tháng:**',
    '• số case, tỷ lệ tiêu cực',
    '• tỷ lệ duyệt đúng hạn',
    '• nghiệm thu, đánh giá MHC',
    '• chi phí ngoài gói',
    '• tiến độ khắc phục',
    '→ BĐH',
  ],
  pm: ['• Dự họp', '• Thống nhất việc khắc phục, thời hạn (kể cả việc của CS, Thu hồi nợ/Collex)', '→ Marketing'],
});

// ===== Tiêu đề + chú thích + bảng =====
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 }, children: [
  new TextRun({ text: 'QUY TRÌNH XỬ LÝ THÔNG TIN TIÊU CỰC TRÊN MẠNG XÃ HỘI', font: FONT, size: 28, bold: true, color: RED })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [
  new TextRun({ text: 'Phối hợp VietCredit và Màn Hình Cộng (MHC). Áp dụng: Thẻ tín dụng VC x DMX, TikTok BNPL. Kênh: Facebook, TikTok, YouTube, Threads, báo điện tử', font: FONT, size: 18, color: GREY })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 140 }, children: [
  new TextRun({ text: 'Cách đọc: ', font: FONT, size: 17, bold: true, color: GREY }),
  new TextRun({ text: '→ chuyển việc   ← nhận việc   ↓ sang bước khác   Ô xám: không có việc.   VC PM/PD là đầu mối nội bộ, tự liên hệ CS, Thu hồi nợ/Collex, Pháp chế để tra thông tin và yêu cầu xử lý', font: FONT, size: 17, color: GREY })] }));
C.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: COLS, rows: [headRow, ...R] }));

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
    properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 1250, bottom: 900, left: 850, right: 850, header: 450, footer: 400 } } },
    headers: { default: header }, footers: { default: footer },
    children: C,
  }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log('wrote', OUT); });
