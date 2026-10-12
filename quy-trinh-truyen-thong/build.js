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
const HEAD = ['Bước / thời hạn', 'MHC - Hệ thống Social listening', 'MHC - Team xử lý tin tiêu cực', 'VC - Marketing', 'VC - PD/PM'];
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

// ===== Phần văn bản =====
const SEC = t => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 120 }, keepNext: true,
  children: [new TextRun({ text: t, bold: true, font: FONT, size: 24, color: RED })] });
const T = (...a) => { C.push(table(...a)); C.push(gap()); };

C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 }, children: [
  new TextRun({ text: 'QUY TRÌNH', font: FONT, size: 30, bold: true, color: RED })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [
  new TextRun({ text: 'QUẢN TRỊ RỦI RO TRUYỀN THÔNG VÀ XỬ LÝ THÔNG TIN TIÊU CỰC TRÊN MẠNG XÃ HỘI', font: FONT, size: 26, bold: true, color: RED })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [
  new TextRun({ text: 'Áp dụng trong giai đoạn pilot. Quy trình sẽ được tiếp tục hoàn thiện nếu quá trình triển khai phát sinh bất cập.', font: FONT, size: 19, italics: true, color: GREY })] }));

T([3200, W - 3200], ['Thông tin', 'Nội dung'], [
  ['Đơn vị chủ trì', 'Marketing VietCredit (MKT VC)'],
  ['Đơn vị phối hợp', 'PD/PM, Chăm sóc khách hàng (CS), Thu hồi nợ (Collex), Anti-fraud của VietCredit; đối tác Màn Hình Cộng (MHC)'],
  ['Sản phẩm áp dụng', 'Tất cả sản phẩm, dịch vụ của VietCredit. Thứ tự ưu tiên: Thẻ tín dụng VC x DMX, TikTok BNPL; nếu còn quota thì đến Tin Vay cashloan'],
  ['Kênh theo dõi', 'Facebook, TikTok, YouTube, Threads, báo điện tử'],
  ['Nơi lưu thông tin xử lý', SHEET_URL],
], { boldFirst: true });

C.push(SEC('I. Mục đích'));
C.push(BL('Chủ động phát hiện, quản trị rủi ro truyền thông và xử lý thông tin tiêu cực của khách hàng đăng trên mạng xã hội, báo điện tử.'));
C.push(BL('Quy định rõ việc nào MHC được xử lý ngay, việc nào phải báo VietCredit, và cách VietCredit phối hợp nội bộ khi có sự vụ hoặc khủng hoảng truyền thông.'));

C.push(SEC('II. Từ viết tắt'));
T([3200, W - 3200], ['Viết tắt', 'Giải thích'], [
  ['MHC - Hệ thống SL', 'Hệ thống Social listening của Màn Hình Cộng: theo dõi, cảnh báo, gửi báo cáo'],
  ['Team XLTC MHC', 'Đội xử lý tin tiêu cực của Màn Hình Cộng'],
  ['MKT VC', 'Phòng Marketing VietCredit'],
  ['PD/PM', 'Đơn vị phát triển, quản lý sản phẩm của VietCredit'],
  ['CS VC', 'Chăm sóc khách hàng VietCredit'],
  ['Collex', 'Bộ phận thu hồi nợ VietCredit'],
  ['Anti-fraud VC', 'Bộ phận phòng chống gian lận VietCredit'],
  ['BĐH', 'Ban Điều hành VietCredit'],
  ['Group phối hợp', 'Nhóm chat giữa Team XLTC MHC và MKT VC'],
], { boldFirst: true });

C.push(SEC('III. Phân loại nhóm tin'));
T([1800, W - 1800], ['Nhóm', 'Nội dung tin'], [
  ['Nhóm 1', 'Hotline không liên hệ được; cho rằng "VietCredit lừa đảo"'],
  ['Nhóm 2', 'Lỗi tính năng, không thao tác được (mới phát sinh)'],
  ['Nhóm 3', 'Tin tiêu cực về doanh nghiệp, tình hình kinh doanh, lãnh đạo'],
  ['Nhóm 4', 'Bài viết hỏi về việc bùng nợ, có bị về nhà đòi nợ không; thu nợ làm ảnh hưởng cá nhân'],
  ['Nhóm 5', 'Tin tiêu cực về dịch vụ hỗ trợ rút hạn mức thẻ'],
], { boldFirst: true });

C.push(SEC('IV. Phân loại cách thức xử lý'));
T([1800, 5000, W - 6800], ['Loại', 'Xác định', 'Cách xử lý'], [
  ['A', 'Tin tiêu cực nằm trong tầm kiểm soát, có thể xử lý nhanh theo quy trình', ['• Seeding comment', '• Seeding comment và liên hệ người đăng', '• Comment, report user/fanpage mạo danh; báo PD/PM và Anti-fraud VC xử lý']],
  ['B', 'Sự vụ có nguy cơ lớn trở thành khủng hoảng truyền thông. Kích hoạt quy trình xử lý sự vụ', ['• MHC làm báo cáo sự vụ, gửi MKT VC trong ngày xảy ra sự vụ', '• MKT VC gửi PD/PM liên quan, xin ý kiến xử lý để sự vụ không lan rộng thành khủng hoảng']],
  ['C', 'Khủng hoảng truyền thông. Kích hoạt quy trình xử lý khủng hoảng', ['• MHC làm báo cáo sự vụ, gửi MKT VC trong ngày xảy ra khủng hoảng', '• MKT VC không chờ báo cáo, song song xin ngay ý kiến BĐH, PD/PM về cách xử lý theo các cấp độ (mục V)']],
], { boldFirst: true });
C.push(P('Khi kết hợp nhóm tin và cách xử lý, tin thuộc loại A được gọi theo nhóm: A1 (Nhóm 1), A2 (Nhóm 2), A3 (Nhóm 3), A4 (Nhóm 4), A5 (Nhóm 5).', { italics: true, color: GREY, size: 19 }));

C.push(SEC('V. Các cấp độ xử lý khủng hoảng (loại C)'));
C.push(P('BĐH, PD/PM và MKT VC chọn một hoặc kết hợp các cấp độ sau:', { after: 60 }));
C.push(BL('Im lặng.'));
C.push(BL('Comment theo hướng người dùng khách quan, đưa thông tin giải thích, tích cực.'));
C.push(BL('Lên tin bài tích cực để pha loãng tin bài tiêu cực.'));
C.push(BL('Lên tin bài đính chính, làm rõ tin tiêu cực.'));
C.push(BL('Họp báo làm rõ thông tin.'));

// ===== Flow =====
C.push(new Paragraph({ spacing: { before: 280, after: 60 }, keepNext: true, children: [
  new TextRun({ text: 'VI. Quy trình phối hợp', bold: true, font: FONT, size: 24, color: RED })] }));
C.push(new Paragraph({ spacing: { after: 140 }, keepNext: true, children: [
  new TextRun({ text: 'Cách đọc: ', font: FONT, size: 17, bold: true, color: GREY }),
  new TextRun({ text: '→ chuyển việc   ← nhận việc   ↓ sang bước khác   Ô xám: không có việc.   CS, Collex, Anti-fraud do PD/PM hoặc MKT VC liên hệ khi cần.', font: FONT, size: 17, color: GREY })] }));

G('BƯỚC CHUNG: THEO DÕI VÀ CẢNH BÁO (áp dụng xuyên suốt)');
S('0. Theo dõi, gửi báo cáo', ['Liên tục'], {
  sl: ['• Theo dõi các kênh, gửi báo cáo Social listening Daily, Weekly, Monthly', '• Cảnh báo realtime tin tiêu cực qua SMS, Email', '→ PD/PM, MKT VC, CS VC, Anti-fraud VC', '→ Team XLTC MHC'],
  mkt: ['← nhận báo cáo, cảnh báo realtime'],
  pm: ['← nhận báo cáo, cảnh báo realtime (cùng CS VC, Anti-fraud VC)'],
});

G('GIAI ĐOẠN 1: TIN TIÊU CỰC TỪ POST THỨ 1 ĐẾN POST THỨ 150 TRONG THÁNG (trong gói MHC)');
S('1. Check tin', ['9h00 hằng ngày', 'Dữ liệu từ 9h00 hôm trước đến 9h00 hôm nay'], {
  sl: ['• Cung cấp báo cáo Social listening và realtime alert', '→ Team XLTC MHC'],
  xl: ['← nhận báo cáo, alert', '• Check toàn bộ tin tiêu cực', '**Phạm vi:** tất cả sản phẩm, dịch vụ', '**Ưu tiên:** Thẻ DMX, TikTok BNPL; nếu còn quota thì đến Tin Vay cashloan'],
});
S('2. Phân loại', ['Ngay sau khi check'], {
  xl: ['• Phân tin vào nhóm 1 đến 5 (mục III)', '• Xác định cách xử lý A, B hoặc C (mục IV)', '↓ A1 đến A5: làm theo các dòng A1 đến A5', '↓ B: dòng B', '↓ C: dòng C'],
});
S('A1. Hotline, "VietCredit lừa đảo"', ['Trong ngày'], {
  xl: ['• Chủ động seeding theo hướng kịch bản, không cần hỏi MKT VC', '**Gợi ý hướng comment:** VietCredit luôn đảm bảo dịch vụ nhanh chóng, tiện lợi, minh bạch, an toàn cho khách hàng. Nếu cần hỗ trợ, mời khách hàng liên hệ CSKH VietCredit qua Zalo OA, Hotline, Email, app VietCredit', '↓ lưu thông tin (bước 3)'],
});
S('A2. Lỗi tính năng, không thao tác được', ['Trong ngày'], {
  xl: ['**Lỗi có trong kịch bản MKT VC đã gửi:**', '• Seeding theo kịch bản đó', '↓ lưu thông tin (bước 3)', '**Lỗi ngoài kịch bản:**', '→ báo MKT VC qua group phối hợp', '← nhận hướng xử lý, xử lý tiếp nếu được giao'],
  mkt: ['• Trước đó: gửi Team XLTC kịch bản trả lời cho các lỗi đã lường trước', '← nhận lỗi ngoài kịch bản', '→ thông báo PD/PM liên quan xác minh', '← nhận hướng xử lý', '→ chuyển Team XLTC MHC (nếu MHC xử lý tiếp)'],
  pm: ['← nhận từ MKT VC', '• Xác minh thông tin với các bộ phận cần thiết', '**Gửi lại hướng xử lý:**', '→ MHC xử lý tiếp (qua MKT VC), hoặc', '→ CS VC tư vấn khách hàng'],
});
S('A3. Doanh nghiệp, kinh doanh, lãnh đạo', ['Trong ngày'], {
  xl: ['→ thông báo MKT VC qua group phối hợp', '← xử lý tiếp nếu được giao'],
  mkt: ['← nhận thông tin', '→ xin ý kiến BĐH, PD liên quan', '**Thống nhất hướng xử lý:**', '→ chuyển MHC xử lý tiếp: nếu là tin tiêu cực kiểm soát được, hoặc sự vụ chưa thành khủng hoảng', '↓ VC tự xử lý: nếu là khủng hoảng truyền thông (dòng C)'],
  pm: ['← nhận từ MKT VC', '• Cho ý kiến', '→ MKT VC'],
});
S('A4. Bùng nợ, đòi nợ, thu nợ ảnh hưởng cá nhân', ['Trong ngày, trong quota gói'], {
  xl: ['• Chủ động lọc post nghiêm trọng, tương tác cao để seeding', '**Hướng comment:** VietCredit thu nợ với khách hàng có quan hệ tín dụng theo quy định của pháp luật, trên nguyên tắc bảo vệ quyền lợi và tuân thủ nghĩa vụ của khách hàng. Khách hàng nên thực hiện nghĩa vụ thanh toán để không ảnh hưởng tín dụng cá nhân', '↓ lưu thông tin (bước 3)'],
});
S('A5. Dịch vụ hỗ trợ rút hạn mức thẻ', ['Trong ngày'], {
  xl: ['• Chủ động lọc post nghiêm trọng, tương tác cao để seeding', '**User mạo danh VC** (tự xưng nhân viên VC, mặc áo VC, tên tài khoản có yếu tố VC):', '• Vẫn seeding comment', '• Report user, fanpage mạo danh', '→ thông báo thêm MKT VC', '← phối hợp xử lý giả mạo thương hiệu'],
  mkt: ['← nhận thông tin mạo danh', '→ chuyển Anti-fraud VC, PD/PM', '→ chuyển Team XLTC MHC xử lý giả mạo thương hiệu'],
  pm: ['← nhận từ MKT VC', '• Cùng Anti-fraud VC xử lý'],
});
S('B. Sự vụ có nguy cơ thành khủng hoảng', ['Trong ngày xảy ra sự vụ'], {
  sl: ['• Tiếp tục cảnh báo realtime diễn biến'],
  xl: ['• Làm báo cáo sự vụ', '→ gửi MKT VC trong ngày', '← triển khai hướng xử lý được giao'],
  mkt: ['← nhận báo cáo sự vụ', '→ gửi PD/PM liên quan, xin ý kiến xử lý để không lan rộng thành khủng hoảng', '← nhận ý kiến', '→ chuyển Team XLTC MHC triển khai'],
  pm: ['← nhận từ MKT VC', '• Cho ý kiến xử lý', '→ MKT VC'],
});
S('C. Khủng hoảng truyền thông', ['Ngay khi phát hiện', 'Báo cáo trong ngày'], {
  sl: ['• Tiếp tục cảnh báo realtime diễn biến'],
  xl: ['• Làm báo cáo sự vụ', '→ gửi MKT VC trong ngày xảy ra khủng hoảng', '← triển khai phần việc được giao'],
  mkt: ['• Không chờ báo cáo của MHC', '→ song song xin ngay ý kiến BĐH, PD/PM', '**Cấp độ xử lý (mục V):** im lặng; comment khách quan; tin bài tích cực pha loãng; tin bài đính chính; họp báo', '→ triển khai theo chỉ đạo BĐH'],
  pm: ['← nhận từ MKT VC', '• Cho ý kiến cách xử lý', '→ MKT VC'],
});
S('3. Lưu thông tin', ['Ngay sau khi xử lý xong'], {
  xl: ['• Lưu thông tin xử lý từng case tại link theo dõi chung', '~ ' + SHEET_URL],
  mkt: ['• Theo dõi kết quả trên link'],
});

G('GIAI ĐOẠN 2: TIN TIÊU CỰC TỪ POST THỨ 151 TRONG THÁNG (VietCredit tự theo dõi, xử lý)');
S('4. Chuyển chế độ', ['Khi hết gói 150 comment tiêu cực trong tháng'], {
  sl: ['• Vẫn gửi báo cáo, cảnh báo realtime như bước 0'],
  xl: ['• Chủ động thông báo đã hoàn thành gói 150 comment tiêu cực trong tháng', '→ MKT VC'],
  mkt: ['← nhận thông báo', '• Kích hoạt chế độ VC tự theo dõi, xử lý', '→ thông báo PD/PM, CS VC, MKT VC cùng theo dõi'],
  pm: ['← nhận thông báo', '• Cùng theo dõi; PM chịu trách nhiệm chính'],
});
S('5. Xử lý case sản phẩm', ['Khi phát sinh'], {
  pm: ['• Chủ động raise trong nội bộ VC (CS, MKT, Collex...)', '• Xử lý theo quy trình nội bộ, thường là CS liên hệ xử lý'],
  mkt: ['← phối hợp khi PM raise'],
});
S('6. Xử lý case liên quan toàn công ty', ['Khi phát sinh'], {
  mkt: ['• Chủ động theo dõi', '→ báo cáo PD, BĐH khi xảy ra sự vụ có nguy cơ (loại B) hoặc sự vụ đã thành khủng hoảng (loại C)'],
  pm: ['← nhận báo cáo, phối hợp xử lý'],
});

C.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: COLS, rows: [headRow, ...R] }));

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
    properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 1250, bottom: 900, left: 850, right: 850, header: 450, footer: 400 } } },
    headers: { default: header }, footers: { default: footer },
    children: C,
  }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log('wrote', OUT); });
