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
const W = 9026; // A4 dọc, lề 1440
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
const T = (...a) => { C.push(table(...a)); C.push(gap()); };
const SEC = t => C.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 140 }, keepNext: true,
  children: [new TextRun({ text: t, bold: true, font: FONT, size: 26, color: RED })] }));
const SUB = t => C.push(new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 220, after: 100 }, keepNext: true,
  children: [new TextRun({ text: t, bold: true, font: FONT, size: 22, color: DARK })] }));
const PP = (t, o) => C.push(P(t, o));
const LEAD = t => C.push(new Paragraph({ keepNext: true, spacing: { after: 100, line: 300 }, alignment: AlignmentType.JUSTIFIED, children: runs(t) }));
const BLS = items => items.forEach(t => C.push(BL(t)));
const NOTE = t => PP(t, { italics: true, color: GREY, size: 19 });

// Ô "Đề xuất bổ sung": nền xám nhạt, viền trái đỏ, để tách rõ với nội dung đã chốt
function SUGG(title, items) {
  const kids = [new Paragraph({ spacing: { after: 80 }, keepNext: true, children: [
    new TextRun({ text: 'ĐỀ XUẤT BỔ SUNG' + (title ? ': ' + title : ''), bold: true, font: FONT, size: 19, color: RED })] })];
  items.forEach(t => {
    if (t.startsWith('• ')) kids.push(new Paragraph({ numbering: { reference: 'cbul', level: 0 }, spacing: { after: 50, line: 276 }, children: runs(t.slice(2), { size: 19 }) }));
    else kids.push(new Paragraph({ spacing: { after: 70, line: 276 }, alignment: AlignmentType.JUSTIFIED, children: runs(t, { size: 19 }) }));
  });
  C.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ children: [new TableCell({
    width: { size: W, type: WidthType.DXA }, margins: { top: 100, bottom: 80, left: 180, right: 140 },
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: ALT },
    borders: { top: NB, bottom: NB, right: NB, left: { style: BorderStyle.SINGLE, size: 24, color: RED } },
    children: kids })] })] }));
  C.push(gap());
}

const SHEET_URL = 'https://docs.google.com/spreadsheets/d/1xp3EIZZXDeLg45_3XUYrvG8A6sjjVtlwraIty1DJtLE';

// ===== Bìa =====
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 60 }, children: [
  new TextRun({ text: 'TÀI LIỆU DIỄN GIẢI', font: FONT, size: 22, bold: true, color: GREY })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [
  new TextRun({ text: 'QUY TRÌNH QUẢN TRỊ RỦI RO TRUYỀN THÔNG', font: FONT, size: 30, bold: true, color: RED })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [
  new TextRun({ text: 'VÀ XỬ LÝ THÔNG TIN TIÊU CỰC', font: FONT, size: 30, bold: true, color: RED })] }));
T([2600, W - 2600], ['Thông tin', 'Nội dung'], [
  ['Phiên bản', 'v4, áp dụng giai đoạn pilot. Quy trình tiếp tục được hoàn thiện nếu triển khai phát sinh bất cập'],
  ['Đơn vị chủ trì', 'Marketing VietCredit (MKT VC)'],
  ['Đơn vị phối hợp', 'PD/PM, CS, Collection, Antifraud của VietCredit; CS và PD/PM đối tác; Màn Hình Cộng (MHC)'],
  ['Sản phẩm', 'Tất cả sản phẩm, dịch vụ. Trong gói MHC ưu tiên: Thẻ tín dụng VC x DMX, TikTok BNPL; còn quota thì đến Tin Vay cashloan'],
  ['Gói dịch vụ MHC', '150 phản hồi tiêu cực mỗi tháng, đến hết 31/12/2026. Gói năm 2027 đàm phán sau'],
], { boldFirst: true });
C.push(new Paragraph({ spacing: { before: 120, after: 80 }, children: runs('**Cách đọc tài liệu:** nội dung chính là quy trình v4 đã chốt. Các ô nền xám có chữ "ĐỀ XUẤT BỔ SUNG" là gợi ý thêm cho những điểm được ghi chú cần làm rõ hoặc tối ưu; các gợi ý này cần được duyệt trước khi đưa vào bản chính thức.', { size: 19, color: GREY }) }));

// ===== 0. Tổng quan =====
SEC('Tổng quan: ba luồng xử lý');
LEAD('Quy trình chia thành ba luồng, tùy nguồn thông tin và thời điểm trong tháng:');
T([1300, 2700, 2500, 2526], ['Luồng', 'Áp dụng cho', 'Thời điểm', 'Người phụ trách chính'], [
  ['Mục V', 'Nguồn (1): báo cáo Social listening của MHC', 'Từ ngày 1 của tháng đến khi hết quota 150 phản hồi', 'Team MHC XLTC, MKT VC, PD/PM VC'],
  ['Mục VI', 'Nguồn (1): báo cáo Social listening của MHC', 'Từ phản hồi thứ 151 trong tháng', 'PM chính hoặc PM vận hành do PD chỉ định'],
  ['Mục VII', 'Nguồn (2), (3): báo cáo CS, theo dõi thủ công, thông tin từ cá nhân', 'Luôn áp dụng', 'PD/PM (sản phẩm); Phòng Truyền thông & Vận hành marketing (thương hiệu)'],
], { boldFirst: true });
LEAD('Cả ba luồng dùng chung: cách phân loại nhóm tin (mục II), mức độ (mục III) và bảng mã (mục IV).');
SUGG('ranh giới giữa mục VI và mục VII', [
  'Bản v4 ghi mục VI áp dụng cả với nguồn (2), (3), trong khi mục VII cũng dành cho nguồn (2), (3). Đề xuất tách rõ:',
  '• Mục VI chỉ áp dụng cho nguồn (1) từ phản hồi thứ 151.',
  '• Nguồn (2), (3) luôn đi theo mục VII (quy trình hiện hành của PD/PM hoặc của Phòng Truyền thông & Vận hành marketing).',
  '• Riêng các case từ nguồn (2), (3) được xếp mức B, C hoặc thuộc Nhóm 5 thì người phát hiện báo ngay MKT VC để xử lý theo bước B, C ở mục VI.',
]);

// ===== I. Nguồn thông tin =====
SEC('I. Phân loại nguồn thông tin');
SUB('1. Nguồn bị động: phản hồi của khách hàng và thị trường');
T([700, 3400, 2300, 2626], ['Mã', 'Nguồn', 'Người cung cấp', 'Người nhận'], [
  ['(1)', 'Báo cáo Social listening của MHC (Realtime Alert, Daily, Weekly, Monthly) và các báo cáo phân tích từ báo cáo này', 'MHC; PD/PM gửi báo cáo phân tích weekly, monthly hoặc theo giai đoạn', 'PD, PM, MKT, CS, Collection, Antifraud VC'],
  ['(2)', 'Báo cáo của CS VC và CS đối tác', 'CS VC, CS đối tác', 'PD, PM, MKT VC'],
  ['(3)', 'Theo dõi thủ công của PD/PM, các BU; thông tin do cá nhân thuộc đối tác, BĐH, PD/PM, CBNV báo', 'Người phát hiện', 'PD/PM sản phẩm liên quan'],
], { boldFirst: true });
SUB('2. Nguồn chủ động phòng ngừa: update, bảo trì, tạm ngừng tính năng');
PP('PD/PM chủ động thông tin cho CS và MKT về các đợt update, bảo trì, tạm ngừng tính năng để cùng xây dựng trước kịch bản trả lời khách hàng và kịch bản phản hồi comment tiêu cực trên mạng xã hội.');
SUGG('để nguồn thông tin chạy đều', [
  '• Nguồn (2): CS VC và CS đối tác gửi báo cáo theo tuần (ví dụ thứ Hai), gồm số khiếu nại theo nhóm tin ở mục II, để so sánh được với số liệu Social listening.',
  '• Nguồn (3): chọn một kênh tiếp nhận chung (một nhóm chat hoặc một form ngắn) với 4 thông tin: link hoặc ảnh chụp, sản phẩm, nội dung tóm tắt, người báo. Tránh việc thông tin gửi rải rác qua tin nhắn cá nhân.',
  '• Nguồn phòng ngừa: PD/PM báo CS, MKT trước ít nhất 2 ngày làm việc với thay đổi có kế hoạch; với sự cố đột xuất thì báo ngay khi phát hiện. MKT VC gửi kịch bản đã thống nhất cho Team MHC XLTC trước giờ thay đổi có hiệu lực.',
]);

// ===== II. Nhóm tin =====
SEC('II. Phân loại nhóm tin');
NOTE('Giai đoạn pilot, tiếp tục hoàn thiện nếu triển khai có bất cập.');
T([1500, 3400, 4126], ['Nhóm', 'Định nghĩa (bản chốt)', 'Ví dụ, dấu hiệu nhận biết (đề xuất bổ sung)'], [
  ['Nhóm 1: CSKH', 'Hotline không liên hệ được và các phản hồi khác về cách trả lời, hỗ trợ của CSKH', ['• Gọi hotline không ai nghe, chờ lâu, bị ngắt máy', '• Nhắn Zalo OA, email không được trả lời', '• Tổng đài viên trả lời sai, thiếu, thái độ chưa tốt', '• Khiếu nại đã gửi nhưng chưa được giải quyết', '• "VietCredit lừa đảo" nói chung chung, không kèm sự việc cụ thể']],
  ['Nhóm 2: Collection', 'Hỏi về việc bùng nợ vay, thẻ; có bị về nhà đòi nợ không; bức xúc với thái độ team thu nợ làm ảnh hưởng cá nhân; bức xúc nên bôi nhọ VietCredit; không vay, không mở thẻ mà bị thu nợ', ['• "Bùng VietCredit có sao không", "có bị xuống nhà không"', '• Bị gọi cho người thân, đồng nghiệp, bị nhắn tin đe dọa', '• Đăng ảnh, ghi âm cuộc gọi nhắc nợ', '• Bị đòi nợ dù không vay (có thể bị mạo danh hồ sơ hoặc nhầm số)']],
  ['Nhóm 3: Sản phẩm, tính năng', 'Lỗi tính năng mới phát sinh, không thao tác được; phản hồi tiêu cực về tính năng, chính sách, biểu lãi phí chưa hợp lý, chưa thuận tiện, chưa đúng mong muốn; thông tin sai lệch, hiểu nhầm về tính năng, chính sách, biểu lãi phí', ['• App lỗi đăng nhập, không nhận OTP, không thanh toán được', '• Giao dịch thẻ bị từ chối, TikTok BNPL không hiện lựa chọn trả sau', '• Chê lãi, phí cao; có khoản phí không hiểu', '• Hạn mức thấp, duyệt hồ sơ lâu, bị từ chối không rõ lý do', '• Hiểu nhầm chính sách (ví dụ nghĩ 0% nhưng có phí)']],
  ['Nhóm 4: Fraud, trục lợi', 'Đối tượng có dấu hiệu trục lợi: dịch vụ trung gian mở thẻ, rút hạn mức thẻ, ví trả sau, hỗ trợ mở khoản vay; quảng cáo chương trình của VietCredit để nhận thưởng; thủ đoạn lừa khách hàng; đặc biệt có dấu hiệu mạo danh VietCredit', ['• Bài "rút tiền thẻ tín dụng, ví trả sau phí thấp"', '• "Mở thẻ VietCredit bao đậu, hỗ trợ nợ xấu"', '• Tài khoản dùng tên, logo, đồng phục VietCredit; tự xưng nhân viên', '• App, web, page, tài khoản TikTok trùng hoặc gần giống tên thương hiệu', '• Yêu cầu khách chuyển phí trước khi giải ngân']],
  ['Nhóm 5: Thương hiệu doanh nghiệp', 'Tin tiêu cực về doanh nghiệp, tình hình kinh doanh, lãnh đạo', ['• Tin về tình hình tài chính, cổ đông, sở hữu', '• Tin về lãnh đạo, nhân sự cấp cao', '• Cáo buộc VietCredit "cho vay nặng lãi", "tín dụng đen" mang tính hệ thống', '• Báo điện tử, kênh tin tức đưa tin tiêu cực về công ty']],
  ['Nhóm 6: Khác', 'Thông tin tiêu cực không thể phân loại vào các nhóm trên', ['• Tuyển dụng, môi trường làm việc (nếu lan rộng thì xét Nhóm 5)', '• So sánh với đối thủ theo hướng chê VietCredit', '• Tin chung về ngành tài chính tiêu dùng có nhắc tên VietCredit', '• Bài cũ được đăng lại']],
], { boldFirst: true });
SUGG('quy tắc khi một bài thuộc nhiều nhóm', [
  'Xếp vào nhóm có rủi ro cao hơn theo thứ tự: Nhóm 5 > Nhóm 4 > Nhóm 2 > Nhóm 3 > Nhóm 1 > Nhóm 6. Ghi chú nhóm phụ ở cột ghi chú để các bên liên quan cùng nắm.',
  'Một số trường hợp hay nhầm:',
  '• "VietCredit lừa đảo" kèm sự việc thu phí, lãi cụ thể: Nhóm 3. Kèm chuyện đòi nợ: Nhóm 2. Do tài khoản giả danh VietCredit gây ra: Nhóm 4. Không kèm sự việc: Nhóm 1.',
  '• "Không vay mà bị đòi nợ": Nhóm 2, đồng thời báo Antifraud kiểm tra khả năng hồ sơ bị mạo danh.',
  '• Bài hỏi "nên vay bên nào" có chê VietCredit: xếp theo nội dung bị chê (lãi phí là Nhóm 3, thu nợ là Nhóm 2).',
]);

// ===== III. Mức độ =====
SEC('III. Phân loại mức độ');
T([1300, W - 1300], ['Mức độ', 'Định nghĩa (bản chốt)'], [
  ['A', 'Tin tiêu cực trong tầm kiểm soát'],
  ['B', 'Sự vụ có nguy cơ trở thành khủng hoảng truyền thông, có thể khống chế, xử lý, ngăn chặn'],
  ['C', 'Sự vụ đã trở thành khủng hoảng truyền thông'],
], { boldFirst: true });
SUGG('tiêu chí tham khảo để xếp mức độ', [
  'Một case được xếp ở mức cao nhất mà nó đạt từ bất kỳ tiêu chí nào trong bảng dưới. Khi phân vân giữa hai mức thì chọn mức cao hơn. Mức độ có thể nâng lên trong quá trình theo dõi. Các ngưỡng số là đề xuất ban đầu, chốt lại sau tháng pilot đầu tiên dựa trên số liệu thực tế.',
]);
T([1700, 2400, 2500, 2426], ['Tiêu chí', 'Mức A', 'Mức B', 'Mức C'], [
  ['Lan tỏa (nguồn 1)', 'Bài đơn lẻ, dưới 100 tương tác (thích, bình luận, chia sẻ)', ['• Từ 100 tương tác trở lên, hoặc', '• từ 5 bài cùng chủ đề trong 24 giờ, hoặc', '• nguồn đăng có ảnh hưởng: page, group từ 100.000 thành viên, KOL, KOC'], ['• Báo điện tử, truyền hình đưa tin, hoặc', '• lan từ 2 nền tảng trở lên với tổng trên 1.000 tương tác trong 24 giờ, hoặc', '• có hashtag, kêu gọi tẩy chay, hành động tập thể']],
  ['Tính chất nội dung', 'Phàn nàn cá nhân, câu hỏi, hiểu nhầm; trả lời được bằng kịch bản sẵn có', ['• Cáo buộc vi phạm pháp luật, quy định (thu nợ sai luật, thu phí sai)', '• Nhiều khách hàng cùng phản ánh một lỗi', '• Mạo danh đã có người bị hại'], ['• Lỗi hệ thống ảnh hưởng diện rộng', '• Sự cố bảo mật, lộ dữ liệu khách hàng', '• Thông tin sai về tài chính, lãnh đạo lan rộng']],
  ['Nguồn (2): CS', 'Khiếu nại đơn lẻ', ['• Từ 20 khiếu nại cùng vấn đề trong 1 ngày, hoặc gấp đôi mức trung bình', '• Khách hàng dọa đăng mạng xã hội, báo chí'], ['• Khiếu nại hàng loạt đã lên mạng xã hội, báo chí', '• Cơ quan quản lý (NHNN, công an) nhận đơn hoặc đề nghị làm việc']],
  ['Nguồn (3): cá nhân báo', 'Thông tin đơn lẻ, chưa lan truyền', 'BĐH, đối tác cảnh báo một vấn đề có nguy cơ lan rộng', 'Đã có trên báo chí hoặc cơ quan chức năng vào cuộc'],
  ['Riêng Nhóm 5', 'Lan tỏa thấp, chỉ là bình luận lẻ', 'Mặc định từ mức B nếu lan tỏa từ trung bình trở lên', 'Theo tiêu chí chung'],
], { boldFirst: true });
SUGG('ai xác định mức độ', [
  '• Trong gói MHC (mục V): Team MHC XLTC đề xuất mức độ; MKT VC xác nhận các case mức B, C.',
  '• Ngoài gói (mục VI) và nguồn (2), (3): PM phụ trách đề xuất; MKT VC xác nhận các case mức B, C.',
]);

// ===== IV. Bảng mã =====
SEC('IV. Bảng mã các trường hợp cần xử lý');
LEAD('Mỗi case được gắn một mã = mức độ + số nhóm. Ví dụ: tin về hotline ở mức A là A1; sự vụ về thu nợ có nguy cơ khủng hoảng là B2.');
const MW = [1700, 1221, 1221, 1221, 1221, 1221, 1221];
T(MW, ['Mức độ', 'Nhóm 1 CSKH', 'Nhóm 2 Collection', 'Nhóm 3 Sản phẩm', 'Nhóm 4 Fraud', 'Nhóm 5 Thương hiệu', 'Nhóm 6 Khác'], [
  ['A: trong tầm kiểm soát', 'A1', 'A2', 'A3', 'A4', 'A5', 'A6'],
  ['B: có nguy cơ khủng hoảng', 'B1', 'B2', 'B3', 'B4', 'B5', 'B6'],
  ['C: đã là khủng hoảng', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6'],
], { boldFirst: true });
NOTE('Bản v4 ghi "chia theo bảng mã từ A1 đến C5"; theo bảng mã đầy đủ thì phạm vi là A1 đến C6.');

// ===== V. Trong gói MHC =====
SEC('V. Xử lý trong phạm vi gói MHC XLTC (nguồn 1, phản hồi thứ 1 đến 150)');
T([2400, W - 2400], ['Hạng mục', 'Quy định'], [
  ['Áp dụng cho', 'Tin tiêu cực ghi nhận, thống kê từ nguồn (1) báo cáo Social listening của MHC; 150 phản hồi tiêu cực đầu tiên trong tháng'],
  ['Thời điểm', 'Từ ngày 1 của tháng đến ngày hết quota 150. MHC thông báo khi hết quota. Gói áp dụng đến hết 31/12/2026'],
  ['Nguồn thông tin', ['• MHC gửi Realtime Alert, Daily, Weekly, Monthly cho PD, PM, MKT, CS, Collection, Antifraud VC', '• PD/PM gửi báo cáo phân tích weekly, monthly hoặc theo giai đoạn']],
  ['Theo dõi, phân loại', 'PIC: Team MHC XLTC. Phân loại theo mục II, III, IV'],
  ['Xử lý, báo cáo', 'PIC và cách thức theo từng mã. Các bên tham gia: Team MHC XLTC, MKT VC, PD/PM VC (PD/PM tự làm việc với nội bộ và đối tác)'],
  ['Lịch check tin', '9h00 hằng ngày, dữ liệu từ 9h00 hôm trước đến 9h00 hôm nay. Phạm vi: tất cả sản phẩm, ưu tiên Thẻ DMX, TikTok BNPL, còn quota thì đến Tin Vay cashloan'],
], { boldFirst: true });

SUB('Cách xử lý từng mã');
T([1000, 2700, W - 3700], ['Mã', 'PIC', 'Cách xử lý'], [
  ['A1', 'Team MHC XLTC', ['Chủ động seeding theo kịch bản.', 'Hướng comment: VietCredit luôn đảm bảo dịch vụ nhanh chóng, tiện lợi, minh bạch, an toàn cho khách hàng; nếu cần hỗ trợ, mời khách hàng liên hệ CSKH VietCredit qua Zalo OA, Hotline, Email, app VietCredit.']],
  ['A2', 'Team MHC XLTC', ['Chủ động lọc post nghiêm trọng, tương tác cao để seeding.', 'Hướng comment: VietCredit thu nợ với khách hàng có quan hệ tín dụng theo quy định của pháp luật, trên nguyên tắc bảo vệ quyền lợi và yêu cầu khách hàng tuân thủ nghĩa vụ. Khách hàng nên thực hiện nghĩa vụ thanh toán để không ảnh hưởng tín dụng cá nhân.']],
  ['A3', 'Team MHC XLTC; MKT VC; PD/PM', ['• Lỗi đã có kịch bản (PD/PM đã thông báo, CS và MKT đã chuẩn bị, MKT đã gửi MHC): MHC xử lý theo kịch bản.', '• Lỗi ngoài kịch bản: MHC báo MKT qua group phối hợp. MKT báo PD/PM liên quan. PD/PM xác minh với các bộ phận cần thiết, gửi lại hướng xử lý cho MKT. Sau đó: MKT chuyển MHC xử lý tiếp, hoặc PD/PM chuyển CS VC tư vấn khách hàng.']],
  ['A4', 'Team MHC XLTC; MKT VC; Antifraud VC', ['Chủ động lọc post nghiêm trọng, tương tác cao để seeding.', 'User mạo danh VietCredit (tự xưng nhân viên, mặc áo VC, tên tài khoản có yếu tố VC): vẫn seeding, đồng thời báo MKT VC. MKT chuyển thông tin cho Antifraud VC và Team MHC xử lý giả mạo app, web, page, tài khoản TikTok trùng tên thương hiệu hoặc gây nhầm lẫn.']],
  ['A5', 'MKT VC; BĐH; PD', ['MHC báo MKT qua group phối hợp. MKT xin ý kiến BĐH, PD liên quan để thống nhất hướng:', '• tin kiểm soát được hoặc sự vụ chưa thành khủng hoảng: MKT chuyển MHC xử lý tiếp;', '• đã là khủng hoảng: MKT tự xử lý (theo bước C).']],
  ['A6', 'Đề xuất: Team MHC XLTC; MKT VC', ['(Bản v4 chưa quy định.) Đề xuất: MHC báo MKT qua group phối hợp kèm hướng đề xuất. MKT quyết định theo dõi, seeding hoặc chuyển đơn vị liên quan. Nếu một dạng tin Nhóm 6 lặp lại nhiều lần, bổ sung thành kịch bản để lần sau MHC tự xử lý.']],
  ['B1 đến B6', 'MKT VC điều phối; PD/PM', ['• MHC làm báo cáo sự vụ, gửi MKT VC trong ngày xảy ra sự vụ.', '• MKT VC gửi PD/PM liên quan, xin ý kiến xử lý để không lan rộng thành khủng hoảng.']],
  ['C1 đến C6', 'MKT VC; BĐH; PD/PM', ['• MHC làm báo cáo sự vụ, gửi MKT VC trong ngày xảy ra khủng hoảng.', '• MKT VC không chờ báo cáo, song song xin ngay ý kiến BĐH, PD/PM về cách xử lý theo các cấp độ: im lặng; comment theo hướng người dùng khách quan đưa thông tin giải thích, tích cực; lên tin bài tích cực để pha loãng; lên tin bài đính chính, làm rõ; họp báo làm rõ thông tin.']],
], { boldFirst: true });
PP('**Đóng sự việc:** sau sự việc, các PIC xử lý báo cáo lại PD/PM. PD/PM đóng sự việc, hoặc tùy case làm thêm báo cáo rủi ro hoạt động.');

SUGG('B và C: giữ một quy trình chung, chỉ khác đầu mối theo nhóm', [
  'Không cần chia nhỏ B1 đến B6 thành sáu quy trình riêng vì các bước giống nhau (báo cáo sự vụ, MKT điều phối, xin ý kiến, triển khai). Khác biệt nằm ở đơn vị MKT VC cần hỏi và đơn vị cung cấp thông tin. Đề xuất bảng đầu mối dưới đây áp dụng cho cả B và C.',
]);
T([2200, 3200, 3626], ['Nhóm', 'Đầu mối chuyên môn (B, C)', 'Ghi chú'], [
  ['1: CSKH', 'PD/PM; Trưởng CS', 'CS kiểm tra số cuộc gọi nhỡ, năng lực tổng đài'],
  ['2: Collection', 'PD/PM; Trưởng Collection', 'Kiểm tra lịch sử nhắc nợ, ghi âm, đối tác nhắc nợ'],
  ['3: Sản phẩm', 'PD/PM sản phẩm', 'Xác nhận lỗi, thời gian khắc phục, số khách hàng bị ảnh hưởng'],
  ['4: Fraud', 'Antifraud VC; Pháp chế', 'Cân nhắc cảnh báo khách hàng trên kênh chính thức, trình báo cơ quan chức năng'],
  ['5: Thương hiệu', 'BĐH; PD', 'Mức B trở lên luôn báo BĐH'],
  ['6: Khác', 'MKT VC quyết định đầu mối', ''],
], { boldFirst: true });
SUGG('mốc thời gian cho B, C', [
  '• Mức B: MHC báo nhanh qua group phối hợp trong 2 giờ từ khi phát hiện, báo cáo sự vụ đầy đủ trong ngày. MKT gửi PD/PM trong 1 giờ làm việc. PD/PM gửi ý kiến trong 4 giờ làm việc.',
  '• Mức C: MHC gọi điện và nhắn group trong 30 phút, kể cả ngoài giờ. MKT báo BĐH trong 1 giờ. Thống nhất cấp độ xử lý trong 4 giờ.',
  '• Đóng sự việc: báo cáo rủi ro hoạt động là bắt buộc với mọi case mức B, C; với mức A khi cùng một vấn đề lặp lại từ 3 lần trong tháng hoặc có thiệt hại cho khách hàng.',
  '• MHC tiếp tục lưu thông tin xử lý từng case tại link theo dõi chung: ' + SHEET_URL,
]);

// ===== VI. VC tự vận hành =====
SEC('VI. Xử lý do nội bộ VietCredit tự vận hành (nguồn 1, từ phản hồi thứ 151)');
T([2400, W - 2400], ['Hạng mục', 'Quy định'], [
  ['Áp dụng cho', 'Tin tiêu cực ghi nhận, thống kê từ nguồn (1), từ phản hồi thứ 151 trong tháng'],
  ['Thời điểm', 'Từ ngày hết quota 150 của MHC. MHC thông báo MKT VC; MKT VC thông báo PD/PM để kích hoạt quy trình'],
  ['Nguồn thông tin', ['• MHC gửi Realtime Alert, Daily, Weekly, Monthly cho PD, PM, MKT, CS, Collection, Antifraud VC', '• PD/PM gửi báo cáo phân tích weekly, monthly hoặc theo giai đoạn']],
  ['Theo dõi, phân loại', 'PIC: PM chính hoặc PM vận hành theo chỉ định của PD. Phân loại theo mục II, III, IV'],
  ['Xử lý, báo cáo', 'PIC và cách thức theo từng mã. Các bên tham gia: MKT VC, PD/PM VC, CS VC, Collection VC, CS đối tác, PD/PM đối tác'],
], { boldFirst: true });
PP('**Nội dung đã chốt (gợi ý trong v4):** PM theo dõi, tự phân loại và raise các case nghiêm trọng, sự vụ có nguy cơ. MKT VC và CS phối hợp chuẩn bị kịch bản để CS VC trả lời comment bằng tài khoản chính thống của VietCredit. Song song, PD/PM chuyển Antifraud nghiên cứu xử lý, khởi tố; hoặc chuyển Team MHC report tài khoản, app, web, page, tài khoản TikTok giả mạo trùng tên thương hiệu hoặc gây nhầm lẫn.');
SUGG('sắp xếp lại theo từng mã để dễ làm', []);
T([1000, 2700, W - 3700], ['Mã', 'PIC', 'Cách xử lý (đề xuất)'], [
  ['A1', 'PM; CS VC', 'CS VC trả lời comment bằng tài khoản chính thống theo kịch bản CSKH; liên hệ khách hàng nếu có thông tin.'],
  ['A2', 'PM; Collection; CS VC', 'Collection xác nhận hướng trả lời; CS VC trả lời bằng tài khoản chính thống. Case "không vay mà bị đòi nợ" chuyển thêm Antifraud.'],
  ['A3', 'PM; CS VC', 'PM xác minh lỗi; dùng kịch bản đã có hoặc cùng CS, MKT soạn kịch bản mới; CS VC trả lời hoặc liên hệ khách hàng.'],
  ['A4', 'PM; Antifraud; MKT VC', 'PM chuyển Antifraud nghiên cứu xử lý, khởi tố. MKT chuyển Team MHC report tài khoản, app, web, page giả mạo (xác nhận với MHC có tính phí ngoài gói không).'],
  ['A5', 'MKT VC', 'PM báo MKT; MKT xin ý kiến BĐH, PD như mã A5 ở mục V.'],
  ['A6', 'PM; MKT VC', 'PM tự xử lý nếu thuộc phạm vi sản phẩm, còn lại chuyển MKT.'],
  ['B, C', 'MKT VC điều phối', 'PM raise ngay cho MKT. Các bước như B, C ở mục V, trừ việc báo cáo sự vụ do PM lập thay cho MHC (hoặc đề nghị MHC hỗ trợ nếu cần).'],
], { boldFirst: true });
SUGG('lưu ý vận hành sau khi hết quota', [
  '• Không dùng tài khoản cá nhân của nhân viên để seeding. Chỉ trả lời bằng tài khoản chính thống của VietCredit.',
  '• PM giữ lịch check tin cố định như MHC (9h00 hằng ngày) và ghi case vào cùng link theo dõi để báo cáo tháng có đủ số liệu cả trong và ngoài gói.',
  '• Làm rõ với MHC cách tính quota: "150 phản hồi" là 150 bài, 150 comment seeding hay 150 case. Bản v4 đang dùng cả ba cách gọi.',
]);
PP('**Đóng sự việc:** sau sự việc, các PIC xử lý báo cáo lại PD/PM. PD/PM đóng sự việc, hoặc tùy case làm thêm báo cáo rủi ro hoạt động.');

// ===== VII. Nguồn 2, 3 =====
SEC('VII. Xử lý các case từ nguồn (2), (3) và nguồn khác');
T([2400, W - 2400], ['Hạng mục', 'Quy định'], [
  ['Áp dụng cho', 'Tin tiêu cực, vấn đề phát sinh từ nguồn (2) báo cáo CS VC, CS đối tác; nguồn (3) theo dõi thủ công của PD/PM, BU, thông tin từ cá nhân thuộc đối tác, BĐH, PD/PM, CBNV; và nguồn khác'],
  ['Thời điểm', 'Luôn áp dụng. VietCredit chủ động thực hiện'],
  ['Phạm vi sản phẩm, dự án', 'Theo quy trình hiện hành do PD/PM dịch vụ quy định về xử lý, phối hợp trong nội bộ và với đối tác'],
  ['Phạm vi thương hiệu doanh nghiệp', 'Theo quy trình hiện hành của Phòng Truyền thông & Vận hành marketing về xử lý tin tiêu cực liên quan danh tiếng doanh nghiệp'],
], { boldFirst: true });
SUGG('nối với quy trình truyền thông', [
  '• Dù xử lý theo quy trình hiện hành, case từ nguồn (2), (3) nếu đạt mức B, C hoặc thuộc Nhóm 5 thì người phát hiện báo ngay MKT VC.',
  '• Gắn mã theo mục IV cho mọi case để báo cáo tháng gộp được số liệu từ cả ba nguồn.',
]);

// ===== VIII. Điểm cần chốt =====
SEC('VIII. Tóm tắt các điểm cần chốt');
T([600, 5200, 3226], ['#', 'Nội dung', 'Mục'], [
  ['1', 'Ranh giới mục VI và mục VII với nguồn (2), (3)', 'Tổng quan'],
  ['2', 'Kênh tiếp nhận chung cho nguồn (3); tần suất báo cáo của CS', 'I'],
  ['3', 'Ví dụ nhận biết từng nhóm; thứ tự ưu tiên khi một bài thuộc nhiều nhóm', 'II'],
  ['4', 'Tiêu chí và ngưỡng số để xếp mức A, B, C; ai xác định mức độ', 'III'],
  ['5', 'Phạm vi bảng mã: A1 đến C6', 'IV'],
  ['6', 'Cách xử lý A6', 'V'],
  ['7', 'Giữ B, C là một quy trình chung, kèm bảng đầu mối theo nhóm; mốc thời gian B, C', 'V'],
  ['8', 'Khi nào bắt buộc báo cáo rủi ro hoạt động', 'V'],
  ['9', 'Cách xử lý từng mã sau khi hết quota; chỉ trả lời bằng tài khoản chính thống', 'VI'],
  ['10', 'Cách tính quota 150 với MHC; phí report tài khoản giả mạo ngoài gói', 'VI'],
], { boldFirst: true });

// ---------- document ----------
const doc = new Document({
  creator: 'Marketing VietCredit',
  title: 'Tài liệu diễn giải quy trình quản trị rủi ro truyền thông',
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
