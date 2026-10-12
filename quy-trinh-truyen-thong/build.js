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
const COLS = [1350, 1450, 1500, 6188, 2000, 2650]; // tổng = W (A4 ngang, lề 850)

function groupRow(text) {
  return new TableRow({ cantSplit: true, children: [new TableCell({
    columnSpan: COLS.length, width: { size: W, type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: 120, right: 100 },
    borders: { top: TB, bottom: TB, left: TB, right: TB },
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: 'D8D9E0' },
    children: [new Paragraph({ keepNext: true, children: [new TextRun({ text, bold: true, font: FONT, size: 19, color: RED })] })],
  })] });
}
let rowIdx = 0;
function stepRow(cells) {
  const fill = rowIdx++ % 2 ? ALT : 'FFFFFF';
  // Chỉ cho phép ngắt trang giữa hàng với các bước dài, tránh mẩu lẻ ở đầu trang
  const len = cells.flat().join(' ').length;
  return new TableRow({ cantSplit: len < 1100, children: cells.map((c, i) => new TableCell({
    width: { size: COLS[i], type: WidthType.DXA }, verticalAlign: VerticalAlign.TOP,
    margins: { top: 70, bottom: 50, left: 100, right: 100 },
    borders: { top: TB, bottom: TB, left: TB, right: TB },
    shading: { type: ShadingType.CLEAR, color: 'auto', fill },
    children: cellParas(c, { bold: i === 0, color: '000000' }),
  })) });
}
const headRow = new TableRow({ tableHeader: true, cantSplit: true, children:
  ['Bước', 'Thời gian / hạn', 'Ai làm', 'Làm thế nào', 'Kết quả, ghi trên link', 'Lưu ý'].map((h, i) => new TableCell({
    width: { size: COLS[i], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
    margins: { top: 70, bottom: 50, left: 100, right: 100 },
    borders: { top: TB, bottom: TB, left: TB, right: TB },
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: RED },
    children: cellParas(h, { bold: true, color: 'FFFFFF', align: AlignmentType.CENTER }),
  })) });

const R = [];
const G = t => R.push(groupRow(t));
const S = (...cells) => R.push(stepRow(cells));

// ===== A =====
G('A. THIẾT LẬP BAN ĐẦU (làm một lần trước khi áp dụng, rà lại mỗi quý)');
S('0. Chuẩn bị và chốt quy trình', 'Tuần đầu triển khai', ['Marketing VC', 'Phối hợp: PD/PM, CS, Thu hồi nợ, MHC, BĐH'], [
  '**Phạm vi:** Thẻ tín dụng VC x DMX và TikTok BNPL. Các dự án khác đang có Social listening: MHC theo dõi như hiện tại; nếu có sự vụ khẩn thì xử lý theo bước K1 đến K4.',
  '**Kênh:** Facebook, TikTok, YouTube, Threads, báo điện tử.',
  '**Đầu mối:** VietCredit là Marketing [tên, số điện thoại]; MHC là Hằng, Lam. Lập nhóm chat [Zalo/Telegram] gồm đầu mối hai bên và Trưởng Marketing.',
  '**Link theo dõi chung:** ' + SHEET_URL + ' (sheet "Tư vấn XLKH" mỗi dòng một case; sheet "Seeding" mỗi dòng một bình luận). Thêm các cột: Tương tác, Ý kiến VC, Người duyệt, Nghiệm thu, Yêu cầu bổ sung, Kiểm tra sau 48 giờ.',
  '**Bộ từ khóa:** PD/PM cung cấp tên sản phẩm, cách khách hàng hay gọi, tên đối tác (Điện Máy Xanh, TikTok Shop); Marketing gửi MHC.',
  '**Kịch bản mẫu Nhóm 1, 4:** MHC soạn, CS và Thu hồi nợ xác nhận nội dung, Marketing duyệt.',
  '**BĐH quyết:** phạm vi giai đoạn 1; cho MHC tự seeding Nhóm 1, 4; hạn mức chi phí Trưởng Marketing được duyệt; người phát ngôn khi có sự vụ khẩn.',
  '**MHC xác nhận:** giờ cập nhật link 13:30; các cột mới trên link; thời hạn bổ sung (bước 9); luồng sự vụ khẩn (K1 đến K4).',
], 'Link có đủ cột; kịch bản mẫu đã duyệt; nhóm chat đã lập', [
  'Ý kiến, duyệt, nghiệm thu đều ghi trên link để lưu vết. Nhóm chat chỉ dùng để báo "đã cập nhật" và việc gấp.',
  'Khi đổi phí, chính sách, bảo trì hệ thống, chạy chiến dịch lớn: PD/PM báo Marketing trước ít nhất 2 ngày làm việc để chuẩn bị câu trả lời và báo MHC.',
]);

// ===== B =====
G('B. XỬ LÝ HẰNG NGÀY (thứ Hai đến thứ Sáu)');
S('1. Ghi nhận, cảnh báo', ['Liên tục', 'Khung tổng hợp: 9h00 hôm trước đến 9h00 hôm nay'], 'MHC', [
  'Theo dõi trên hệ thống Crisis theo bộ từ khóa, gửi alert cho VietCredit.',
  'Nếu tin thuộc sự vụ khẩn thì chuyển ngay sang bước K1: tin về lãnh đạo, tình hình kinh doanh, bảo mật, lỗi sản phẩm hàng loạt, truyền thông sai sự thật; hoặc bài có trên 100 tương tác; hoặc lên báo điện tử.',
], 'Tin đã alert, có trên hệ thống Crisis', '');
S('2. Phân nhóm, đề xuất phương án', ['Buổi sáng', 'Cập nhật link lúc 13:30'], 'MHC', [
  '**Phân nhóm mỗi tin:**',
  '• Nhóm 1: hotline không liên hệ được, cho rằng "VietCredit lừa đảo". MHC seeding luôn theo kịch bản mẫu (mời liên hệ CSKH, hotline 1900 6515).',
  '• Nhóm 2: lỗi tính năng, không thao tác được, mới phát sinh. Chờ VietCredit duyệt.',
  '• Nhóm 3: tin tiêu cực về doanh nghiệp, tình hình kinh doanh, lãnh đạo. Chờ VietCredit duyệt.',
  '• Nhóm 4: hỏi bùng nợ được không, có bị đến nhà đòi nợ không, nên vay bên nào. MHC seeding luôn trong quota gói, không cần làm 100%, ưu tiên bài nhiều tương tác.',
  '• Khác: ngoài 4 nhóm (mạo danh, bên thứ ba dùng tên VietCredit, tuyển dụng...). Chờ VietCredit duyệt.',
  '**Nhóm 1, 4 vẫn chờ duyệt nếu:** khách hàng cụ thể tố VietCredit làm sai có chi tiết (ví dụ nhắc nợ gọi người thân, đe dọa); hoặc bài trên 100 tương tác; hoặc bên thứ ba dùng tên VietCredit quảng cáo dịch vụ.',
  '**Ghi mỗi case một dòng trên link:** ngày (dd/mm/yyyy), nền tảng, link, nhóm, tương tác, đánh giá case, phương án đề xuất, bên triển khai.',
  '**Phương án có thể đề xuất:** theo dõi thêm; liên hệ người đăng; seeding (số bình luận, hướng nội dung); report; gỡ bài (ghi chi phí, thời gian, bảo hành, tỷ lệ thành công).',
  'Xong thì nhắn nhóm chat báo đã cập nhật.',
], ['Nhóm 1, 4: "Đang triển khai"', 'Nhóm 2, 3, Khác: "Chờ VC duyệt"'], 'MHC đang ghi "đầu giờ chiều"; VietCredit đề nghị chốt 13:30.');
S('3. Review đề xuất', '13:30 đến 14:30', 'Marketing VC', [
  'Xem từng case "Chờ VC duyệt" và lướt kiểm tra các case Nhóm 1, 4 MHC đã làm. Với mỗi case, hỏi:',
  '• MHC phân nhóm đúng chưa? Có case tố VietCredit làm sai, case mạo danh, case lan tỏa cao bị xếp nhầm vào Nhóm 1, 4 không?',
  '• Thông tin trong bài đúng hay sai? Có cần tra cứu khách hàng không?',
  '• Phương án hợp lý chưa? (ví dụ bài quảng cáo rút tiền thẻ dùng tên VietCredit thì nên report, không seeding)',
  '• Nội dung bình luận đúng quy định ở bước 6 chưa?',
  '• Có chi phí ngoài gói không?',
  '**Case cần ý kiến chuyên môn thì gửi kèm link bài và đề xuất của MHC cho:**',
  '• lỗi tính năng, phí, lãi, hạn mức, chính sách: PD/PM;',
  '• nhắc nợ, đòi nợ, gọi người thân, đến nhà: Thu hồi nợ;',
  '• khách hàng cụ thể phàn nàn dịch vụ: CS;',
  '• mạo danh, vu khống, lộ thông tin khách hàng: Pháp chế;',
  '• doanh nghiệp, lãnh đạo: trình BĐH.',
], 'Case cần ý kiến đã gửi đơn vị liên quan', 'Case Marketing tự quyết được thì làm luôn bước 5, không cần chờ.');
S('4. Cho ý kiến chuyên môn', 'Trước 16:00', 'PD/PM, CS, Thu hồi nợ, Pháp chế (BĐH với Nhóm 3)', [
  'Trả lời Marketing: thông tin đúng hay sai; nên trả lời theo hướng nào; có cần liên hệ khách hàng không, ai liên hệ.',
  'Nếu khách hàng có vấn đề thật: CS hoặc Thu hồi nợ tra cứu và nhận việc liên hệ khách hàng.',
], 'Ý kiến gửi Marketing qua chat hoặc email', 'Quá 16:00 chưa trả lời: Marketing duyệt phương án an toàn (theo dõi thêm, hoặc bình luận mời liên hệ CSKH) và báo trưởng đơn vị đó.');
S('5. Duyệt phương án', 'Trước 16:30', 'Marketing VC', [
  'Ghi lên link cột "Ý kiến VC" một trong các lựa chọn: Đồng ý / Đồng ý, sửa nội dung thành... / Đổi phương án sang... / Không xử lý. Ghi tên người duyệt.',
  '**Chi phí ngoài gói (report, gỡ bài):** đến [hạn mức] mỗi lần thì Trưởng Marketing duyệt; vượt hạn mức thì trình BĐH.',
  'Nhắn nhóm chat báo MHC đã duyệt.',
], ['"Đang triển khai"', 'hoặc "Không xử lý"'], '');
S('6. Triển khai', 'Trong ngày, chậm nhất 9h00 hôm sau', ['MHC', 'CS, Thu hồi nợ nếu phải liên hệ khách hàng'], [
  'Làm đúng phương án đã duyệt hoặc kịch bản mẫu.',
  '**Seeding:** ghi vào sheet "Seeding" nội dung từng bình luận, link trực tiếp hoặc ảnh chụp bình luận đã đăng, thời điểm đăng.',
  '**Report, gỡ bài:** ghi ngày gửi, thời hạn dự kiến, kết quả khi có.',
  '**Liên hệ người đăng:** MHC cung cấp thông tin người đăng; CS hoặc Thu hồi nợ của VietCredit liên hệ. MHC không tự liên hệ nhân danh VietCredit.',
], '"Chờ nghiệm thu"', [
  '**Quy định bình luận:**',
  '• không đăng phần nhãn trong ngoặc vuông;',
  '• viết đúng tên "VietCredit";',
  '• không viết câu có thể hiểu là VietCredit làm sai khi chưa xác minh;',
  '• không lên lớp, dọa người đăng; không gợi ý trốn nợ; không hứa ưu đãi, giãn nợ chưa được xác nhận;',
  '• các bình luận trong một bài khác câu chữ, đăng giãn cách, từ tài khoản khác nhau.',
]);
S('7. Kiểm tra sau 48 giờ', '48 giờ sau khi đăng', 'MHC', [
  'Mở lại từng bình luận đã đăng, kiểm tra còn hiển thị không.',
  'Bị ẩn hoặc bị xóa thì đăng lại và cập nhật bằng chứng mới.',
], 'Cột "Kiểm tra sau 48 giờ": Còn / Bị ẩn / Bị xóa, đã đăng lại', 'Đăng lại không tính thêm vào quota.');
S('8. Nghiệm thu', ['Case đã duyệt: trong 1 ngày làm việc', 'Nhóm 1, 4: thứ Sáu hằng tuần'], 'Marketing VC', [
  'Mở bằng chứng trên link và kiểm tra:',
  '• seeding: đúng bài, đúng nội dung đã duyệt, đủ số lượng, đúng quy định bình luận, còn hiển thị sau 48 giờ;',
  '• report, gỡ bài: xong trong thời hạn MHC cam kết, có bằng chứng; bài gỡ không xuất hiện lại trong thời gian bảo hành;',
  '• theo dõi: MHC cập nhật diễn biến đúng hẹn.',
  'Nhóm 1, 4 kiểm tra ngẫu nhiên ít nhất 20% số case trong tuần. Nếu hơn 1/5 số case kiểm tra không đạt, tuần sau kiểm tra 100%.',
], ['Cột "Nghiệm thu":', '• Đạt hoặc Đạt có góp ý: tình trạng "Hoàn thành"', '• Cần bổ sung: ghi rõ cần làm gì, hạn, tình trạng "Cần bổ sung"'], 'Case "Cần bổ sung" chưa được tính vào khối lượng nghiệm thu cho đến khi đạt.');
S('9. Bổ sung, làm lại', ['Xác nhận: trong 2 giờ làm việc', 'Hoàn thành: trong 1 ngày làm việc hoặc theo hạn ghi trên link'], 'MHC', [
  'Làm theo yêu cầu ghi trên link: sửa hoặc xóa bình luận sai; thêm bình luận khi thảo luận tiêu cực tăng; đăng lại bình luận bị ẩn, xóa; đổi phương án (ví dụ sang report); theo dõi thêm.',
  'Cập nhật bằng chứng mới, nhắn nhóm chat. Marketing nghiệm thu lại như bước 8.',
], '"Chờ nghiệm thu"', 'Lỗi do MHC, hoặc bài đã gỡ đăng lại trong thời gian bảo hành: MHC làm lại, không tính thêm phí, quota.');
S('10. Chuyển xử lý tận gốc', 'Khi review hoặc nghiệm thu thấy vấn đề có thật', ['Marketing VC chuyển', 'Đơn vị phụ trách xử lý'], [
  '• Nhiều tin hotline không gọi được: CS kiểm tra năng lực tổng đài.',
  '• Lỗi tính năng lặp lại: PD/PM sửa, Marketing theo dõi đến khi xong.',
  '• Phản ánh nhắc nợ gọi người thân, đe dọa: Thu hồi nợ kiểm tra cuộc gọi, nhân sự hoặc đối tác nhắc nợ.',
  '• Bên thứ ba dùng tên VietCredit quảng cáo dịch vụ: Pháp chế xử lý; cân nhắc cảnh báo khách hàng trên kênh chính thức.',
], 'Đơn vị nhận việc, có hạn xử lý', '');

// ===== C =====
G('C. SỰ VỤ KHẨN (lãnh đạo, tình hình kinh doanh, bảo mật, lỗi sản phẩm hàng loạt, bài trên 100 tương tác hoặc lên báo điện tử)');
S('K1. Báo động', 'Trong 30 phút từ khi phát hiện, kể cả ngoài giờ và cuối tuần', 'MHC', [
  'Gọi điện cho đầu mối Marketing và nhắn nhóm chat.',
  'Gửi ngay form sự vụ gồm: tên sự vụ; tổng số bài và số bài trên từng nền tảng; nguồn đăng (link, thông tin liên hệ nếu có); nội dung sự vụ và bình luận nổi bật, tỷ lệ tiêu cực, trung tính, tích cực; tương tác hiện tại; đánh giá mức lan tỏa, ảnh hưởng; đề xuất phương án (kèm chi phí, thời gian, bảo hành nếu report, gỡ bài); những việc VietCredit nên tránh trong lúc chờ.',
], 'Form sự vụ đã gửi', '');
S('K2. Tiếp nhận', 'Trong 1 giờ', 'Marketing VC', [
  'Xác nhận với MHC đã nhận. Báo BĐH.',
  'Gọi nhóm xử lý: Trưởng Marketing, PD/PM sản phẩm liên quan, Trưởng CS, Thu hồi nợ (nếu liên quan), Pháp chế, thành viên BĐH phụ trách.',
], 'Nhóm xử lý đã họp', '');
S('K3. Chốt phương án', 'Trong 4 giờ', 'Nhóm xử lý; BĐH duyệt', [
  'Các đơn vị kiểm tra sự việc: đúng hay sai, bao nhiêu khách hàng bị ảnh hưởng, tổng đài đã nhận khiếu nại chưa.',
  'BĐH duyệt phương án, nội dung phát ngôn, chi phí. Marketing gửi MHC triển khai; CS dùng cùng nội dung để trả lời khách hàng gọi tổng đài.',
], 'Phương án và nội dung phát ngôn đã duyệt', [
  'Chỉ người được BĐH giao mới phát ngôn.',
  'Nhân viên không tự bình luận trên tài khoản cá nhân.',
  'Không xóa bình luận hàng loạt trên fanpage khi chưa có ý kiến nhóm xử lý.',
]);
S('K4. Theo dõi, đóng sự vụ', ['Đến khi ổn định', 'Báo cáo: trong 5 ngày làm việc sau khi đóng'], ['MHC', 'Marketing VC'], [
  'MHC cập nhật diễn biến theo mốc đã hẹn cho đến khi thảo luận lắng xuống.',
  'Marketing báo cáo BĐH: nguyên nhân, cách xử lý, chi phí, việc cần sửa để không lặp lại.',
], 'Báo cáo sự vụ gửi BĐH', '');

// ===== D =====
G('D. ĐỊNH KỲ');
S('D1. Báo cáo tuần', 'Thứ Hai', ['MHC gửi', 'Marketing VC nhận'], [
  'Số case theo sản phẩm, nền tảng, nhóm; tình trạng xử lý; số bình luận đã đăng và tỷ lệ còn hiển thị; case nổi bật; quota gói đã dùng.',
], 'Báo cáo tuần', '');
S('D2. Báo cáo, đối chiếu tháng', 'Ngày làm việc thứ 3 của tháng', ['MHC gửi', 'Marketing VC đối chiếu'], [
  'MHC: như báo cáo tuần, thêm xu hướng so với tháng trước, tỷ lệ tiêu cực trong thảo luận, kết quả report, gỡ bài, 5 vấn đề bị phàn nàn nhiều nhất, bảng khối lượng đã làm.',
  'Marketing: đối chiếu bảng khối lượng với kết quả nghiệm thu "Đạt" trên link trước khi xác nhận thanh toán. Đánh giá MHC: tỷ lệ đạt ngay lần đầu, tỷ lệ đúng hạn, số yêu cầu bổ sung, số lần phân nhóm sai.',
], 'Xác nhận khối lượng; nhận xét gửi MHC', '');
S('D3. Họp tháng, báo cáo BĐH', 'Tuần đầu tháng', ['Marketing VC chủ trì', 'PD/PM, CS, Thu hồi nợ'], [
  'Xem 5 vấn đề bị phàn nàn nhiều nhất, thống nhất việc khắc phục và thời hạn.',
  'Marketing gửi BĐH báo cáo tháng: số case, tỷ lệ tiêu cực, tỷ lệ duyệt đúng hạn, kết quả nghiệm thu và đánh giá MHC, chi phí ngoài gói, tiến độ khắc phục.',
], 'Báo cáo tháng gửi BĐH', '');

// ===== Tiêu đề + bảng =====
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 }, children: [
  new TextRun({ text: 'QUY TRÌNH XỬ LÝ THÔNG TIN TIÊU CỰC TRÊN MẠNG XÃ HỘI', font: FONT, size: 28, bold: true, color: RED })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 }, children: [
  new TextRun({ text: 'Phối hợp VietCredit và Màn Hình Cộng (MHC). Áp dụng: Thẻ tín dụng VC x DMX, TikTok BNPL', font: FONT, size: 20, color: GREY })] }));
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
