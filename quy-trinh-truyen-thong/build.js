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

// Title block
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 240, after: 80 }, children: [
  new TextRun({ text: 'QUY TRÌNH', font: FONT, size: 32, bold: true, color: RED })] }));
C.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [
  new TextRun({ text: 'QUẢN TRỊ RỦI RO TRUYỀN THÔNG VÀ XỬ LÝ THÔNG TIN TIÊU CỰC TRÊN MẠNG XÃ HỘI', font: FONT, size: 26, bold: true, color: RED })] }));
C.push(table([2600, 6426], ['Thông tin', 'Nội dung'], [
  ['Tình trạng', 'Dự thảo trình Ban Điều hành'],
  ['Ngày soạn', '07/10/2026'],
  ['Đơn vị chủ trì', 'Marketing VietCredit'],
  ['Đơn vị phối hợp', 'Quản lý sản phẩm (PD/PM), Chăm sóc khách hàng (CS), đối tác Màn Hình Cộng (MHC)'],
  ['Sản phẩm áp dụng giai đoạn 1', 'Thẻ tín dụng VietCredit x Điện Máy Xanh (VC x DMX); TikTok BNPL'],
], { boldFirst: true }));

// 1
C.push(H1('1. Mục đích'));
C.push(P('Quy trình này quy định cách VietCredit phát hiện, phân loại, xử lý và báo cáo các thông tin tiêu cực về công ty và sản phẩm trên mạng xã hội và báo điện tử. Mục tiêu cụ thể:'));
C.push(BL('Phát hiện sớm thông tin tiêu cực, trước khi bài viết lan rộng.'));
C.push(BL('Phân rõ case nào đối tác MHC được tự xử lý ngay theo kịch bản đã duyệt, case nào phải hỏi ý kiến VietCredit, case nào phải báo Ban Điều hành.'));
C.push(BL('Rút ngắn thời gian từ lúc phát hiện đến lúc có phương án, giảm việc trao đổi thủ công giữa hai bên.'));
C.push(BL('Chuyển các phàn nàn lặp lại về cho PD/PM và CS để sửa nguyên nhân gốc, thay vì chỉ xử lý từng bài đăng.'));

// 2
C.push(H1('2. Phạm vi áp dụng'));
C.push(H2('2.1. Sản phẩm'));
C.push(P('Giai đoạn 1 áp dụng đầy đủ quy trình (giám sát, xử lý chủ động, báo cáo) cho hai sản phẩm: Thẻ tín dụng VC x DMX và TikTok BNPL.'));
C.push(P('Các dự án khác đang có báo cáo Social listening vẫn được MHC theo dõi và tổng hợp như hiện tại. Riêng sự vụ Cấp độ 1 (mục 5.4) thuộc bất kỳ sản phẩm nào cũng xử lý theo luồng khẩn của quy trình này.'));
C.push(P('Sau giai đoạn 1, Marketing đề xuất Ban Điều hành bổ sung sản phẩm vào danh mục khi sản phẩm đáp ứng ít nhất một tiêu chí sau:'));
C.push(BL('VietCredit chịu toàn bộ rủi ro của dự án.'));
C.push(BL('Dự án có biên lợi nhuận (margin) tốt, cần bảo vệ doanh số.'));
C.push(BL('Dự án đang thử nghiệm, hoặc đang có vấn đề cần đo lường, giám sát, xử lý.'));
C.push(H2('2.2. Kênh theo dõi'));
C.push(P('Facebook (fanpage, nhóm, trang cá nhân công khai), TikTok, YouTube, Threads, báo điện tử và trang tin. Danh sách kênh có thể mở rộng theo năng lực hệ thống Social listening của MHC.'));

// 3
C.push(H1('3. Từ viết tắt và thuật ngữ'));
C.push(table([2300, 6726], ['Thuật ngữ', 'Giải thích'], [
  ['VC', 'VietCredit'],
  ['MHC', 'Màn Hình Cộng, đối tác cung cấp dịch vụ Social listening và quản trị tiêu cực chủ động (Crisis management) cho VietCredit'],
  ['BĐH', 'Ban Điều hành'],
  ['PD/PM', 'Đơn vị phát triển và quản lý sản phẩm phụ trách sản phẩm liên quan'],
  ['CS', 'Chăm sóc khách hàng'],
  ['Social listening', 'Theo dõi, cảnh báo, báo cáo các thảo luận về VietCredit, sản phẩm của VietCredit và đối thủ'],
  ['Seeding', 'Bình luận định hướng, giải thích thông tin, pha loãng thảo luận tiêu cực dưới bài viết công khai'],
  ['Report', 'Báo cáo bài viết với nền tảng để hạn chế tương tác, hạn chế tiếp cận người dùng khác'],
  ['Gỡ bài', 'Xử lý bằng biện pháp kỹ thuật để bài viết sai sự thật hoặc gây ảnh hưởng nghiêm trọng bị gỡ khỏi nền tảng'],
  ['Alert OTT', 'Cảnh báo khẩn qua ứng dụng nhắn tin (Zalo, Telegram...) gửi ngay khi phát hiện sự vụ lớn'],
  ['Tương tác', 'Tổng lượt thích, bình luận và chia sẻ của một bài viết'],
], { boldFirst: true }));

// 4
C.push(H1('4. Vai trò và trách nhiệm'));
C.push(table([1700, 7326], ['Đơn vị', 'Trách nhiệm'], [
  ['MHC', [
    '• Giám sát liên tục theo bộ từ khóa của từng sản phẩm.',
    '• Phân loại tin theo nhóm nội dung và chấm cấp độ sự vụ.',
    '• Tự xử lý các nhóm được ủy quyền (Nhóm 1, Nhóm 4) theo kịch bản đã duyệt.',
    '• Gửi file tổng hợp hằng ngày, gửi Alert OTT và Form đề xuất khi có sự vụ lớn.',
    '• Triển khai phương án đã chốt (seeding, report, gỡ bài), theo dõi sau xử lý và báo cáo định kỳ.',
    '• Đầu mối theo đề xuất của MHC: Hằng, Lam.']],
  ['Marketing VC', [
    '• Đầu mối duy nhất làm việc với MHC: tiếp nhận đề xuất, điều phối ý kiến nội bộ, gửi phương án đã chốt cho MHC.',
    '• Duyệt kịch bản seeding mẫu và cập nhật thư viện kịch bản.',
    '• Báo cáo BĐH sự vụ Cấp độ 1 và báo cáo tháng.',
    '• Quản lý quota gói dịch vụ và chi phí xử lý phát sinh ngoài gói.']],
  ['PD/PM sản phẩm', [
    '• Cho ý kiến về nội dung chuyên môn (tính năng, lỗi, phí, chính sách) trong thời hạn quy định.',
    '• Báo trước cho Marketing các thay đổi có thể gây phản ứng: bảo trì, lỗi đã biết, thay đổi phí, chính sách, chiến dịch.',
    '• Nhận danh sách vấn đề lặp lại từ báo cáo tháng và lên kế hoạch khắc phục.']],
  ['CS', [
    '• Xác minh thông tin khách hàng khi bài viết có đủ dữ liệu nhận diện.',
    '• Liên hệ trực tiếp khách hàng khi phương án là "Liên hệ".',
    '• Cung cấp câu trả lời chuẩn cho các thắc mắc thường gặp; báo Marketing khi tổng đài quá tải (liên quan Nhóm 1).']],
  ['Pháp chế', ['• Tham gia khi bài viết có dấu hiệu vu khống, xúc phạm, lộ thông tin khách hàng, hoặc khi cần gửi yêu cầu gỡ bài chính thức tới nền tảng, cơ quan báo chí.']],
  ['BĐH', ['• Phê duyệt quy trình và danh mục sản phẩm áp dụng.', '• Chỉ đạo xử lý sự vụ Cấp độ 1, duyệt phát ngôn chính thức và chi phí xử lý vượt hạn mức.']],
], { boldFirst: true }));

// 5
C.push(H1('5. Tiêu chí phân loại thông tin tiêu cực'));
C.push(P('Mỗi tin được MHC đánh giá theo hai trục: nội dung (đúng hay sai sự thật, mức ảnh hưởng tới uy tín, có xúc phạm, vu khống hay vi phạm chính sách nền tảng không) và tương tác (lượt thích, bình luận, chia sẻ, tốc độ lan truyền, mức ảnh hưởng của nguồn đăng). Kết quả là một nhóm nội dung (mục 5.1) và một cấp độ sự vụ (mục 5.4).'));

C.push(H2('5.1. Nhóm nội dung và quyền xử lý'));
C.push(table([1250, 3400, 1700, 2676], ['Nhóm', 'Nội dung', 'Quyền xử lý', 'Hướng xử lý'], [
  ['Nhóm 1', 'Gọi hotline mãi không được; cho rằng "VietCredit lừa đảo"', 'MHC tự xử lý, báo cáo sau', 'Seeding theo kịch bản mẫu. Hướng bình luận: VietCredit đảm bảo dịch vụ nhanh chóng, minh bạch, an toàn; khách hàng cần hỗ trợ xin liên hệ CSKH VietCredit (hotline 1900 6515).'],
  ['Nhóm 2', 'Lỗi tính năng, không thao tác được, mới phát sinh', 'Hỏi ý kiến VC', 'Marketing chuyển PD/PM xác nhận lỗi và hướng trả lời; CS liên hệ khách hàng nếu cần.'],
  ['Nhóm 3', 'Thông tin tiêu cực về doanh nghiệp, tình hình kinh doanh, lãnh đạo', 'Hỏi ý kiến VC', 'Marketing xin ý kiến PD/PM liên quan; báo BĐH nếu thuộc Cấp độ 1.'],
  ['Nhóm 4', 'Hỏi bùng nợ được không, có về nhà đòi nợ không, nên vay bên nào', 'MHC tự xử lý theo quota còn lại của gói', 'Dùng bộ kịch bản đang áp dụng. Không cần xử lý 100% số tin vì số lượng lớn; ưu tiên bài có tương tác cao.'],
  ['Nhóm 5', 'Nội dung khác chưa có kịch bản', 'Hỏi ý kiến VC', 'MHC đề xuất hướng xử lý; sau khi chốt, kịch bản được đưa vào thư viện để lần sau tự xử lý.'],
], { boldFirst: true }));
C.push(gap());
C.push(P('Quyền tự xử lý của Nhóm 1 và Nhóm 4 chỉ áp dụng khi sự vụ ở Cấp độ 3. Nếu bài viết thuộc hai nhóm này nhưng có lan tỏa cao (trên 100 tương tác), xuất hiện trên báo điện tử, hoặc nêu một trường hợp khách hàng cụ thể có bằng chứng, MHC không tự xử lý mà chuyển sang Luồng B hoặc Luồng C theo cấp độ.'));

C.push(H2('5.2. Mức độ lan tỏa'));
C.push(table([2200, 6826], ['Mức lan tỏa', 'Tiêu chí'], [
  ['Cao', ['• Bài viết trên mạng xã hội có tổng tương tác trên 100, hoặc', '• Bài viết xuất hiện trên trang tin, báo điện tử.']],
  ['Trung bình', 'Bài viết trên mạng xã hội có tổng tương tác từ 50 đến 100.'],
  ['Thấp', 'Bài viết trên mạng xã hội có tổng tương tác dưới 50.'],
], { boldFirst: true }));
C.push(gap());
C.push(P('Bài chia sẻ lại một bài gốc đã được cảnh báo thì cảnh báo ở mức thấp để tránh trùng lặp. Riêng bài chia sẻ lại link website có nội dung tiêu cực thì đánh giá theo mức lan tỏa như bài gốc.'));

C.push(H2('5.3. Ma trận mức độ nghiêm trọng'));
C.push(P('Kết hợp loại thông tin và mức lan tỏa theo bảng dưới (theo gợi ý của MHC).'));
C.push(table([4826, 1400, 1400, 1400], ['Loại thông tin tiêu cực', 'Lan tỏa cao', 'Lan tỏa trung bình', 'Lan tỏa thấp'], [
  ['Lãnh đạo: đời tư, quá trình công tác, bổ nhiệm; vi phạm pháp luật; phát ngôn bị xuyên tạc hoặc gây hiểu lầm', 'Nghiêm trọng', 'Nghiêm trọng', 'Nghiêm trọng'],
  ['Bảo mật, bí mật kinh doanh', 'Nghiêm trọng', 'Nghiêm trọng', 'Nghiêm trọng'],
  ['Hoạt động kinh doanh', 'Nghiêm trọng', 'Nghiêm trọng', 'Nghiêm trọng'],
  ['Công bố thông tin, truyền thông không đúng thực tế', 'Nghiêm trọng', 'Nghiêm trọng', 'Nghiêm trọng'],
  ['Sản phẩm, dịch vụ có lỗi mang tính hệ thống', 'Nghiêm trọng', 'Nghiêm trọng', 'Nghiêm trọng'],
  ['Nhân viên vi phạm pháp luật trên cương vị công tác hoặc lợi dụng danh nghĩa công ty (thông đồng làm giả hồ sơ, tham nhũng, lừa đảo, bị bắt, khởi tố)', 'Nghiêm trọng', 'Nghiêm trọng', 'Nghiêm trọng'],
  ['Tuyển dụng, đãi ngộ, điều kiện làm việc, khen thưởng, kỷ luật của bộ phận lớn nhân viên', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Thái độ của nhân viên', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Quy định, chính sách, thủ tục', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Đầu tư, xây dựng cơ bản, đấu thầu', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Cơ sở vật chất, công nghệ yếu kém, lạc hậu', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Nhân viên vi phạm luật giao thông, đánh nhau, phát ngôn nhạy cảm (chính trị, người khuyết tật, cộng đồng LGBT) có liên quan đến công việc', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Phàn nàn về sản phẩm, dịch vụ mang tính cục bộ (một cá nhân, một nhóm khách hàng, một hoặc vài tỉnh)', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Chia sẻ lại link website có nội dung tiêu cực', 'Nghiêm trọng', 'Vừa', 'Thấp'],
  ['Chia sẻ lại bài viết tiêu cực trên page, group, trang cá nhân', 'Thấp', 'Thấp', 'Thấp'],
  ['Đời tư vi phạm pháp luật của nhân viên, không liên quan công việc', 'Thấp', 'Thấp', 'Thấp'],
]));

C.push(H2('5.4. Cấp độ sự vụ và luồng xử lý'));
C.push(table([1500, 3326, 4200], ['Cấp độ', 'Áp dụng khi', 'Luồng xử lý'], [
  [['Cấp độ 1', 'Nghiêm trọng'], 'Kết quả ma trận là "Nghiêm trọng". Điển hình: thông tin về lãnh đạo, hoạt động và bí mật kinh doanh, truyền thông sai thực tế, lỗi sản phẩm mang tính hệ thống, phần lớn tin có lan tỏa cao.', 'Luồng C (khẩn): Alert OTT ngay, Marketing báo BĐH, họp nhóm xử lý.'],
  [['Cấp độ 2', 'Trung bình'], 'Kết quả ma trận là "Vừa". Điển hình: đãi ngộ nhân sự, chất lượng dịch vụ, quy định, chính sách, thủ tục, công nghệ, thái độ nhân viên.', 'Luồng B: MHC đề xuất, Marketing lấy ý kiến PD/PM/CS, chốt và xử lý trong ngày.'],
  [['Cấp độ 3', 'Thấp'], 'Kết quả ma trận là "Thấp".', 'Luồng A nếu thuộc Nhóm 1 hoặc Nhóm 4 (MHC tự xử lý). Các nhóm khác theo Luồng B, có thể chỉ theo dõi.'],
], { boldFirst: true }));
C.push(gap());
C.push(P('Lưu ý: hình minh họa "Cấp độ khủng hoảng" của MHC xếp "Nhân viên vi phạm pháp luật" vào Cấp độ 3, trong khi bảng chi tiết xếp trường hợp vi phạm trên cương vị công tác là Nghiêm trọng. Quy trình này áp dụng theo bảng chi tiết ở mục 5.3.', { italics: true, color: GREY }));

// 6
C.push(H1('6. Các hình thức xử lý'));
C.push(table([1500, 3000, 2300, 2226], ['Hình thức', 'Khi nào dùng', 'Người thực hiện', 'Lưu ý'], [
  ['Theo dõi', 'Tin lan tỏa thấp, chưa có thảo luận, hoặc nguồn đăng ít ảnh hưởng.', 'MHC', 'Chuyển sang hình thức khác nếu tương tác tăng nhanh.'],
  ['Liên hệ', 'Khách hàng có vấn đề thật, có thể nhận diện được (tên, số điện thoại, mã hợp đồng).', 'CS liên hệ khách hàng; MHC tìm thông tin nguồn đăng nếu có', 'Ưu tiên giải quyết vấn đề của khách hàng. Sau khi xử lý, có thể đề nghị khách hàng cập nhật bài viết.'],
  ['Seeding', 'Bài có tương tác nhưng chưa lan rộng; bài hỏi đáp có thể giải thích được.', 'MHC', 'Dùng kịch bản đã duyệt; giãn cách thời gian giữa các bình luận; chỉ áp dụng ở trang, nhóm, tài khoản công khai cho phép bình luận.'],
  ['Report hạn chế tương tác', 'Bài tương tác cao, cần hạn chế tiếp cận người dùng khác.', 'MHC', 'Theo dõi và báo cáo kết quả theo deadline đã tư vấn.'],
  ['Gỡ bài', 'Thông tin sai sự thật hoặc gây ảnh hưởng nghiêm trọng; video livestream đã kết thúc không seeding được.', 'MHC (kỹ thuật); Pháp chế (yêu cầu chính thức nếu cần)', 'Có chi phí ngoài gói. Form đề xuất phải nêu chi phí, thời gian, thời hạn bảo hành, tỷ lệ thành công. Ví dụ MHC đưa ra: 12.000.000 đồng, 1 đến 15 ngày, bảo hành 7 ngày, tỷ lệ 85%.'],
  ['Phản hồi chính thức', 'Sự vụ Cấp độ 1, hoặc khi cần VietCredit lên tiếng trên kênh chính thức.', 'Marketing soạn, BĐH duyệt', 'Không dùng seeding thay cho phản hồi chính thức trong sự vụ lớn.'],
], { boldFirst: true }));

C.push(H2('6.1. Nguyên tắc khi seeding và phản hồi'));
C.push(BL('Không phủ nhận sai sót đã được xác nhận là có thật. Khi lỗi có thật, nội dung là thừa nhận, nêu hướng xử lý và mời khách hàng liên hệ CS.'));
C.push(BL('Không công kích, chế giễu người đăng hoặc người bình luận.'));
C.push(BL('Không nêu thông tin cá nhân, thông tin khoản vay của khách hàng trên mạng xã hội.'));
C.push(BL('Không hứa hẹn ưu đãi, miễn giảm hoặc kết quả mà sản phẩm không có.'));
C.push(BL('Với Nhóm 4, không tư vấn hay ngầm khuyến khích việc trốn nợ. Nội dung nhấn vào nghĩa vụ trả nợ, hậu quả với lịch sử tín dụng và kênh hỗ trợ khi khách hàng gặp khó khăn.'));
C.push(BL('Trong thời gian chờ phương án chính thức, MHC tư vấn cho VietCredit những việc cần tránh (ví dụ: phản hồi vội trên fanpage, xóa bình luận hàng loạt).'));

// 7
C.push(H1('7. Quy trình xử lý'));
C.push(H2('7.1. Các bước'));
C.push(table([700, 2100, 1500, 4726], ['Bước', 'Công việc', 'Thực hiện', 'Mô tả'], [
  ['1', 'Giám sát', 'MHC', 'Theo dõi liên tục trên hệ thống Social listening theo bộ từ khóa từng sản phẩm. Ghi nhận mọi tin tiêu cực vào file tổng hợp (Mẫu 01).'],
  ['2', 'Phân loại', 'MHC', 'Xác định nhóm nội dung (mục 5.1), mức lan tỏa (5.2), cấp độ (5.4). Từ đó chọn Luồng A, B hoặc C.'],
  ['3A', 'Luồng A: tự xử lý', 'MHC', 'Nhóm 1 và Nhóm 4 ở Cấp độ 3: seeding ngay theo kịch bản đã duyệt, không cần trình duyệt từng bài. Ghi kết quả vào file tổng hợp.'],
  ['3B', 'Luồng B: hỏi ý kiến VC', 'MHC, Marketing, PD/PM, CS', [
    '• MHC đưa case vào file 10:30 kèm đánh giá và đề xuất phương án.',
    '• Marketing tiếp nhận, chuyển PD/PM hoặc CS của sản phẩm cho ý kiến.',
    '• Marketing chốt kịch bản và gửi lại MHC.',
    '• MHC triển khai trong ngày.']],
  ['3C', 'Luồng C: khẩn', 'MHC, Marketing, BĐH, Pháp chế', [
    '• MHC gửi Alert OTT ngay khi phát hiện, sau đó gửi Form đề xuất sự vụ lớn (Mẫu 02).',
    '• Marketing báo BĐH và triệu tập nhóm xử lý (PD/PM, CS, Pháp chế khi cần).',
    '• Nhóm xử lý chốt phương án, BĐH duyệt phát ngôn và chi phí.',
    '• MHC cập nhật diễn biến theo mốc thời gian đã thống nhất cho đến khi đóng sự vụ.']],
  ['4', 'Theo dõi sau xử lý', 'MHC', 'Theo dõi bài viết sau seeding, report, gỡ bài; theo dõi trong thời gian bảo hành; báo ngay nếu bài viết xuất hiện lại hoặc lan sang kênh khác.'],
  ['5', 'Đóng case', 'MHC, Marketing', 'Cập nhật trạng thái và kết quả. Kịch bản mới được duyệt đưa vào thư viện để lần sau xử lý theo Luồng A.'],
  ['6', 'Báo cáo và cải tiến', 'MHC, Marketing, PD/PM', 'Báo cáo tuần, tháng (mục 9). Vấn đề lặp lại được chuyển PD/PM, CS để sửa nguyên nhân gốc.'],
], { boldFirst: true }));

C.push(H2('7.2. Lịch phối hợp hằng ngày và thời hạn phản hồi'));
C.push(P('Thời hạn dưới đây là đề xuất, cần chốt với MHC theo năng lực của gói dịch vụ.'));
C.push(table([2600, 2000, 4426], ['Mốc', 'Thực hiện', 'Nội dung'], [
  ['10:30, thứ Hai đến thứ Sáu', 'MHC', 'Gửi file tổng hợp tin tiêu cực phát sinh, gồm: case đã tự xử lý (Luồng A) và case cần ý kiến VC (Luồng B) kèm đề xuất.'],
  ['Trước 14:00 cùng ngày', 'Marketing (sau khi lấy ý kiến PD/PM/CS)', 'Phản hồi phương án cho các case Luồng B. VC chỉ rà soát các case cần ý kiến, không phải xác nhận lại toàn bộ file.'],
  ['Trong ngày', 'MHC', 'Triển khai phương án đã chốt.'],
  ['Trong 30 phút từ khi phát hiện', 'MHC', 'Alert OTT với sự vụ Cấp độ 1, kể cả ngoài giờ hành chính và cuối tuần.'],
  ['Trong 1 giờ từ khi nhận alert', 'Marketing', 'Xác nhận đã nhận, báo BĐH, triệu tập nhóm xử lý.'],
  ['Trong 4 giờ từ khi nhận alert', 'Nhóm xử lý, BĐH', 'Chốt phương án xử lý Cấp độ 1.'],
], { boldFirst: true }));
C.push(gap());
C.push(P('Ghi chú: trong đề xuất của MHC, phần nguyên tắc ghi 10h và bảng timeline ghi 10:30. Quy trình chọn 10:30 để thống nhất.', { italics: true, color: GREY }));

// 8
C.push(H1('8. Phòng ngừa chủ động'));
C.push(P('Để giảm số tin phải xử lý sau khi đã đăng, các bên thực hiện thêm các việc sau:'));
C.push(NL('**Bộ từ khóa theo sản phẩm.** PD/PM cung cấp tên sản phẩm, tên gọi khách hàng hay dùng, tên đối tác (Điện Máy Xanh, TikTok Shop), các lỗi và thắc mắc thường gặp. MHC cấu hình vào hệ thống Social listening và rà lại mỗi quý.', 'num'));
C.push(NL('**Thư viện kịch bản duyệt sẵn.** Mỗi nhóm nội dung có bộ câu trả lời mẫu do Marketing duyệt, PD/PM và CS xác nhận nội dung. Thư viện được bổ sung sau mỗi case Luồng B.', 'num'));
C.push(NL('**Thông báo trước cho MHC.** Khi có bảo trì hệ thống, lỗi đã biết, thay đổi phí hoặc chính sách, chiến dịch lớn, PD/PM báo Marketing để Marketing gửi MHC kèm câu trả lời chuẩn, trước khi khách hàng đăng bài.', 'num'));
C.push(NL('**Theo dõi sớm khi ra mắt.** Trong 2 tuần đầu sau khi ra mắt sản phẩm, tính năng hoặc chiến dịch, MHC gửi báo cáo nhanh hằng ngày cho sản phẩm đó.', 'num'));
C.push(NL('**Phối hợp với CS.** Khi tin Nhóm 1 (hotline không gọi được) tăng đột biến, Marketing báo CS kiểm tra năng lực tổng đài. Khi CS nhận nhiều khiếu nại cùng một vấn đề, CS báo Marketing để MHC chuẩn bị kịch bản.', 'num'));
C.push(NL('**Sửa nguyên nhân gốc.** Báo cáo tháng liệt kê các vấn đề lặp lại nhiều nhất của từng sản phẩm. PD/PM phản hồi kế hoạch khắc phục trong cuộc họp tháng.', 'num'));

// 9
C.push(H1('9. Báo cáo và chỉ số theo dõi'));
C.push(table([1800, 2200, 5026], ['Báo cáo', 'Người nhận', 'Nội dung chính'], [
  ['Hằng ngày (file 10:30)', 'Marketing', 'Danh sách tin tiêu cực mới, nhóm, cấp độ, đề xuất, trạng thái.'],
  ['Hằng tuần', 'Marketing, PD/PM sản phẩm', 'Số tin tiêu cực theo sản phẩm, kênh, nhóm; tình trạng xử lý; case nổi bật; quota gói đã dùng.'],
  ['Hằng tháng', 'BĐH, PD/PM, CS', 'Xu hướng tin tiêu cực, hiệu quả xử lý, chi phí phát sinh ngoài gói, các vấn đề lặp lại cần sửa ở sản phẩm.'],
  ['Theo sự vụ', 'Marketing, BĐH', 'Báo cáo diễn biến và kết quả đối với sự vụ Cấp độ 1.'],
], { boldFirst: true }));
C.push(gap());
C.push(P('Các chỉ số đề xuất theo dõi:'));
C.push(BL('Số tin tiêu cực mới theo sản phẩm, kênh, nhóm nội dung và cấp độ.'));
C.push(BL('Tỷ lệ tiêu cực, trung tính, tích cực trong tổng thảo luận về sản phẩm.'));
C.push(BL('Thời gian từ khi đăng đến khi MHC phát hiện; từ khi phát hiện đến khi có phương án.'));
C.push(BL('Tỷ lệ case Luồng B được VC phản hồi đúng hạn.'));
C.push(BL('Tỷ lệ report, gỡ bài thành công và số bài xuất hiện lại trong thời gian bảo hành.'));
C.push(BL('Quota seeding đã dùng so với gói; chi phí xử lý ngoài gói.'));
C.push(BL('Số vấn đề lặp lại đã được PD/PM khắc phục.'));

// 10
C.push(H1('10. Kế hoạch triển khai'));
C.push(table([2200, 4626, 2200], ['Thời gian', 'Công việc', 'Thực hiện'], [
  ['Tuần 1', 'BĐH duyệt quy trình. Chốt đầu mối từng bên, nhóm OTT cảnh báo, thời hạn phản hồi với MHC.', 'Marketing, MHC'],
  ['Tuần 1 đến tuần 2', 'Chốt bộ từ khóa và thư viện kịch bản Nhóm 1 đến Nhóm 4 cho Thẻ tín dụng VC x DMX và TikTok BNPL.', 'PD/PM, CS, Marketing, MHC'],
  ['Tuần 3 đến tuần 6', 'Chạy thí điểm. Họp nhanh 15 phút mỗi tuần giữa Marketing và MHC để điều chỉnh phân loại.', 'Marketing, MHC'],
  ['Cuối tuần 6', 'Đánh giá thí điểm theo chỉ số ở mục 9, đề xuất BĐH sản phẩm mở rộng tiếp theo.', 'Marketing'],
], { boldFirst: true }));

// 11
C.push(H1('11. Nội dung cần Ban Điều hành cho ý kiến'));
C.push(NL('Phạm vi giai đoạn 1 gồm Thẻ tín dụng VC x DMX và TikTok BNPL, các dự án khác chỉ áp dụng luồng khẩn cho Cấp độ 1.', 'num2'));
C.push(NL('Ủy quyền cho MHC tự xử lý Nhóm 1 và Nhóm 4 ở Cấp độ 3 theo kịch bản đã duyệt.', 'num2'));
C.push(NL('Hạn mức chi phí xử lý ngoài gói (gỡ bài, report) mà Trưởng Marketing được duyệt; phần vượt hạn mức trình BĐH.', 'num2'));
C.push(NL('Người có thẩm quyền duyệt phát ngôn chính thức trong sự vụ Cấp độ 1.', 'num2'));
C.push(NL('Thời hạn phản hồi đề xuất tại mục 7.2.', 'num2'));

// Appendix
C.push(H1('Phụ lục. Mẫu biểu'));
C.push(H2('Mẫu 01. File tổng hợp tin tiêu cực hằng ngày'));
C.push(P('Mỗi dòng là một tin. Các cột:'));
C.push(table([2600, 6426], ['Cột', 'Cách ghi'], [
  ['Ngày phát hiện', 'dd/mm/yyyy, giờ phát hiện'],
  ['Sản phẩm', 'VC x DMX / TikTok BNPL / sản phẩm khác'],
  ['Kênh, link', 'Facebook, TikTok, YouTube, Threads, báo điện tử; đường link bài viết'],
  ['Nguồn đăng', 'Tên trang, nhóm hoặc tài khoản; thông tin liên hệ nếu có'],
  ['Tóm tắt nội dung', 'Một đến hai câu'],
  ['Tương tác', 'Lượt thích, bình luận, chia sẻ, lượt xem tại thời điểm ghi nhận'],
  ['Nhóm, cấp độ', 'Nhóm 1 đến 5; Cấp độ 1, 2, 3'],
  ['Luồng', 'A (đã tự xử lý), B (cần ý kiến VC), C (khẩn)'],
  ['Đề xuất của MHC', 'Hình thức xử lý và hướng nội dung'],
  ['Ý kiến VC', 'Marketing ghi phương án đã chốt'],
  ['Trạng thái, kết quả', 'Đang xử lý / Đã xử lý / Theo dõi; kết quả sau xử lý'],
], { boldFirst: true }));

C.push(H2('Mẫu 02. Form đề xuất sự vụ lớn'));
C.push(P('Dùng khi gửi Alert OTT hoặc cập nhật sự vụ.'));
C.push(table([2600, 6426], ['Mục', 'Nội dung cần điền'], [
  ['0. Tên sự vụ', 'Tên ngắn gọn, sản phẩm liên quan, cấp độ'],
  ['1. Tổng quan', 'Tổng số bài đăng; số bài trên từng nền tảng'],
  ['2.1. Nguồn đăng', 'Tên, link nguồn đăng; số điện thoại, email, kênh xã hội khác (nếu có)'],
  ['2.2. Nội dung', [
    '• Nội dung sự vụ',
    '• Tương tác hiện tại và thời điểm thảo luận gần nhất',
    '• Nội dung thảo luận, tỷ lệ tiêu cực, trung tính, tích cực và ý chính của từng nhóm',
    '• Các bài viết nổi bật (nếu có)']],
  ['3. Đánh giá', 'Mức độ lan tỏa, mức độ ảnh hưởng, lý do xếp cấp độ'],
  ['4. Đề xuất', [
    '• Phương án: theo dõi, seeding (số bình luận, giãn cách, hướng nội dung), report, gỡ bài',
    '• Với report, gỡ bài: chi phí, thời gian thực hiện, thời gian bảo hành, tỷ lệ thành công',
    '• Những việc VietCredit nên tránh trong lúc chờ phương án']],
], { boldFirst: true }));

C.push(H2('Mẫu 03. Báo cáo tuần, tháng'));
C.push(BL('Tổng quan: số tin tiêu cực mới, so sánh với kỳ trước, theo sản phẩm và kênh.'));
C.push(BL('Phân bổ theo nhóm nội dung và cấp độ.'));
C.push(BL('Tình trạng xử lý: số case Luồng A, B, C; số case đúng hạn, quá hạn.'));
C.push(BL('Hiệu quả: tỷ lệ report, gỡ bài thành công; tỷ lệ tiêu cực trong thảo luận trước và sau xử lý.'));
C.push(BL('Quota gói đã dùng, chi phí ngoài gói.'));
C.push(BL('Vấn đề lặp lại cần PD/PM, CS xử lý và đề xuất kịch bản mới.'));

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
    { reference: 'num', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
    { reference: 'num2', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 540, hanging: 300 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1200, left: 1440, right: 1440, header: 560 } } },
    headers: { default: header }, footers: { default: footer },
    children: C,
  }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(OUT, b); console.log('wrote', OUT); });
